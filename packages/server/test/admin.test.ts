import assert from "node:assert/strict";
import { test } from "node:test";
import { readFileSync } from "node:fs";
import { routes, type RouteSpec } from "../src/api.ts";
import { openapi } from "../src/openapi.ts";
import { ADMIN_TOKEN, AGENT, connectors, harness, ORG, withBy } from "./helpers.ts";

// `legacy`: the deprecated unversioned prefix, kept for one release.
function call(
  h: ReturnType<typeof harness>,
  method: string,
  path: string,
  opts: {
    token?: string | null;
    body?: unknown;
    headers?: Record<string, string>;
    legacy?: boolean;
  } = {},
) {
  const token = opts.token === undefined ? ADMIN_TOKEN : opts.token;
  return h.app.fetch(
    new Request(`http://svc/agent-connect/admin${opts.legacy ? "" : "/v1"}${path}`, {
      method,
      headers: {
        ...(token ? { authorization: `Bearer ${token}` } : {}),
        ...(opts.body ? { "content-type": "application/json" } : {}),
        ...opts.headers,
      },
      body: opts.body ? JSON.stringify(withBy(opts.body)) : undefined,
    }),
  );
}

test("the admin API needs the service token", async () => {
  const h = harness();
  assert.equal((await call(h, "GET", "/connectors", { token: null })).status, 401);
  assert.equal((await call(h, "GET", "/connectors", { token: "wrong" })).status, 401);
  assert.equal((await call(h, "GET", "/connectors", { token: ADMIN_TOKEN + "x" })).status, 401);
  assert.equal(
    (await call(h, "GET", "/connectors", { headers: { authorization: ADMIN_TOKEN }, token: null }))
      .status,
    401,
  );
  assert.equal((await call(h, "GET", "/connectors")).status, 200);
});

test("an agent key does not open the admin API", async () => {
  const h = harness();
  const key = h.addAgent(["read"]);
  assert.equal((await call(h, "GET", "/connectors", { token: key })).status, 401);
});

test("admin requests through the public load balancer are refused even with the token", async () => {
  const h = harness();
  const res = await call(h, "GET", "/connectors", { headers: { "x-forwarded-for": "1.2.3.4" } });
  assert.equal(res.status, 404);
});

test("a short admin token is refused at startup", () => {
  return import("../src/app.ts").then(({ createApp }) =>
    assert.throws(
      () => createApp({ service: {} as never, store: {} as never, adminToken: "short" }),
      /at least 32/,
    ),
  );
});

test("on the old prefix the first request for an organization creates it, with the admin app's id", async () => {
  const h = harness();
  const org = "0b0e5d1c-3f4a-4c2b-9d8e-7a6b5c4d3e2f";
  assert.ok(!h.state.organizations.has(org));
  assert.equal((await call(h, "GET", `/orgs/${org}/connections`, { legacy: true })).status, 200);
  assert.ok(h.state.organizations.has(org));
  assert.equal((await call(h, "GET", "/connectors", { legacy: true })).status, 200);
  assert.equal(h.state.organizations.size, 2);
});

test("v1 answers only for organizations it knows, made by id or by name", async () => {
  const h = harness();
  const org = "0b0e5d1c-3f4a-4c2b-9d8e-7a6b5c4d3e2f";
  const unknown = await call(h, "GET", `/orgs/${org}/connections`);
  assert.equal(unknown.status, 404);
  assert.ok(!h.state.organizations.has(org));
  assert.equal((await call(h, "GET", `/orgs/${org}`)).status, 404);

  const put = await call(h, "PUT", `/orgs/${org}`, { body: {} });
  assert.equal(put.status, 200);
  assert.deepEqual(((await put.json()) as any).name, null);
  assert.equal((await call(h, "GET", `/orgs/${org}/connections`)).status, 200);
  // A name given later is kept; a PUT without one leaves it.
  await call(h, "PUT", `/orgs/${org}`, { body: { name: "Acme" } });
  await call(h, "PUT", `/orgs/${org}`, { body: {} });
  assert.equal(((await (await call(h, "GET", `/orgs/${org}`)).json()) as any).name, "Acme");

  const made = await call(h, "POST", "/orgs", { body: { name: "Beta" } });
  assert.equal(made.status, 201);
  const beta = (await made.json()) as any;
  assert.equal(beta.name, "Beta");
  assert.equal((await call(h, "GET", `/orgs/${beta.id}/agents`)).status, 200);
  assert.equal((await call(h, "POST", "/orgs", { body: { name: " " } })).status, 400);

  const list = (await (await call(h, "GET", "/orgs")).json()) as any[];
  assert.deepEqual(list.map((o) => o.id).sort(), [ORG, org, beta.id].sort());
});

// Values for each kind of path parameter, so every route can be called by its declared path.
const sample: Record<string, string> = {
  org: ORG,
  agent: AGENT,
  app: connectors[0].id,
  provider: "microsoft",
  step: "a-step",
};

test("the old prefix and v1 give the same answer on every route that only reads", async () => {
  const h = harness();
  h.addAgent(["read"]);
  const reads = Object.values(routes as Record<string, RouteSpec>).filter(
    (r) => r.method === "GET" && !r.teams,
  );
  assert.ok(reads.length > 5);
  for (const route of reads) {
    const path = route.path.replace(/\{(\w+)\}/g, (_, p: string) => sample[p]);
    const [v1, legacy] = [await call(h, "GET", path), await call(h, "GET", path, { legacy: true })];
    assert.equal(v1.status, legacy.status, path);
    assert.ok(v1.status < 500, path);
    assert.deepEqual(await v1.text(), await legacy.text(), path);
  }
});

test("openapi.json is the one the routes make, and v1 serves it", async () => {
  const committed = JSON.parse(
    readFileSync(new URL("../openapi.json", import.meta.url), "utf8"),
  ) as { paths: Record<string, Record<string, unknown>> };
  assert.deepEqual(committed, openapi(), "run npm run openapi");
  // The published contract is the project's, not a deployment's.
  assert.doesNotMatch(JSON.stringify(committed), /arbium/i);
  const h = harness();
  const res = await call(h, "GET", "/openapi.json");
  assert.equal(res.status, 200);
  assert.deepEqual(await res.json(), committed);
  assert.equal((await call(h, "GET", "/openapi.json", { token: null })).status, 401);
  // Every declared path is in it, under v1.
  for (const route of Object.values(routes as Record<string, RouteSpec>))
    assert.ok(committed.paths[route.path]?.[route.method.toLowerCase()], route.path);
});

test("connections never include the credential", async () => {
  const h = harness();
  const rows = (await (await call(h, "GET", `/orgs/${ORG}/connections`)).json()) as any;
  assert.equal(rows.length, 5);
  for (const row of rows) assert.ok(!("credential" in row) && !("credentialKeyId" in row));
});

test("connecting verifies before saving and returns what agents can see", async () => {
  const h = harness();
  const res = await call(h, "PUT", `/orgs/${ORG}/connections/jira`, {
    body: {
      capabilities: ["read"],
      fields: { site: "x", client_id: "a", client_secret: "b" },
      connectedBy: "admin@x",
    },
  });
  assert.equal(res.status, 200);
  const body = (await res.json()) as any;
  assert.deepEqual(body.checks, [
    { label: "Connected Jira site", values: ["https://x.atlassian.net"] },
    { label: "Projects agents can see", values: ["KAN", "SECRET"] },
  ]);
  assert.ok(!JSON.stringify(body).includes('"b"'));
});

test("bad input is a 400, unknown routes 404", async () => {
  const h = harness();
  const bad = await call(h, "PUT", `/orgs/${ORG}/connections/jira`, {
    body: { capabilities: ["admin"], fields: {}, connectedBy: "a" },
  });
  assert.equal(bad.status, 400);
  assert.equal((await call(h, "GET", `/orgs/not-a-uuid/connections`)).status, 404);
  assert.equal(
    (
      await call(h, "PUT", `/orgs/${ORG}/connections/nope`, {
        body: { capabilities: ["read"], fields: {}, connectedBy: "a" },
      })
    ).status,
    400,
  );
});

test("creating an agent returns its key once and stores only the hash", async () => {
  const h = harness();
  const res = await call(h, "POST", `/orgs/${ORG}/agents`, {
    body: { name: "Triage", createdBy: "admin" },
  });
  assert.equal(res.status, 201);
  const { key, agent } = (await res.json()) as any;
  assert.match(key, /^arb_agent_/);
  assert.ok(!("keyHash" in agent));
  assert.ok(!JSON.stringify(h.state.agents).includes(key));
});

test("grants are validated against the app's levels and limits", async () => {
  const h = harness();
  h.addAgent(["read"]);
  const path = `/orgs/${ORG}/agents/${AGENT}/grants/jira`;
  assert.equal((await call(h, "PUT", path, { body: { capabilities: ["owner"] } })).status, 400);
  assert.equal(
    (await call(h, "PUT", path, { body: { capabilities: ["read"], limits: { mailbox: "a@b" } } }))
      .status,
    400,
  );
  assert.equal(
    (
      await call(h, "PUT", path, {
        body: { capabilities: ["read"], limits: { projects: ['KAN") OR ("'] } },
      })
    ).status,
    400,
  );
  assert.equal(
    (await call(h, "PUT", path, { body: { capabilities: ["read"], limits: { projects: "KAN" } } }))
      .status,
    400,
  );
  assert.equal(
    (
      await call(h, "PUT", path, {
        body: { capabilities: ["read", "comment"], limits: { projects: ["KAN", "KAN"] } },
      })
    ).status,
    204,
  );
  assert.deepEqual(h.state.grants.at(-1)!.limits, { projects: ["KAN"] });
});

test("reconnecting to a different account clears agents' grants; the same account keeps them", async () => {
  const h = harness();
  const connect = (cloud: string) =>
    call(h, "PUT", `/orgs/${ORG}/connections/jira`, {
      body: {
        capabilities: ["read"],
        fields: { site: "x", client_id: "a", client_secret: "b", cloud },
        connectedBy: "admin@x",
      },
    });
  assert.equal((await connect("c")).status, 200);
  assert.equal((await connect("another-site")).status, 200);
  assert.deepEqual(
    h.state.saved.map((r) => r.clearGrants),
    [false, true],
  );
});

test("after a key rotation a read-only check re-seals the credential with the current key", async () => {
  const h = harness({ rotated: true });
  const row = h.state.connections[0];
  const before = row.credentialKeyId;
  assert.notEqual(before, h.key.id);
  // Pages only read, so the re-seal cannot wait for a recorded check.
  const res = await call(h, "POST", `/orgs/${ORG}/connections/jira/check?record=false`);
  assert.equal(((await res.json()) as any).ok, true);
  assert.equal(h.state.checked.length, 0);
  assert.equal(row.credentialKeyId, h.key.id);
  assert.equal(
    h.key.open(
      { ciphertext: row.credential!, keyId: row.credentialKeyId! },
      `agent-connect:${ORG}:jira`,
    ).client_secret,
    "x",
  );
});

test("a re-seal never overwrites a credential saved by a reconnect during the check", async () => {
  const h = harness({ rotated: true });
  const row = h.state.connections[0];
  const fresh = Buffer.from("saved by a reconnect");
  const store = h.store as {
    getConnection: (o: string, a: string) => Promise<any>;
    markChecked: (id: string, ok: boolean) => Promise<void>;
  };
  const [read, mark] = [store.getConnection, store.markChecked];
  store.getConnection = async (o, a) => ({ ...(await read(o, a)) });
  store.markChecked = async (id, ok) => {
    await mark(id, ok);
    row.credential = fresh;
  };
  const res = await call(h, "POST", `/orgs/${ORG}/connections/jira/check`, {
    body: { by: "admin@x" },
  });
  assert.equal(((await res.json()) as any).ok, true);
  assert.equal(row.credential, fresh);
});

test("a check with record=false reports a failure without marking the connection", async () => {
  const h = harness();
  const real = h.vendor.call;
  h.vendor.call = async () => ({ status: 503, data: null });
  const quiet = await call(h, "POST", `/orgs/${ORG}/connections/jira/check?record=false`);
  assert.equal(((await quiet.json()) as any).ok, false);
  assert.equal(h.state.checked.length, 0);
  const recorded = await call(h, "POST", `/orgs/${ORG}/connections/jira/check`, {
    body: { by: "admin@x" },
  });
  assert.equal(((await recorded.json()) as any).ok, false);
  assert.deepEqual(
    h.state.checked.map((c) => c.ok),
    [false],
  );
  // The recorded failure cut the agents off, so it is audited under the admin who asked.
  const checks = () => h.state.audit.filter((r) => r.tool === "admin:check_app");
  assert.deepEqual(
    checks().map((r) => [r.agentName, r.connectorId]),
    [["admin@x", "jira"]],
  );
  assert.match(checks()[0].target!, /^failed: /);
  h.vendor.call = real;
  await call(h, "POST", `/orgs/${ORG}/connections/jira/check?record=false`);
  assert.equal(h.state.checked.length, 1);
  // Passing again reconnects it (audited); a recorded check needs the admin's name.
  await call(h, "POST", `/orgs/${ORG}/connections/jira/check`, { body: { by: "admin@x" } });
  assert.deepEqual(
    checks().map((r) => r.target),
    [checks()[0].target, "ok"],
  );
  await call(h, "POST", `/orgs/${ORG}/connections/jira/check`, { body: { by: "admin@x" } });
  assert.equal(checks().length, 2);
  assert.equal((await call(h, "POST", `/orgs/${ORG}/connections/jira/check`)).status, 400);
});

test("a key made or rotated whose audit row fails leaves the old key working", async () => {
  const h = harness();
  h.addAgent(["read"]);
  const before = h.state.agents[0].keyHash;
  h.state.auditDown = true;
  const rotated = await call(h, "POST", `/orgs/${ORG}/agents/${AGENT}/key`, { body: {} });
  assert.equal(rotated.status, 500);
  assert.equal(h.state.agents[0].keyHash, before);
  const created = await call(h, "POST", `/orgs/${ORG}/agents`, {
    body: { name: "New", createdBy: "ann@x" },
  });
  assert.equal(created.status, 500);
  assert.equal(h.state.agents.length, 1);
});

test("each admin change to apps, keys and access writes one audit row naming the admin", async () => {
  const h = harness();
  h.addAgent(["read"]);
  const org = `/orgs/${ORG}`;
  const by = { by: "ann@x" };
  const steps: [string, string, unknown, string, string][] = [
    [
      "PATCH",
      "/connections/jira",
      { capabilities: ["read", "comment"], ...by },
      "set_ceiling",
      "jira",
    ],
    [
      "PUT",
      `/agents/${AGENT}/grants/jira`,
      { capabilities: ["read", "comment"], limits: { projects: ["KAN"] }, ...by },
      "set_grant",
      "jira",
    ],
    ["DELETE", `/agents/${AGENT}/grants/jira`, by, "remove_grant", "jira"],
    ["PATCH", `/agents/${AGENT}`, { status: "disabled", ...by }, "key_disabled", "agents"],
    ["POST", `/agents/${AGENT}/key`, by, "rotate_key", "agents"],
    ["DELETE", `/agents/${AGENT}`, by, "delete_key", "agents"],
    ["POST", "/agents", { name: "New", createdBy: "ann@x" }, "create_key", "agents"],
    ["DELETE", "/connections/jira", by, "disconnect_app", "jira"],
    ["PUT", "/confirmed/mail-roles", by, "confirm_step", "outlook-mail"],
    ["DELETE", "/confirmed/mail-roles", by, "unconfirm_step", "outlook-mail"],
    ["PUT", "/confirmed/microsoft-app-registered", by, "confirm_step", "microsoft"],
  ];
  for (const [method, path, body, action, app] of steps) {
    const before = h.state.audit.length;
    const res = await call(h, method, `${org}${path}`, { body });
    assert.ok(res.status < 300, `${method} ${path}: ${res.status}`);
    const rows = h.state.audit.slice(before);
    assert.deepEqual(
      rows.map((r) => [r.tool, r.agentId, r.agentName, r.connectorId]),
      [[`admin:${action}`, null, "ann@x", app]],
      `${method} ${path}`,
    );
  }
  const grant = h.state.audit.find((r) => r.tool === "admin:set_grant")!;
  assert.match(grant.target!, /capabilities read, comment limits \{"projects":\["KAN"\]\}/);
  // A change without the admin's name is refused.
  const res = await h.app.fetch(
    new Request(`http://svc/agent-connect/admin${org}/connections/jira`, {
      method: "DELETE",
      headers: { authorization: `Bearer ${ADMIN_TOKEN}`, "content-type": "application/json" },
      body: "{}",
    }),
  );
  assert.equal(res.status, 400);
});
