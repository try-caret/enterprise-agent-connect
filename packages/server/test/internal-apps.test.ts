import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createServer } from "node:http";
import type { AddressInfo } from "node:net";
import { after, before, test } from "node:test";
import { AgentConnectError, createClient } from "../../client/src/index.ts";
import { createAgentConnect, type AgentConnect } from "../src/index.ts";
import { connectDatabase } from "../src/store.ts";

// Internal apps end to end: an organization's own app saved through the admin API, connected,
// granted and called by an agent against a local server, then replaced and deleted. Needs
// TEST_DATABASE_URL (see store.test.ts).
const url = process.env.TEST_DATABASE_URL;
const SCHEMA = "eac_internal_apps_test";
const TOKEN = "t".repeat(40);
const sql = url ? connectDatabase(url) : null;
const db = { skip: !sql && "TEST_DATABASE_URL not set" };
const BY = "ann@acme.example";

const orders = new Map([
  ["1001", { id: "1001", customer: "acme", status: "shipped" }],
  ["1002", { id: "1002", customer: "globex", status: "open" }],
]);
const app = createServer((req, res) => {
  const send = (status: number, data: unknown) => {
    res.writeHead(status, { "content-type": "application/json" });
    res.end(JSON.stringify(data));
  };
  if (req.headers["x-api-key"] !== "k1") return send(401, { error: "bad key" });
  if (req.url === "/api/health") return send(200, { name: "Orders" });
  const order = orders.get(/^\/api\/orders\/(\w+)$/.exec(req.url ?? "")?.[1] ?? "");
  return order ? send(200, order) : send(404, { error: "no such order" });
});

let service: AgentConnect;
let origin = "";
let A = "";
let B = "";
const fetchImpl = ((u: string, init: RequestInit) => service.fetch(new Request(u, init))) as never;
const client = createClient({ baseUrl: "http://svc", token: TOKEN, fetch: fetchImpl });

const example = JSON.parse(
  readFileSync(
    new URL("../../../examples/internal-app/connectors/orders/connector.json", import.meta.url),
    "utf8",
  ),
);
// The example orders app, reaching the local server.
const definition = (change: (d: any) => void = () => {}) => {
  const d = structuredClone(example);
  d.http.baseUrl = `${origin}/api`;
  change(d);
  return d;
};

before(async () => {
  if (!sql) return;
  await new Promise<void>((done) => app.listen(0, "127.0.0.1", done));
  origin = `http://127.0.0.1:${(app.address() as AddressInfo).port}`;
  await sql.unsafe(`drop schema if exists ${SCHEMA} cascade`);
  service = await createAgentConnect({
    database: sql,
    schema: SCHEMA,
    encryptionKey: Buffer.alloc(32, 3).toString("base64"),
    adminToken: TOKEN,
    log: { log() {}, error() {} },
  });
  A = (await client.organizations.create("Acme")).id;
  B = (await client.organizations.create("Beta")).id;
});
after(async () => {
  app.close();
  await service?.stop();
  await sql?.unsafe(`drop schema if exists ${SCHEMA} cascade`);
  await sql?.end();
});

let rpcId = 0;
async function mcp(key: string, method: string, params: unknown = {}) {
  const res = await service.fetch(
    new Request("http://svc/agent-connect/mcp", {
      method: "POST",
      headers: {
        "content-type": "application/json",
        accept: "application/json, text/event-stream",
        "mcp-protocol-version": "2025-06-18",
        authorization: `Bearer ${key}`,
      },
      body: JSON.stringify({ jsonrpc: "2.0", id: ++rpcId, method, params }),
    }),
  );
  const text = await res.text();
  const line = text.startsWith("{") ? text : text.split("\n").find((l) => l.startsWith("data: "));
  return JSON.parse(line!.replace(/^data: /, ""));
}
const toolNames = async (key: string) =>
  ((await mcp(key, "tools/list")).result.tools as { name: string }[]).map((t) => t.name).sort();

const refused = (status: number, pattern: RegExp) => (e: unknown) => {
  assert.ok(e instanceof AgentConnectError, String(e));
  assert.equal(e.status, status);
  assert.match(e.message, pattern);
  assert.doesNotMatch(e.message, /✖|→/);
  return true;
};

test("an internal app is saved, listed, read and offered to its organization only", db, async () => {
  const saved = await client.saveInternalApp(A, "orders", definition(), BY);
  assert.equal(saved.created, true);
  assert.equal(saved.disconnected, false);
  assert.equal(saved.connector.name, "Orders");
  assert.equal(saved.app.problem, null);
  assert.equal(saved.app.createdBy, BY);
  const list = await client.internalApps(A);
  assert.deepEqual(
    list.map((a) => a.id),
    ["orders"],
  );
  assert.deepEqual((await client.internalApp(A, "orders"))?.definition, definition());
  assert.ok((await client.orgConnectors(A)).some((c) => c.id === "orders"));
  assert.ok(!(await client.connectors()).some((c) => c.id === "orders"));

  // Organization B sees none of it and cannot use it.
  assert.deepEqual(await client.internalApps(B), []);
  assert.equal(await client.internalApp(B, "orders"), null);
  assert.ok(!(await client.orgConnectors(B)).some((c) => c.id === "orders"));
  await assert.rejects(
    client.connect(B, "orders", {
      capabilities: ["read"],
      fields: { api_key: "k1" },
      connectedBy: BY,
    }),
    refused(400, /Unknown app orders/),
  );

  const again = await client.saveInternalApp(A, "orders", definition(), BY);
  assert.equal(again.created, false);
  const audit = await client.audit(A, { app: "orders" });
  assert.deepEqual(
    audit.map((r) => r.tool),
    ["admin:replace_internal_app", "admin:create_internal_app"],
  );
});

test("a definition is checked as a connector.json file is, in plain words", db, async () => {
  const save = (id: string, d: unknown) => client.saveInternalApp(A, id, d as never, BY);
  await assert.rejects(
    save("orders", definition((d) => (d.name = ""))),
    refused(400, /^This definition does not load: .*\(at name\)/),
  );
  await assert.rejects(
    save("orders", definition((d) => delete d.http)),
    refused(400, /needs an http, soap, odata or sql block/),
  );
  await assert.rejects(
    save("orders", definition((d) => d.capabilities[0].tools.push("orders_missing"))),
    refused(400, /tools not implemented: orders_missing/),
  );
  await assert.rejects(save("other", definition()), refused(400, /saved as other/));
  // Nothing was changed by the refusals.
  assert.deepEqual((await client.internalApp(A, "orders"))?.definition, definition());
});

test("an id or a tool name already taken is refused", db, async () => {
  for (const id of ["jira", "salesforce", "http", "agents"])
    await assert.rejects(
      client.saveInternalApp(A, id, definition((d) => (d.id = id)), BY),
      refused(400, /one of this deployment's apps/),
    );
  // Another app of the same organization with the same tools.
  await assert.rejects(
    client.saveInternalApp(A, "orders-2", definition((d) => (d.id = "orders-2")), BY),
    refused(400, /already has a tool named orders_list, orders_get/),
  );
  // A built-in app's tool name.
  await assert.rejects(
    client.saveInternalApp(
      A,
      "orders",
      definition((d) => {
        d.http.tools[0].name = "jira_search_issues";
        d.capabilities[0].tools[0] = "jira_search_issues";
      }),
      BY,
    ),
    refused(400, /tool named jira_search_issues/),
  );
  // Another organization may use the same id and tools: its apps are its own.
  assert.equal((await client.saveInternalApp(B, "orders", definition(), BY)).created, true);
  await client.deleteInternalApp(B, "orders", BY);
});

let agentId = "";
let key = "";

test("an agent with a grant calls the internal app's tools, within its limits", db, async () => {
  const { checks } = await client.connect(A, "orders", {
    capabilities: ["read", "cancel"],
    fields: { api_key: "k1" },
    connectedBy: BY,
  });
  assert.deepEqual(checks[0].values, ["Orders"]);
  ({
    agent: { id: agentId },
    key,
  } = await client.createAgent(A, { name: "Clerk", createdBy: BY }));
  await client.setGrant(A, agentId, "orders", ["read"], { customers: ["acme"] }, BY);
  assert.deepEqual(await toolNames(key), ["orders_get", "orders_list"]);

  const call = async (id: string) =>
    (await mcp(key, "tools/call", { name: "orders_get", arguments: { id } })).result;
  const mine = await call("1001");
  assert.ok(!mine.isError);
  assert.equal(JSON.parse(mine.content[0].text).customer, "acme");
  assert.equal((await call("1002")).isError, true);
  // Not granted: denied before any tool runs.
  const cancel = await mcp(key, "tools/call", { name: "orders_cancel", arguments: { id: "1001" } });
  assert.ok(cancel.error || cancel.result.isError);
  const rows = await client.audit(A, { agent: agentId });
  assert.deepEqual(
    rows.map((r) => [r.tool, r.decision]),
    [
      ["orders_cancel", "denied"],
      ["orders_get", "denied"],
      ["orders_get", "allowed"],
    ],
  );
});

test("replacing a connected app cuts what the new definition no longer has", db, async () => {
  // Without the cancel capability: the connection stays, its ceiling loses cancel.
  const noCancel = definition((d) => {
    d.capabilities = d.capabilities.filter((c: { id: string }) => c.id !== "cancel");
    d.http.tools = d.http.tools.filter((t: { name: string }) => t.name !== "orders_cancel");
  });
  const cut = await client.saveInternalApp(A, "orders", noCancel, BY);
  assert.equal(cut.disconnected, false);
  assert.deepEqual(cut.grantsRemoved, []);
  const [conn] = (await client.connections(A)).filter((c) => c.connectorId === "orders");
  assert.deepEqual(conn.capabilities, ["read"]);
  assert.deepEqual(await toolNames(key), ["orders_get", "orders_list"]);

  // Without the customers limit, the agent's narrowed access no longer fits: it is removed,
  // not widened.
  const noLimit = definition((d) => {
    d.limits = [];
    delete d.http.tools[0].params.customer.limit;
    d.http.tools[1].answerLimits = [];
  });
  const widened = await client.saveInternalApp(A, "orders", noLimit, BY);
  assert.deepEqual(widened.grantsRemoved, [agentId]);
  assert.deepEqual(await toolNames(key), []);
  assert.deepEqual((await client.agents(A)).find((a) => a.id === agentId)?.grants, []);

  // Another address: the credential was entered for the old one, so the connection goes.
  await client.saveInternalApp(A, "orders", definition(), BY);
  await client.setGrant(A, agentId, "orders", ["read"], { customers: ["acme"] }, BY);
  const moved = await client.saveInternalApp(
    A,
    "orders",
    definition((d) => (d.http.baseUrl = "https://orders.example.com/api")),
    BY,
  );
  assert.equal(moved.disconnected, true);
  assert.ok(!(await client.connections(A)).some((c) => c.connectorId === "orders"));
  assert.deepEqual(await toolNames(key), []);
  const [row] = await client.audit(A, { app: "orders" }, 1);
  assert.equal(row.tool, "admin:replace_internal_app");
  assert.match(row.target ?? "", /^disconnected: /);
});

test("deleting an internal app removes its connection and every grant at once", db, async () => {
  await client.saveInternalApp(A, "orders", definition(), BY);
  await client.connect(A, "orders", {
    capabilities: ["read"],
    fields: { api_key: "k1" },
    connectedBy: BY,
  });
  await client.setGrant(A, agentId, "orders", ["read"], {}, BY);
  assert.deepEqual(await toolNames(key), ["orders_get", "orders_list"]);

  await client.deleteInternalApp(A, "orders", BY);
  assert.equal(await client.internalApp(A, "orders"), null);
  assert.ok(!(await client.connections(A)).some((c) => c.connectorId === "orders"));
  assert.deepEqual((await client.agents(A)).find((a) => a.id === agentId)?.grants, []);
  assert.deepEqual(await toolNames(key), []);
  assert.ok(!(await client.orgConnectors(A)).some((c) => c.id === "orders"));
  const [row] = await client.audit(A, { app: "orders" }, 1);
  assert.equal(row.tool, "admin:delete_internal_app");
  await assert.rejects(client.deleteInternalApp(A, "orders", BY), refused(404, /Not found/));
});

test("a draft from an OpenAPI spec saves as an internal app", db, async () => {
  const spec = JSON.parse(
    readFileSync(
      new URL("../../../scripts/fixtures/tracker.openapi.json", import.meta.url),
      "utf8",
    ),
  );
  const draft = await client.draftInternalApp({ spec, id: "tracker", tags: ["issues"] });
  assert.equal(draft.problem, null);
  assert.ok(draft.notes.length > 0);
  const saved = await client.saveInternalApp(A, "tracker", draft.definition, BY);
  assert.equal(saved.connector.id, "tracker");
  await assert.rejects(
    client.draftInternalApp({ spec: { swagger: "2.0" }, id: "tracker", tags: ["issues"] }),
    refused(400, /OpenAPI 3/),
  );
});
