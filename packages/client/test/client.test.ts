import assert from "node:assert/strict";
import { test } from "node:test";
import { ADMIN_TOKEN, connectors, harness, ORG } from "../../server/test/helpers.ts";
import {
  AgentConnectError,
  chosenSteps,
  CLAUDE_RUNNER,
  createClient,
  discovered,
  discoveredChoices,
  triggerForm,
  triggerInput,
  type TeamsTrigger,
} from "../src/index.ts";

const TOKEN = "t".repeat(40);

function recording(answer: (url: string, init: RequestInit) => Response) {
  const seen: { url: string; init: RequestInit }[] = [];
  const fetch = (async (url: string, init: RequestInit) => {
    seen.push({ url, init });
    return answer(url, init);
  }) as typeof globalThis.fetch;
  return { seen, client: createClient({ baseUrl: "http://svc:3000/", token: TOKEN, fetch }) };
}

test("the client sends the token to v1, maps refusals and fails loudly otherwise", async () => {
  let next = () => Response.json([]);
  const { seen, client } = recording(() => next());
  assert.deepEqual(await client.connectors(), []);
  assert.equal(seen[0].url, "http://svc:3000/agent-connect/admin/v1/connectors");
  assert.equal((seen[0].init.headers as Record<string, string>).authorization, `Bearer ${TOKEN}`);

  next = () => Response.json({ error: "The app answered 401" }, { status: 422 });
  await assert.rejects(
    client.disconnect("o", "an-app", "ann@x"),
    (e: unknown) => e instanceof AgentConnectError && e.status === 422 && /401/.test(e.message),
  );
  assert.equal(seen[1].init.body, JSON.stringify({ by: "ann@x" }));

  next = () => new Response(null, { status: 204 });
  assert.equal(await client.disconnect("o", "an-app", "ann@x"), undefined);

  next = () => Response.json({ error: "Internal error" }, { status: 500 });
  await assert.rejects(
    client.connectors(),
    (e: unknown) => !(e instanceof AgentConnectError) && /answered 500/.test(String(e)),
  );

  // A row that may not exist yet is null on 404.
  next = () => Response.json({ error: "Not found" }, { status: 404 });
  assert.equal(await client.teamsTrigger("o", "a"), null);
  assert.equal(await client.organizations.get("o"), null);
});

test("without its own fetch the client uses the global one of the moment", async (t) => {
  const real = globalThis.fetch;
  t.after(() => {
    globalThis.fetch = real;
  });
  const client = createClient({ baseUrl: "http://svc", token: TOKEN });
  const urls: string[] = [];
  globalThis.fetch = (async (url: string) => {
    urls.push(url);
    return Response.json([]);
  }) as typeof fetch;
  await client.audit("o", { agent: "a b" }, 30);
  await client.audit("o", { app: "x" });
  assert.deepEqual(
    urls.map((u) => u.split("/admin/v1")[1]),
    ["/orgs/o/audit?limit=30&agent=a+b", "/orgs/o/audit?limit=50&app=x"],
  );
});

test("live app states only read checks; one the service refuses is null", async (t) => {
  t.mock.method(console, "error", () => {});
  const { seen, client } = recording((url) =>
    url.includes("/check")
      ? Response.json({ error: "The vendor refused the sign-in" }, { status: 422 })
      : Response.json({ set: false }),
  );
  const apps = connectors.slice(0, 2);
  const connections = apps.map((c) => ({
    id: c.id,
    organizationId: "o",
    connectorId: c.id,
    providerId: null,
    capabilities: [],
    config: {},
    status: "connected" as const,
    accountLabel: null,
    connectedBy: null,
    connectedAt: "",
    checkedAt: "",
  }));
  const states = await client.appStates("o", apps, connections);
  const checks = await client.appChecks(
    "o",
    apps.map((c) => ({ ...c, provider: undefined })),
    connections,
  );
  for (const c of apps) {
    assert.equal(states[c.id].health, null);
    assert.equal(checks[c.id], null);
  }
  const checked = seen.filter((s) => s.url.includes("/check"));
  assert.ok(checked.length >= apps.length);
  for (const s of checked) assert.match(s.url, /[?]record=false$/);
});

test("the client's calls reach the service's routes and read its answers", async () => {
  const h = harness();
  const client = createClient({
    baseUrl: "http://svc",
    token: ADMIN_TOKEN,
    fetch: ((url: string, init: RequestInit) => h.app.fetch(new Request(url, init))) as never,
  });
  assert.deepEqual(await client.connectors(), h.service.connectors());
  assert.equal((await client.branding()).productName, "Enterprise Agent Connect");
  const org = "0b0e5d1c-3f4a-4c2b-9d8e-7a6b5c4d3e2f";
  await assert.rejects(client.agents(org), (e: unknown) => (e as AgentConnectError).status === 404);
  assert.equal((await client.organizations.ensure(org, "Acme")).name, "Acme");
  const { agent, key } = await client.createAgent(org, { name: "Triage", createdBy: "ann@x" });
  assert.ok(key);
  await client.setAgentStatus(org, agent.id, "disabled", "ann@x");
  assert.deepEqual(
    (await client.agents(org)).map((a) => [a.name, a.status]),
    [["Triage", "disabled"]],
  );
  const step = h.service
    .connectors()
    .flatMap((c) => [...c.setup, ...c.afterConnect])
    .find((s) => s.confirm)!.confirm!;
  await client.setConfirmed(org, step, true, "ann@x");
  assert.deepEqual(await client.confirmed(org), [step]);
  assert.ok((await client.connections(ORG)).length > 0);
  assert.ok((await client.audit(org, {})).length > 0);
  const made = await client.organizations.create("Beta");
  assert.deepEqual(await client.organizations.get(made.id), made);
});

test("the trigger form keeps another runner and its settings, asking Claude's fields only of Claude", () => {
  const trigger = {
    enabled: true,
    runner: "some-runner",
    settings: { url: "https://hooks.example/x" },
    managedAgentId: null,
    environmentId: null,
    vaultIds: [],
    budgetCents: 100,
    maxCentsPerDay: 500,
    maxRunsPerHour: 20,
    replyMode: "mentions",
    allowedTeams: [],
    allowedPeople: [],
    anyoneInOrganization: true,
  } as unknown as TeamsTrigger;
  const kept = triggerInput(triggerForm(trigger));
  assert.ok(kept.ok);
  assert.equal(kept.value.runner, "some-runner");
  assert.deepEqual(kept.value.settings, { url: "https://hooks.example/x" });
  assert.ok(!("managedAgentId" in kept.value));

  const fresh = triggerForm(null);
  assert.equal(fresh.runner, CLAUDE_RUNNER);
  const claude = triggerInput({ ...fresh, anyoneInOrganization: true });
  assert.deepEqual(claude, { ok: false, error: "Enter the Claude agent's ID" });
});

test("steps for a capability show only while it is ticked; discover labels are offered", () => {
  const netsuite = connectors.find((c) => c.id === "netsuite")!;
  const script = (ticked?: string[]) =>
    chosenSteps(netsuite, netsuite.afterConnect, {}, ticked).some((s) => s.capabilities);
  assert.equal(script(["read"]), false);
  assert.equal(script(["read", "saved_searches"]), true);
  assert.equal(script(), true);
  const checks = [
    { label: netsuite.discover!.label, values: ["1", "2"], labels: { "1": "Parent Co" } },
  ];
  assert.deepEqual(discovered(netsuite, checks), ["1", "2"]);
  assert.deepEqual(discoveredChoices(netsuite, checks), [
    { value: "1", label: "Parent Co" },
    { value: "2", label: "2" },
  ]);
});
