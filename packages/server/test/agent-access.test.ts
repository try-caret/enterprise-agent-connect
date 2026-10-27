import assert from "node:assert/strict";
import { test } from "node:test";
import { connectorSchema, loadProviders } from "../../connector-kit/src/connector.ts";
import { createService } from "../src/service.ts";
import { ADMIN_TOKEN, harness, jira, microsoft, ORG } from "./helpers.ts";

type H = ReturnType<typeof harness>;

async function check(h: H, app: string, value: unknown) {
  const res = await h.app.fetch(
    new Request(`http://svc/agent-connect/admin/v1/orgs/${ORG}/connections/${app}/agent-access`, {
      method: "POST",
      headers: { authorization: `Bearer ${ADMIN_TOKEN}`, "content-type": "application/json" },
      body: JSON.stringify({ value, by: "ann@x.com" }),
    }),
  );
  return { status: res.status, body: (await res.json()) as Record<string, string> };
}

test("the address must fit the app's limit, the app must have one, and be on", async () => {
  const h = harness();
  const bad = await check(h, "outlook-mail", "agent@x.com/../../users/ceo@x.com");
  assert.equal(bad.status, 400);
  assert.match(bad.body.error, /not a valid mailbox/);

  const none = await check(h, "jira", "agent@x.com");
  assert.equal(none.status, 400);
  assert.match(none.body.error, /no agent address/);

  h.state.connections = h.state.connections.filter((c) => c.connectorId !== "outlook-mail");
  const off = await check(h, "outlook-mail", "agent@x.com");
  assert.equal(off.status, 422);
  assert.match(off.body.error, /Turn the app on first/);
});

test("agentAccess must name a text limit; an app has no add-an-agent steps of its own", () => {
  const wrong = connectorSchema.safeParse({
    ...microsoft,
    agentAccess: { ...microsoft.agentAccess, limit: "nope" },
  });
  assert.equal(wrong.success, false);
  const extra = { ...jira, addAgent: { intro: "x", steps: [] } };
  assert.equal(connectorSchema.safeParse(extra).success, false);
});

test("every Microsoft add-agent step is wired to its apps", () => {
  const microsoftProvider = loadProviders().find((p) => p.id === "microsoft")!;
  assert.ok(microsoftProvider.addAgent.some((s) => s.agentCheck));
  const step = { ...microsoftProvider.addAgent[0], apps: ["jira"] };
  assert.throws(
    () =>
      createService({
        apps: [{ connector: jira, vendor: {} as never, toolset: {} as never }],
        providers: [{ provider: { ...microsoftProvider, addAgent: [step] }, client: {} as never }],
        store: {} as never,
        key: { derive: () => ({}) } as never,
      }),
    /names jira, not one of its apps/,
  );
});

test("a check of an address no key holds is audited; hostile addresses are refused", async () => {
  const h = harness();
  await check(h, "outlook-mail", "outside@x.com");
  const row = h.state.audit.at(-1)!;
  assert.equal(row.tool, "admin:agent_access_check");
  assert.equal(row.target, "outside@x.com");
  assert.equal(row.agentName, "ann@x.com");
  for (const bad of [
    "a@x.co;Remove-Mailbox x",
    "$(iex 1)@x.com",
    "a'b@x.com",
    'a"@x.com',
    "a`b@x.com",
  ]) {
    const res = await check(h, "outlook-mail", bad);
    assert.equal(res.status, 400, bad);
  }
});
