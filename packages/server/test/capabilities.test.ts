import assert from "node:assert/strict";
import { test } from "node:test";
import { ADMIN_TOKEN, AGENT, harness, ORG, withBy } from "./helpers.ts";

type H = ReturnType<typeof harness>;

const toolNames = async (h: H, key: string) =>
  ((await h.rpc(key, "tools/list")).body.result.tools as { name: string }[])
    .map((t) => t.name)
    .sort();

// A refused call never reaches the app and leaves a denied audit row.
async function refused(h: H, key: string, tool: string, args: unknown, sent = h.vendor.requests) {
  const before = sent.length;
  const r = await h.callTool(key, tool, args);
  assert.ok(r.error || r.isError, `${tool} ${JSON.stringify(args)} was allowed`);
  assert.equal(h.state.audit.at(-1)!.tool, tool);
  assert.equal(h.state.audit.at(-1)!.decision, "denied");
  return { r, sentAfter: sent.length - before };
}

test("only tools of capabilities both ticked for the agent and allowed by the connection are listed", async () => {
  const h = harness();
  const key = h.addAgent(["read", "comment", "delete"]);
  h.state.connections[0].capabilities = ["read", "comment", "transition"];
  assert.deepEqual(await toolNames(h, key), [
    "jira_add_comment",
    "jira_get_issue",
    "jira_list_projects",
    "jira_search_issues",
  ]);
  // Delete is ticked for the agent but not allowed by the connection: refused and audited.
  const { sentAfter } = await refused(h, key, "jira_delete_issue", { issue_key: "KAN-1" });
  assert.equal(sentAfter, 0);
  // Move is allowed by the connection but not ticked for the agent.
  await refused(h, key, "jira_transition_issue", { issue_key: "KAN-1", transition: "Done" });
  assert.equal(h.vendor.requests.length, 0);
});

const admin = (h: H, method: string, path: string, body?: unknown) =>
  h.app.fetch(
    new Request(`http://svc/agent-connect/admin/v1/orgs/${ORG}${path}`, {
      method,
      headers: { authorization: `Bearer ${ADMIN_TOKEN}`, "content-type": "application/json" },
      body: body === undefined ? undefined : JSON.stringify(withBy(body)),
    }),
  );

test("the admin API validates capabilities and always keeps the required one", async () => {
  const h = harness();
  h.addAgent(["read"]);
  const grant = `/agents/${AGENT}/grants/jira`;
  assert.equal((await admin(h, "PUT", grant, { capabilities: ["read", "admin"] })).status, 400);
  assert.equal((await admin(h, "PUT", grant, { capabilities: "read" })).status, 400);
  assert.equal((await admin(h, "PUT", grant, { capabilities: ["delete", "comment"] })).status, 204);
  assert.deepEqual(h.state.grants.at(-1)!.capabilities, ["read", "comment", "delete"]);

  assert.equal(
    (await admin(h, "PATCH", "/connections/jira", { capabilities: ["nope"] })).status,
    400,
  );
  assert.equal((await admin(h, "PATCH", "/connections/jira", { capabilities: [] })).status, 204);
  assert.deepEqual(h.state.connections[0].capabilities, ["read"]);
});

test("a grant is never stored above the app's ceiling, nor for an app not connected", async () => {
  const h = harness();
  h.addAgent(["read"]);
  const grant = `/agents/${AGENT}/grants/jira`;
  h.state.connections[0].capabilities = ["read", "comment"];
  const over = await admin(h, "PUT", grant, { capabilities: ["read", "delete"] });
  assert.equal(over.status, 400);
  assert.match(((await over.json()) as { error: string }).error, /does not allow Delete issues/);
  assert.deepEqual(h.state.grants.at(-1)!.capabilities, ["read"]);
  h.state.connections = h.state.connections.filter((c) => c.connectorId !== "jira");
  const off = await admin(h, "PUT", grant, { capabilities: ["read"] });
  assert.equal(off.status, 400);
  assert.match(((await off.json()) as { error: string }).error, /Connect Jira before/);
});

test("lowering the connection's limit takes the capability off agents for good", async () => {
  const h = harness();
  const key = h.addAgent(["read", "delete"]);
  h.state.connections[0].capabilities = ["read", "delete"];
  assert.ok((await toolNames(h, key)).includes("jira_delete_issue"));
  const ceiling = (capabilities: string[]) =>
    admin(h, "PATCH", "/connections/jira", { capabilities });
  assert.equal((await ceiling(["read"])).status, 204);
  assert.equal((await ceiling(["read", "delete"])).status, 204);
  assert.ok(!(await toolNames(h, key)).includes("jira_delete_issue"));
  assert.deepEqual(h.state.grants[0].capabilities, ["read"]);
});

test("recent activity can be narrowed to one app", async () => {
  const h = harness();
  const key = h.addAgent(["read"]);
  await h.callTool(key, "jira_list_projects", {});
  await h.callTool(key, "mail_list_messages", {});
  const rows = (await (await admin(h, "GET", "/audit?app=jira")).json()) as {
    connectorId: string;
  }[];
  assert.deepEqual(
    rows.map((r) => r.connectorId),
    ["jira"],
  );
  assert.equal((await admin(h, "GET", "/audit?app=Not%20An%20App")).status, 400);
});
