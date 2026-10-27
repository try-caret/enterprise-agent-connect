import assert from "node:assert/strict";
import { test } from "node:test";
import { audited, json } from "../../connector-kit/src/common.ts";
import { AGENT, harness } from "./helpers.ts";

const toolNames = async (h: ReturnType<typeof harness>, key: string) =>
  ((await h.rpc(key, "tools/list")).body.result.tools as { name: string }[])
    .map((t) => t.name)
    .sort();

test("an agent without a valid key gets 401", async () => {
  const h = harness();
  assert.equal((await h.rpc(null, "tools/list")).status, 401);
  assert.equal((await h.rpc("arb_agent_" + "x".repeat(43), "tools/list")).status, 401);
  assert.equal((await h.rpc("not-a-key", "tools/list")).status, 401);
});

test("a disabled agent gets 401", async () => {
  const h = harness();
  const key = h.addAgent(["read"], {}, "disabled");
  assert.equal((await h.rpc(key, "tools/list")).status, 401);
});

test("browser requests are refused", async () => {
  const h = harness();
  const key = h.addAgent(["read"]);
  assert.equal(
    (await h.rpc(key, "tools/list", {}, { origin: "https://evil.example" })).status,
    403,
  );
});

test("read lists read tools only; ticking Comment adds commenting", async () => {
  const h = harness();
  const key = h.addAgent(["read"]);
  assert.deepEqual(await toolNames(h, key), [
    "jira_get_issue",
    "jira_list_projects",
    "jira_search_issues",
  ]);
  h.state.grants[0].capabilities = ["read", "comment"];
  assert.ok((await toolNames(h, key)).includes("jira_add_comment"));
});

test("the connection's ceiling caps the agent's grant", async () => {
  const h = harness();
  const key = h.addAgent(["read", "comment"]);
  h.state.connections[0].capabilities = ["read"];
  assert.ok(!(await toolNames(h, key)).includes("jira_add_comment"));
});

test("calling a tool outside the capabilities fails and nothing reaches Jira", async () => {
  const h = harness();
  const key = h.addAgent(["read"]);
  const r = await h.callTool(key, "jira_add_comment", { issue_key: "KAN-1", body: "hi" });
  assert.ok(r.error || r.isError);
  assert.equal(h.vendor.requests.length, 0);
  assert.deepEqual(h.state.audit.at(-1), {
    organizationId: h.state.agents[0].organizationId,
    agentId: h.state.agents[0].id,
    agentName: "Triage",
    connectorId: "jira",
    tool: "jira_add_comment",
    target: '{"issue_key":"KAN-1","body":"hi"}',
    decision: "denied",
    reason: "This agent is not allowed to do this",
    vendorStatus: null,
  });
});

test("arguments the tool's schema rejects are audited as denied, with what was sent", async () => {
  const h = harness();
  const key = h.addAgent(["read"]);
  const r = await h.callTool(key, "jira_get_issue", { issue_key: 7 });
  assert.ok(r.error || r.isError);
  assert.equal(h.vendor.requests.length, 0);
  const row = h.state.audit.at(-1)!;
  assert.equal(row.tool, "jira_get_issue");
  assert.equal(row.decision, "denied");
  assert.equal(row.target, '{"issue_key":7}');
  assert.match(row.reason!, /^Invalid arguments: issue_key: /);
  // A valid call is audited once, by the tool itself.
  const before = h.state.audit.length;
  await h.callTool(key, "jira_get_issue", { issue_key: "KAN-1" });
  assert.equal(h.state.audit.length, before + 1);
  assert.equal(h.state.audit.at(-1)!.decision, "allowed");
});

test("a call to an app the agent has no grant for is audited as denied too", async () => {
  const h = harness();
  const key = h.addAgent(["read"]);
  const r = await h.callTool(key, "mail_list_messages", {});
  assert.ok(r.error || r.isError);
  assert.equal(h.state.audit.at(-1)!.connectorId, "outlook-mail");
  assert.equal(h.state.audit.at(-1)!.decision, "denied");
  // Granted calls and unknown names add no extra row.
  await h.callTool(key, "jira_list_projects", {});
  await h.callTool(key, "no_such_tool", {});
  assert.equal(h.state.audit.length, 2);
});

test("an internal failure is logged, audited as failed and hidden from the agent", async (t) => {
  const h = harness();
  const key = h.addAgent(["read"]);
  t.mock.method(console, "error", () => {});
  h.vendor.call = async () => {
    throw new Error("Credential was sealed with key abc, which is not loaded");
  };
  const r = await h.callTool(key, "jira_list_projects", {});
  assert.equal(r.isError, true);
  assert.equal(r.text, "The call failed. Try again later.");
  assert.deepEqual(
    [h.state.audit.at(-1)!.decision, h.state.audit.at(-1)!.reason],
    ["failed", "Internal error"],
  );
});

test("a call that went through is not reported failed when its audit write fails", async (t) => {
  const errors: unknown[][] = [];
  t.mock.method(console, "error", (...args: unknown[]) => errors.push(args));
  let runs = 0;
  const r = await audited(
    {
      limits: {},
      config: {},
      call: async () => ({ status: 200, data: null }),
      audit: async () => {
        throw new Error("database down");
      },
    },
    "mail_send",
    "someone@example.com",
    async () => {
      runs++;
      return json({ sent: true });
    },
  );
  assert.equal(runs, 1);
  assert.equal(r.isError, undefined);
  assert.match((r.content[0] as { text: string }).text, /"sent": true/);
  assert.equal(errors.length, 1);
});

test("an oversized body or a long batch is refused before anything is audited", async () => {
  const h = harness();
  const key = h.addAgent(["read"]);
  const send = (body: string) =>
    h.app.fetch(
      new Request("http://svc/agent-connect/mcp", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          accept: "application/json, text/event-stream",
          authorization: `Bearer ${key}`,
        },
        body,
      }),
    );
  const call = (i: number) => ({
    jsonrpc: "2.0",
    id: i,
    method: "tools/call",
    params: { name: "jira_add_comment", arguments: { issue: "KAN-1", body: "x" } },
  });
  const big = await send(JSON.stringify({ ...call(1), pad: "x".repeat(4 * 1024 * 1024) }));
  assert.equal(big.status, 413);
  const batch = await send(JSON.stringify(Array.from({ length: 21 }, (_, i) => call(i))));
  assert.equal(batch.status, 400);
  assert.equal(h.state.audit.length, 0);
  // Within both, an ungranted call in a batch is still audited.
  await send(JSON.stringify([call(1)]));
  assert.equal(h.state.audit.length, 1);
});

test("no tools when the app is not connected or needs reconnecting", async () => {
  const h = harness();
  const key = h.addAgent(["read"]);
  h.state.connections[0].status = "needs_reconnect";
  assert.deepEqual(await toolNames(h, key), []);
});

test("a vendor error is returned and audited with its status", async () => {
  const h = harness();
  const key = h.addAgent(["read"]);
  const r = await h.callTool(key, "jira_get_issue", { issue_key: "NOPE-1" });
  assert.equal(r.isError, true);
  assert.match(r.text!, /404/);
  assert.equal(h.state.audit.at(-1)!.vendorStatus, 404);
  assert.equal(h.state.audit.at(-1)!.decision, "failed");
});

test("an agent's own MCP address takes only that agent's key", async () => {
  const h = harness();
  const mine = h.addAgent(["read"]);
  const other = h.addAgent(["read"]);
  const OTHER = "33333333-3333-3333-3333-333333333333";
  h.state.agents.at(-1)!.id = OTHER;
  h.state.grants.at(-1)!.agentId = OTHER;
  const list = (key: string, agentId: string) =>
    h.app.fetch(
      new Request(`http://svc/agent-connect/mcp/${agentId}`, {
        method: "POST",
        headers: {
          "content-type": "application/json",
          accept: "application/json, text/event-stream",
          "mcp-protocol-version": "2025-06-18",
          authorization: `Bearer ${key}`,
        },
        body: JSON.stringify({ jsonrpc: "2.0", id: 1, method: "tools/list", params: {} }),
      }),
    );
  assert.equal((await list(mine, AGENT)).status, 200);
  assert.equal((await list(other, AGENT)).status, 401);
  // The shared address still takes any agent's key.
  assert.equal((await h.rpc(other, "tools/list")).status, 200);
});
