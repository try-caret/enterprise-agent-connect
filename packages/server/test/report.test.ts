import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtempSync, symlinkSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import {
  errorReporter,
  noReport,
  reportingVendorRefusals,
  type AgentConnectError,
  type ReportTags,
} from "../src/report.ts";
import type { AuditEntry } from "../src/store.ts";

test("without a hook nothing is reported", () => {
  assert.equal(errorReporter(null), noReport);
  assert.equal(errorReporter(undefined), noReport);
});

test("an error reaches the hook once: kind, tags, class, code and status, never its message", () => {
  const seen: AgentConnectError[] = [];
  const report = errorReporter((e) => seen.push(e), { now: () => 1_700_000_000_000 });
  const error = Object.assign(new TypeError("token sk-secret for ann@example.com"), {
    code: "ECONNRESET",
  });
  report("teams_send_failed", error, { app: "teams", agent: "a1", bot: null, status: 502 });

  assert.equal(seen.length, 1);
  const [event] = seen;
  assert.equal(event.kind, "teams_send_failed");
  assert.deepEqual(event.tags, { app: "teams", agent: "a1", status: "502" });
  assert.equal(event.errorClass, "TypeError");
  assert.equal(event.code, "ECONNRESET");
  assert.equal(event.status, 502);
  assert.equal(event.timestamp, "2023-11-14T22:13:20.000Z");
  assert.ok(event.frames!.length > 0);
  assert.ok(event.frames!.every((f) => f.line > 0 && f.file));
  assert.match(event.frames![0].file, /report\.test\.ts$/);
  const body = JSON.stringify(event);
  assert.ok(!body.includes("sk-secret") && !body.includes("ann@example.com"), body);
});

test("a status on the error wins over the tag's", () => {
  const seen: AgentConnectError[] = [];
  errorReporter((e) => seen.push(e))(
    "event_failed",
    Object.assign(new Error("x"), { status: 503 }),
    { app: "jira", status: 400 },
  );
  assert.equal(seen[0].status, 503);
});

test("one event per kind and app a minute", () => {
  const seen: AgentConnectError[] = [];
  let now = 0;
  const report = errorReporter((e) => seen.push(e), { now: () => now });
  report("vendor_auth_failed", null, { app: "jira" });
  report("vendor_auth_failed", null, { app: "jira" });
  report("vendor_auth_failed", null, { app: "slack" });
  report("event_failed", null, { app: "jira" });
  now = 59_999;
  report("vendor_auth_failed", null, { app: "jira" });
  now = 60_000;
  report("vendor_auth_failed", null, { app: "jira" });
  assert.deepEqual(
    seen.map((e) => `${e.kind} ${e.tags.app}`),
    [
      "vendor_auth_failed jira",
      "vendor_auth_failed slack",
      "event_failed jira",
      "vendor_auth_failed jira",
    ],
  );
  // No error object: the kind and tags alone.
  assert.deepEqual(seen[0], {
    kind: "vendor_auth_failed",
    tags: { app: "jira" },
    errorClass: undefined,
    code: undefined,
    status: undefined,
    frames: undefined,
    timestamp: "1970-01-01T00:00:00.000Z",
  });
});

test("a hook that throws is logged, never thrown into the service", () => {
  const logged: unknown[] = [];
  const report = errorReporter(
    () => {
      throw new Error("sink down");
    },
    { log: { error: (...a: unknown[]) => logged.push(a) } },
  );
  assert.doesNotThrow(() => report("event_failed", new Error("x")));
  assert.equal(logged.length, 1);
});

test("an async hook that rejects is logged, never an unhandled rejection", async () => {
  const logged: unknown[] = [];
  const report = errorReporter(
    async () => {
      throw new Error("sink down");
    },
    { log: { error: (...a: unknown[]) => logged.push(a) } },
  );
  report("event_failed", new Error("x"));
  await new Promise((resolve) => setImmediate(resolve));
  assert.equal(logged.length, 1);
});

test("the minute's limit is per organization, subscription and bot too", () => {
  const seen: AgentConnectError[] = [];
  const report = errorReporter((e) => seen.push(e), { now: () => 0 });
  report("event_poll_failed", null, { app: "jira", org: "o1", subscription: "s1" });
  report("event_poll_failed", null, { app: "jira", org: "o1", subscription: "s1" });
  report("event_poll_failed", null, { app: "jira", org: "o1", subscription: "s2" });
  report("event_poll_failed", null, { app: "jira", org: "o2", subscription: "s3" });
  report("teams_send_failed", null, { app: "teams", org: "o1", bot: "b1" });
  report("teams_send_failed", null, { app: "teams", org: "o1", bot: "b1" });
  report("teams_send_failed", null, { app: "teams", org: "o1", bot: "b2" });
  assert.deepEqual(
    seen.map((e) => `${e.kind} ${e.tags.org} ${e.tags.subscription ?? e.tags.bot}`),
    [
      "event_poll_failed o1 s1",
      "event_poll_failed o1 s2",
      "event_poll_failed o2 s3",
      "teams_send_failed o1 b1",
      "teams_send_failed o1 b2",
    ],
  );
});

test("the server entry starts when launched through a symlink", () => {
  const link = join(mkdtempSync(join(tmpdir(), "eac-entry-")), "server.ts");
  symlinkSync(new URL("../src/server.ts", import.meta.url).pathname, link);
  const env = { ...process.env };
  delete env.DATABASE_URL;
  const run = spawnSync(process.execPath, ["--import", "tsx", link], { env, encoding: "utf8" });
  // It ran serveFromEnv, which stops at the first missing setting.
  assert.notEqual(run.status, 0);
  assert.match(run.stderr, /Missing required env var: DATABASE_URL/);
});

test("an audited call the app refused with 401 or 403 is reported after its row is written", async () => {
  const rows: AuditEntry[] = [];
  const reports: [string, ReportTags | undefined][] = [];
  const store = reportingVendorRefusals(
    { audit: async (entry: AuditEntry) => void rows.push(entry), other: 1 },
    (kind, _error, tags) => reports.push([kind, tags]),
  );
  const entry = (vendorStatus: number | null): AuditEntry => ({
    organizationId: "o1",
    agentId: "a1",
    agentName: "Triage",
    connectorId: "jira",
    tool: "jira_get_issue",
    target: "KAN-1 summary text",
    decision: "failed",
    reason: "Jira answered with a body",
    vendorStatus,
  });
  for (const status of [401, 403, 404, 500, null]) await store.audit(entry(status));
  assert.equal(store.other, 1);
  assert.equal(rows.length, 5);
  assert.deepEqual(reports, [
    [
      "vendor_auth_failed",
      { app: "jira", org: "o1", agent: "a1", tool: "jira_get_issue", status: 401 },
    ],
    [
      "vendor_auth_failed",
      { app: "jira", org: "o1", agent: "a1", tool: "jira_get_issue", status: 403 },
    ],
  ]);
});

test("a vendor refusal through the audit wrapper reaches the hook with no message", async () => {
  const seen: AgentConnectError[] = [];
  const store = reportingVendorRefusals(
    { audit: async (_entry: AuditEntry) => {} },
    errorReporter((e) => seen.push(e)),
  );
  await store.audit({
    organizationId: "o1",
    agentId: "a1",
    agentName: "Triage",
    connectorId: "jira",
    tool: "jira_get_issue",
    target: "KAN-1",
    decision: "failed",
    reason: "Jira said: invalid token for ann@example.com",
    vendorStatus: 401,
  });
  assert.equal(seen.length, 1);
  assert.equal(seen[0].status, 401);
  assert.deepEqual(seen[0].tags, {
    app: "jira",
    org: "o1",
    agent: "a1",
    tool: "jira_get_issue",
    status: "401",
  });
  assert.ok(!JSON.stringify(seen[0]).includes("ann@example.com"));
});
