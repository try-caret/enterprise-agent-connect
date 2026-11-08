import assert from "node:assert/strict";
import { test } from "node:test";
import {
  errorReporter,
  noReport,
  reportingVendorRefusals,
  type ReportTags,
} from "../src/report.ts";
import type { AuditEntry } from "../src/store.ts";

function fakeSentry() {
  const sent: { url: string; headers: Record<string, string>; lines: any[] }[] = [];
  const fetch = (async (url: string, init: RequestInit) => {
    sent.push({
      url,
      headers: init.headers as Record<string, string>,
      lines: String(init.body)
        .split("\n")
        .map((l) => JSON.parse(l)),
    });
    return new Response(null, { status: 200 });
  }) as typeof globalThis.fetch;
  return { sent, fetch };
}
const settle = () => new Promise((r) => setImmediate(r));

test("without a DSN nothing is reported", () => {
  assert.equal(errorReporter(null), noReport);
  assert.equal(errorReporter({ dsn: "" }), noReport);
});

test("a DSN that is not Sentry's stops startup", () => {
  assert.throws(() => errorReporter({ dsn: "https://sentry.example.com/1" }), /SENTRY_DSN/);
  assert.throws(() => errorReporter({ dsn: "not a url" }));
});

test("an error goes as one envelope: fingerprint, tags, class and code, never its message", async () => {
  const s = fakeSentry();
  const report = errorReporter(
    {
      dsn: "https://pub123@sentry.example.com/prefix/42",
      environment: "staging",
      release: "1.2.3",
    },
    { fetch: s.fetch, now: () => 1_700_000_000_000 },
  );
  const error = Object.assign(new TypeError("token sk-secret for ann@example.com"), {
    code: "ECONNRESET",
  });
  report("teams_send_failed", error, { app: "teams", agent: "a1", bot: null, status: 502 });
  await settle();

  assert.equal(s.sent.length, 1);
  const [{ url, headers, lines }] = s.sent;
  assert.equal(url, "https://sentry.example.com/prefix/api/42/envelope/");
  assert.match(headers["x-sentry-auth"], /sentry_version=7, sentry_key=pub123,/);
  const [head, item, event] = lines;
  assert.equal(head.event_id, event.event_id);
  assert.deepEqual(item, { type: "event" });
  assert.deepEqual(event.fingerprint, ["agent-connect", "teams_send_failed"]);
  assert.deepEqual(event.tags, {
    kind: "teams_send_failed",
    app: "teams",
    agent: "a1",
    status: "502",
  });
  assert.equal(event.environment, "staging");
  assert.equal(event.release, "1.2.3");
  assert.equal(event.timestamp, 1_700_000_000);
  const [exception] = event.exception.values;
  assert.equal(exception.type, "teams_send_failed");
  assert.equal(exception.value, "TypeError ECONNRESET 502");
  assert.ok(exception.stacktrace.frames.length > 0);
  assert.ok(exception.stacktrace.frames.every((f: { lineno: number }) => f.lineno > 0));
  const body = JSON.stringify(lines);
  assert.ok(!body.includes("sk-secret") && !body.includes("ann@example.com"), body);
});

test("one event per kind and app a minute", async () => {
  const s = fakeSentry();
  let now = 0;
  const report = errorReporter(
    { dsn: "https://k@sentry.example.com/1" },
    { fetch: s.fetch, now: () => now },
  );
  report("vendor_auth_failed", null, { app: "jira" });
  report("vendor_auth_failed", null, { app: "jira" });
  report("vendor_auth_failed", null, { app: "slack" });
  report("event_failed", null, { app: "jira" });
  now = 59_999;
  report("vendor_auth_failed", null, { app: "jira" });
  now = 60_000;
  report("vendor_auth_failed", null, { app: "jira" });
  await settle();
  assert.deepEqual(
    s.sent.map((e) => `${e.lines[2].tags.kind} ${e.lines[2].tags.app}`),
    [
      "vendor_auth_failed jira",
      "vendor_auth_failed slack",
      "event_failed jira",
      "vendor_auth_failed jira",
    ],
  );
  // No error object: the kind alone, no stack.
  assert.deepEqual(s.sent[0].lines[2].exception.values[0], {
    type: "vendor_auth_failed",
    value: "",
  });
});

test("a report that cannot be sent is logged, never thrown", async () => {
  const logged: unknown[] = [];
  const report = errorReporter(
    { dsn: "https://k@sentry.example.com/1" },
    {
      fetch: (async () => {
        throw new Error("offline");
      }) as typeof fetch,
      log: { error: (...a: unknown[]) => logged.push(a) },
    },
  );
  report("event_failed", new Error("x"));
  await settle();
  assert.equal(logged.length, 1);
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
