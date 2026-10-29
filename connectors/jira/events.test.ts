import assert from "node:assert/strict";
import { createHmac } from "node:crypto";
import { test } from "node:test";
import {
  fill,
  fresh,
  locate,
  outsideLimits,
  pickFields,
  signed,
  valueAt,
} from "../../packages/server/src/events/deliver.ts";
import { jira } from "../../packages/server/test/helpers.ts";

// Jira's events as its connector.json declares them, on payloads shaped like Atlassian's
// webhook documentation (developer.atlassian.com/cloud/jira/platform/webhooks, read October 27,
// 2026). Never run against a real Jira site.

const events = jira.events!;
const hook = events.webhook!;
const SECRET = "c".repeat(64);
const now = Date.now();
const created = {
  timestamp: now,
  webhookEvent: "jira:issue_created",
  user: { accountId: "557058:human" },
  issue: {
    key: "KAN-7",
    fields: {
      summary: "Printer on fire",
      project: { key: "KAN" },
      issuetype: { name: "Bug" },
      status: { name: "To Do" },
      priority: { name: "High" },
      labels: ["hardware"],
      reporter: { accountId: "557058:human" },
      assignee: null,
      updated: "2026-10-27T10:15:30.123+0000",
    },
  },
};

function delivery(payload: unknown, secret = SECRET) {
  const body = Buffer.from(JSON.stringify(payload));
  const headers = new Headers({
    "X-Hub-Signature": `sha256=${createHmac("sha256", secret).update(body).digest("hex")}`,
    "X-Atlassian-Webhook-Identifier": "1234-abcd",
  });
  return { body, headers };
}

test("a Jira delivery is proven by X-Hub-Signature and named by its identifier", () => {
  const { body, headers } = delivery(created);
  assert.ok(signed(hook.verify, SECRET, body, headers));
  assert.ok(!signed(hook.verify, "d".repeat(64), body, headers));
  assert.equal(locate(hook.type, created, headers), "jira:issue_created");
  assert.equal(locate(hook.deliveryId!, created, headers), "1234-abcd");
  // Jira retries for over an hour; the window allows a day, dedupe keeps ids 30 days.
  assert.ok(fresh(hook.timestamp!, created, headers, now + 2 * 3600_000));
  assert.ok(!fresh(hook.timestamp!, created, headers, now + 2 * 86_400_000));
});

test("an agent sees only the kept fields, and is held to its projects", () => {
  const data = pickFields(events.fields, created);
  assert.equal(data.key, "KAN-7");
  assert.equal(data.project, "KAN");
  assert.equal(data.issue_type, "Bug");
  assert.deepEqual(data.labels, ["hardware"]);
  assert.equal(data.assignee, null);
  assert.equal(outsideLimits(events.limits, data, { projects: ["KAN"] }), null);
  assert.ok(outsideLimits(events.limits, data, { projects: ["HELP"] }));
});

test("the service account's own comments and changes are recognized", () => {
  const comment = {
    timestamp: now,
    webhookEvent: "comment_created",
    comment: { body: "Looking into it", author: { accountId: "svc-1" } },
    issue: created.issue,
  };
  const data = pickFields(events.fields, comment);
  assert.ok(events.self!.fields.some((f) => data[f] === "svc-1"));
  assert.equal(valueAt({ accountId: "svc-1" }, events.self!.id), "svc-1");
  assert.equal(events.self!.request.path, "/rest/api/3/myself");
});

test("a poll asks Jira for what changed since the cursor, in the subscription's JQL", () => {
  const poll = events.poll!;
  const request = fill(poll.request, {
    minutes: "6",
    cursor: "x",
    filter: poll.filter!.replace("{value}", "project = KAN"),
  }) as { body: { jql: string } };
  assert.equal(request.body.jql, "(project = KAN) AND updated >= -6m ORDER BY updated ASC");
  // Jira stamps a new issue's update just after its creation: created is what makes it new.
  const found = {
    key: "KAN-7",
    fields: { created: "2026-10-28T08:58:46.211+0200", updated: "2026-10-28T08:58:46.437+0200" },
  };
  const item = { [poll.wrap!]: found };
  assert.deepEqual(poll.types[0], { type: "jira:issue_created", since: "issue.fields.created" });
  assert.equal(valueAt(item, poll.types[0].since!), found.fields.created);
  assert.deepEqual(
    poll.deliveryId.map((p) => valueAt(item, p)),
    ["KAN-7", found.fields.updated],
  );
});
