import assert from "node:assert/strict";
import { test } from "node:test";
import type { VendorRequest, VendorResponse } from "../../packages/connector-kit/src/types.ts";
import { harness } from "../../packages/server/test/helpers.ts";

// The Jira tools against an in-memory Jira (the server's test harness). Projects KAN (a
// software project) and HELP (a service project), issues KAN-1 and HELP-1.

const ACCOUNT = (id: string, type: string, email: string | null) => ({
  accountId: id,
  accountType: type,
  displayName: `Person ${id}`,
  ...(email ? { emailAddress: email } : {}),
});
const AGENT = ACCOUNT("agent-1", "atlassian", "agent@acme.example");
const CUSTOMER = ACCOUNT("c-1", "customer", "customer@example.com");

const change = (author: object, day: number, from: string, to: string, field = "labels") => ({
  id: String(day),
  author,
  created: `2026-10-${String(day).padStart(2, "0")}T00:00:00.000+0000`,
  items: [{ field, fieldtype: "jira", fromString: from, toString: to }],
});

// HELP-1's history, oldest first: a customer adds urgent, an agent removes it and adds it back,
// then the customer adds vip. legacy was set when the issue was created, so it is not there.
const HISTORY = [
  change(CUSTOMER, 1, "", "urgent"),
  change(AGENT, 2, "urgent", ""),
  change(AGENT, 2, "Open", "Waiting", "status"),
  change(AGENT, 3, "", "urgent"),
  change(CUSTOMER, 4, "urgent", "urgent vip"),
];

type Options = {
  /** jsdPublic in every comment POST's answer; by default false only when flagged internal. */
  jsdPublic?: boolean;
  /** HELP-1's changelog, oldest first, or a failing call. */
  changelog?: object[] | "fail";
  comments?: "fail";
};

function jira(h: ReturnType<typeof harness>, options: Options = {}) {
  const sent: VendorRequest[] = [];
  h.vendor.call = async (_conn, req): Promise<VendorResponse> => {
    sent.push(req);
    const [path, query] = req.path.split("?");
    const params = new URLSearchParams(query);
    const issue = /^\/rest\/api\/3\/issue\/((KAN|HELP)-1)$/.exec(path);
    if (issue)
      return {
        status: 200,
        data: {
          key: issue[1],
          fields: {
            summary: "Printer",
            project: {
              key: issue[2],
              projectTypeKey: issue[2] === "HELP" ? "service_desk" : "software",
            },
            reporter: ACCOUNT(`r-${issue[2]}`, "customer", "reporter@example.com"),
            labels: issue[2] === "HELP" ? ["urgent", "vip", "legacy"] : [],
            description: null,
          },
        },
      };
    if (path === "/rest/api/3/issue/HELP-1/changelog") {
      const changelog = options.changelog ?? HISTORY;
      if (changelog === "fail") return { status: 500, data: { errorMessages: ["Down"] } };
      const startAt = Number(params.get("startAt"));
      const count = Math.min(Number(params.get("maxResults")), 100);
      const values = changelog.slice(startAt, startAt + count);
      return {
        status: 200,
        data: {
          startAt,
          total: changelog.length,
          values,
          isLast: startAt + values.length >= changelog.length,
        },
      };
    }
    if (/^\/rest\/api\/3\/issue\/(KAN|HELP)-1\/comment$/.test(path) && req.method === "GET") {
      if (options.comments === "fail") return { status: 503, data: { errorMessages: ["Busy"] } };
      // Newest first, as asked with orderBy=-created; 25 of them, more than are returned.
      const comments = Array.from({ length: 25 }, (_, i) => 25 - i).map((n) => ({
        id: String(n),
        created: `2026-10-${String(n).padStart(2, "0")}T00:00:00.000+0000`,
        author: n % 2 ? AGENT : CUSTOMER,
        body: {
          type: "doc",
          version: 1,
          content: [{ type: "paragraph", content: [{ type: "text", text: `c${n}` }] }],
        },
        ...(n === 25 ? {} : { jsdPublic: n % 2 === 0 }),
      }));
      return { status: 200, data: { comments: comments.slice(0, 20), total: 25 } };
    }
    if (/^\/rest\/api\/3\/issue\/(KAN|HELP)-1\/comment$/.test(path)) {
      // Jira answers jsdPublic false for a comment flagged internal (and, live, for any comment
      // from an account without a Service Management seat).
      const flagged = JSON.stringify(req.body ?? {}).includes("sd.public.comment");
      return {
        status: 201,
        data: {
          id: flagged ? "101" : "100",
          created: "2026-10-25T10:00:00.000+0000",
          jsdPublic: options.jsdPublic ?? !flagged,
        },
      };
    }
    return { status: 404, data: { errorMessages: ["Not found"] } };
  };
  return sent;
}

const filler = (n: number) =>
  Array.from({ length: n }, (_, i) => change(AGENT, 5, `s${i}`, `s${i + 1}`, "status"));
const getIssue = async (h: ReturnType<typeof harness>, key: string, issue_key: string) =>
  JSON.parse((await h.callTool(key, "jira_get_issue", { issue_key })).text!);

test("jira_get_issue says who the reporter and comment authors are, without emails", async () => {
  const h = harness();
  const key = h.addAgent(["read"]);
  const sent = jira(h);
  const r = await h.callTool(key, "jira_get_issue", { issue_key: "HELP-1" });
  assert.equal(r.isError, false);
  assert.ok(!r.text!.includes("@example.com") && !r.text!.includes("email"));
  const issue = JSON.parse(r.text!);
  assert.equal(issue.reporter, "Person r-HELP");
  assert.equal(issue.reporter_account_id, "r-HELP");
  assert.equal(issue.reporter_account_type, "customer");
  // The 20 most recent, oldest first.
  assert.equal(issue.comments.length, 20);
  assert.deepEqual(
    issue.comments.map((c: { body: string }) => c.body),
    Array.from({ length: 20 }, (_, i) => `c${i + 6}`),
  );
  const [first, second] = issue.comments;
  assert.deepEqual(
    [first, second].map((c) => [c.author_account_id, c.author_account_type, c.internal]),
    [
      ["c-1", "customer", false],
      ["agent-1", "atlassian", true],
    ],
  );
  // A comment without jsdPublic says nothing either way.
  assert.equal(issue.comments.at(-1).internal, null);
  const comments = sent.find((s) => s.path.includes("/comment"))!;
  assert.equal(comments.path, "/rest/api/3/issue/HELP-1/comment?orderBy=-created&maxResults=20");
  assert.ok(!sent[0].path.includes("comment"), "the issue is read without its comment field");
  // The project type comes with the issue: no project call.
  assert.ok(!sent.some((s) => s.path.startsWith("/rest/api/3/project/")));
});

test("jira_get_issue leaves internal null outside a service project", async () => {
  const h = harness();
  const key = h.addAgent(["read"]);
  const sent = jira(h);
  const issue = JSON.parse((await h.callTool(key, "jira_get_issue", { issue_key: "KAN-1" })).text!);
  assert.ok(issue.comments.length > 0);
  assert.ok(issue.comments.every((c: { internal: unknown }) => c.internal === null));
  // No labels, no changelog read.
  assert.deepEqual(issue.labels_added_by, []);
  assert.ok(!sent.some((s) => s.path.includes("changelog")));
});

test("jira_get_issue fails clearly when the comments cannot be read", async () => {
  const h = harness();
  const key = h.addAgent(["read"]);
  jira(h, { comments: "fail" });
  const r = await h.callTool(key, "jira_get_issue", { issue_key: "HELP-1" });
  assert.equal(r.isError, true);
  assert.match(r.text!, /Jira answered 503: Busy/);
});

test("jira_get_issue says who last added each label", async () => {
  const h = harness();
  const key = h.addAgent(["read"]);
  const sent = jira(h);
  const issue = await getIssue(h, key, "HELP-1");
  assert.deepEqual(issue.labels_added_by, [
    // Removed and added back: the latest add, by an agent, not the customer's first one.
    {
      label: "urgent",
      author: "Person agent-1",
      author_account_id: "agent-1",
      author_account_type: "atlassian",
      added: "2026-10-03T00:00:00.000+0000",
    },
    {
      label: "vip",
      author: "Person c-1",
      author_account_id: "c-1",
      author_account_type: "customer",
      added: "2026-10-04T00:00:00.000+0000",
    },
    {
      label: "legacy",
      author: null,
      author_account_id: null,
      author_account_type: null,
      added: null,
    },
  ]);
  assert.deepEqual(
    sent.filter((s) => s.path.includes("changelog")).map((s) => s.path),
    ["/rest/api/3/issue/HELP-1/changelog?startAt=0&maxResults=100"],
  );
});

test("jira_get_issue reads a long changelog newest page first, at most five calls", async () => {
  const h = harness();
  const key = h.addAgent(["read"]);
  // vip in the first page, urgent near the end: both found, the middle never read twice.
  const long = [
    ...filler(10),
    change(CUSTOMER, 6, "", "vip"),
    ...filler(189),
    change(AGENT, 7, "vip", "vip urgent"),
    ...filler(49),
  ];
  const sent = jira(h, { changelog: long });
  const issue = await getIssue(h, key, "HELP-1");
  assert.deepEqual(
    issue.labels_added_by.map((l: { label: string; author_account_id: string | null }) => [
      l.label,
      l.author_account_id,
    ]),
    [
      ["urgent", "agent-1"],
      ["vip", "c-1"],
      ["legacy", null],
    ],
  );
  assert.deepEqual(
    sent.filter((s) => s.path.includes("changelog")).map((s) => s.path.split("?")[1]),
    ["startAt=0&maxResults=100", "startAt=150&maxResults=100", "startAt=100&maxResults=50"],
  );

  // Past five calls the oldest changes are not read, and a label only they add stays unknown.
  const h2 = harness();
  const key2 = h2.addAgent(["read"]);
  const sent2 = jira(h2, {
    changelog: [...filler(10), change(CUSTOMER, 6, "", "vip"), ...filler(989)],
  });
  const issue2 = await getIssue(h2, key2, "HELP-1");
  assert.equal(issue2.labels_added_by[1].author_account_id, null);
  assert.equal(sent2.filter((s) => s.path.includes("changelog")).length, 5);
});

test("jira_get_issue still answers when the changelog cannot be read", async () => {
  const h = harness();
  const key = h.addAgent(["read"]);
  jira(h, { changelog: "fail" });
  const r = await h.callTool(key, "jira_get_issue", { issue_key: "HELP-1" });
  assert.equal(r.isError, false);
  const issue = JSON.parse(r.text!);
  assert.equal(issue.labels_added_by, null);
  assert.deepEqual(issue.labels, ["urgent", "vip", "legacy"]);
  assert.equal(h.state.audit.at(-1)!.decision, "allowed");
});

test("jira_get_issue reads no changelog for an issue outside the agent's projects", async () => {
  const h = harness();
  const key = h.addAgent(["read"], { projects: ["KAN"] });
  const sent = jira(h);
  const r = await h.callTool(key, "jira_get_issue", { issue_key: "HELP-1" });
  assert.equal(r.isError, true);
  assert.deepEqual(
    sent.map((s) => s.path),
    [
      "/rest/api/3/issue/HELP-1?fields=summary,status,project,issuetype,assignee,priority," +
        "updated,reporter,created,labels,description",
    ],
  );
});

test("jira_add_comment stays public by default and says so", async () => {
  const h = harness();
  const key = h.addAgent(["read", "comment"]);
  const sent = jira(h);
  const r = await h.callTool(key, "jira_add_comment", { issue_key: "HELP-1", body: "Hello" });
  assert.deepEqual(JSON.parse(r.text!), {
    id: "100",
    created: "2026-10-25T10:00:00.000+0000",
    internal: false,
  });
  const post = sent.find((s) => s.method === "POST")!;
  assert.equal(post.path, "/rest/api/3/issue/HELP-1/comment");
  assert.ok(!sent.some((s) => s.path.includes("servicedeskapi")));
});

test("jira_add_comment says when Jira kept a plain comment internal", async () => {
  const h = harness();
  const key = h.addAgent(["read", "comment"]);
  jira(h, { jsdPublic: false });
  const r = await h.callTool(key, "jira_add_comment", { issue_key: "HELP-1", body: "Hello" });
  assert.equal(JSON.parse(r.text!).internal, true);
});

test("jira_add_comment gives internal null outside a service project", async () => {
  const h = harness();
  const key = h.addAgent(["read", "comment"]);
  jira(h);
  const r = await h.callTool(key, "jira_add_comment", { issue_key: "KAN-1", body: "Hello" });
  assert.equal(JSON.parse(r.text!).internal, null);
});

test("jira_add_comment with internal adds a service project's internal note", async () => {
  const h = harness();
  const key = h.addAgent(["read", "comment"]);
  const sent = jira(h);
  const r = await h.callTool(key, "jira_add_comment", {
    issue_key: "HELP-1",
    body: "Checked the logs",
    internal: true,
  });
  assert.deepEqual(JSON.parse(r.text!), {
    id: "101",
    created: "2026-10-25T10:00:00.000+0000",
    internal: true,
  });
  const post = sent.find((s) => s.method === "POST")!;
  assert.equal(post.path, "/rest/api/3/issue/HELP-1/comment");
  assert.deepEqual((post.body as { properties: unknown }).properties, [
    { key: "sd.public.comment", value: { internal: true } },
  ]);
  assert.ok(!sent.some((s) => s.path.includes("servicedeskapi") || s.path.includes("/project/")));
  assert.equal(h.state.audit.at(-1)!.target, "HELP-1 (internal)");
});

test("jira_add_comment says plainly when Jira posted an internal note publicly", async () => {
  const h = harness();
  const key = h.addAgent(["read", "comment"]);
  const sent = jira(h, { jsdPublic: true });
  const r = await h.callTool(key, "jira_add_comment", { issue_key: "HELP-1", body: "x", internal: true });
  // Not an error: the comment exists, and a retry would post a second public copy.
  assert.equal(r.isError, false);
  const answer = JSON.parse(r.text!);
  assert.equal(answer.id, "101");
  assert.equal(answer.internal, false);
  assert.match(answer.warning, /comment 101 WAS posted on HELP-1 and is PUBLIC/);
  assert.match(answer.warning, /Do not retry/);
  assert.equal(sent.filter((s) => s.method === "POST").length, 1);
  const row = h.state.audit.at(-1)!;
  assert.equal(row.decision, "allowed");
  assert.equal(row.target, "HELP-1 (internal asked, posted PUBLIC as comment 101)");
});

test("jira_add_comment refuses an internal note outside a service project", async () => {
  const h = harness();
  const key = h.addAgent(["read", "comment"]);
  const sent = jira(h);
  const r = await h.callTool(key, "jira_add_comment", {
    issue_key: "KAN-1",
    body: "x",
    internal: true,
  });
  assert.equal(r.isError, true);
  assert.match(r.text, /not in a service project/);
  assert.ok(!sent.some((s) => s.method === "POST"));
  assert.equal(h.state.audit.at(-1)!.decision, "denied");
});

test("jira_add_comment refuses an issue outside the agent's projects before any write", async () => {
  const h = harness();
  const key = h.addAgent(["read", "comment"], { projects: ["KAN"] });
  const sent = jira(h);
  const r = await h.callTool(key, "jira_add_comment", {
    issue_key: "HELP-1",
    body: "x",
    internal: true,
  });
  assert.equal(r.isError, true);
  assert.ok(!sent.some((s) => s.method === "POST" || s.path.includes("/project/")));
});
