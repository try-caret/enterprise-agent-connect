import assert from "node:assert/strict";
import { test } from "node:test";
import { gongTools, MAX_PAGES } from "./tools.ts";
import { loadConnectors } from "../../packages/connector-kit/src/connector.ts";
import type { ToolContext } from "../../packages/connector-kit/src/common.ts";
import type { VendorRequest, VendorResponse } from "../../packages/connector-kit/src/types.ts";

// The Gong tools against a fake Gong (answers shaped as /v2/calls/extensive, /v2/calls/transcript,
// /v2/users and /v2/workspaces return them). Call 1 is with Acme (acme.com, CRM account
// 001ACME) in workspace 10, call 2 with Globex in workspace 20, call 3 is private.

const party = (name: string, email: string, affiliation: string, speakerId: string) => ({
  id: `p-${speakerId}`,
  name,
  emailAddress: email,
  title: "VP",
  affiliation,
  speakerId,
  userId: affiliation === "Internal" ? "900" : undefined,
  phoneNumber: "+1 555 0100",
  methods: ["Invitee"],
});
const call = (id: string, ws: string, customer: string, account: string, extra: object = {}) => ({
  metaData: {
    id,
    url: `https://app.gong.io/call?id=${id}`,
    title: `Call with ${customer}`,
    started: "2026-10-20T15:00:00Z",
    duration: 1800,
    primaryUserId: "900",
    direction: "Conference",
    scope: "External",
    workspaceId: ws,
    purpose: "Renewal",
    isPrivate: false,
    meetingUrl: "https://zoom.us/j/1",
    ...extra,
  },
  context: [
    {
      system: "Salesforce",
      objects: [
        {
          objectType: "Account",
          objectId: account,
          fields: [
            { name: "Name", value: customer },
            { name: "Industry", value: "Manufacturing" },
          ],
        },
        {
          objectType: "Opportunity",
          objectId: `006${id}`,
          fields: [
            { name: "Name", value: `${customer} renewal` },
            { name: "Amount", value: 50000 },
          ],
        },
      ],
    },
  ],
  parties: [
    party("Rita Rep", "rita@seller.com", "Internal", "s1"),
    party(`Cora ${customer}`, `cora@${customer.toLowerCase()}.com`, "External", "s2"),
  ],
  content: {
    brief: "They want to renew with more seats.",
    keyPoints: [{ text: "Budget approved" }],
    highlights: [
      {
        title: "Next steps",
        items: [{ text: "Send quote", startTimes: [100] }],
      },
    ],
    outline: [
      {
        section: "Pricing",
        startTime: 60,
        duration: 300,
        items: [{ text: "Seats" }],
      },
    ],
    topics: [
      { name: "Pricing", duration: 300 },
      { name: "Small talk", duration: 0 },
    ],
    callOutcome: { id: "1", category: "Positive", name: "Moving forward" },
    trackers: [{ name: "Competitors", count: 2 }],
  },
  interaction: { speakers: [] },
});
const CALLS = [
  call("1", "10", "Acme", "001ACME"),
  call("2", "20", "Globex", "001GLOBEX"),
  call("3", "10", "Acme", "001ACME", { isPrivate: true }),
];
const TRANSCRIPT = [
  {
    speakerId: "s1",
    topic: "Pricing",
    sentences: [
      { start: 1000, end: 4000, text: "Thanks for joining." },
      { start: 4000, end: 9000, text: "Let us talk pricing." },
    ],
  },
  {
    speakerId: "s2",
    topic: "Pricing",
    sentences: [
      {
        start: 70000,
        end: 80000,
        text: "Our budget is approved for PRICING next year.",
      },
    ],
  },
  {
    speakerId: "s9",
    sentences: [{ start: 120000, end: 121000, text: "Bye." }],
  },
];
const USERS = Array.from({ length: 3 }, (_, i) => ({
  id: String(900 + i),
  emailAddress: `user${i}@seller.com`,
  emailAliases: i === 2 ? ["alias@seller.com"] : [],
  firstName: ["Rita", "Sam", "Tom"][i],
  lastName: "Rep",
  title: "AE",
  active: true,
  managerId: "800",
  phoneNumber: "+1 555 0101",
  settings: { webConferencesRecorded: true },
}));

type Options = { forced?: Record<string, VendorResponse>; pageSize?: number };

function setup(limits: Record<string, unknown> = {}, options: Options = {}) {
  const requests: VendorRequest[] = [];
  const audit: {
    tool: string;
    target: string | null;
    decision: string;
    reason: string | null;
  }[] = [];
  const size = options.pageSize ?? 100;
  // Gong's cursor is opaque; here it is the next page's offset.
  const paged = (all: object[], cursor: unknown, key: string) => {
    const at = cursor ? Number(cursor) : 0;
    const items = all.slice(at, at + size);
    if (!items.length) return { status: 404, data: { errors: ["No records found"] } };
    const more = at + size < all.length;
    return {
      status: 200,
      data: {
        requestId: "r",
        records: {
          totalRecords: all.length,
          ...(more ? { cursor: String(at + size) } : {}),
        },
        [key]: items,
      },
    };
  };
  const call = async (req: VendorRequest): Promise<VendorResponse> => {
    requests.push(req);
    const [path, query = ""] = req.path.split("?");
    const forced = options.forced?.[`${req.method} ${path}`];
    if (forced) return forced;
    const body = req.body as { filter: Record<string, any>; cursor?: string } | undefined;
    if (path === "/v2/calls/extensive") {
      const f = body!.filter;
      const found = CALLS.filter(
        (c) =>
          (!f.callIds || f.callIds.includes(c.metaData.id)) &&
          (!f.workspaceId || f.workspaceId === c.metaData.workspaceId),
      ).map((c) =>
        (req.body as any).contentSelector.exposedFields.content ? c : { ...c, content: undefined },
      );
      return paged(found, body!.cursor, "calls");
    }
    if (path === "/v2/calls/transcript")
      return {
        status: 200,
        data: {
          callTranscripts: [{ callId: body!.filter.callIds[0], transcript: TRANSCRIPT }],
        },
      };
    if (path === "/v2/workspaces")
      return {
        status: 200,
        data: {
          workspaces: [
            { id: "10", name: "Sales", description: "Field" },
            { id: "20", name: "Partners", description: null },
          ],
        },
      };
    if (path === "/v2/users")
      return paged(USERS, new URLSearchParams(query).get("cursor"), "users");
    const m = /^\/v2\/users\/(\d+)$/.exec(path);
    if (m) {
      const u = USERS.find((x) => x.id === m[1]);
      return u
        ? { status: 200, data: { user: u } }
        : { status: 404, data: { errors: ["User not found"] } };
    }
    return { status: 500, data: {} };
  };
  const ctx = {
    limits,
    config: { host: "us-1.api.gong.io" },
    call,
    async audit(e) {
      audit.push(e);
    },
  } as ToolContext;
  const tools = new Map<string, (args: unknown) => Promise<any>>();
  const server = {
    registerTool: (name: string, def: any, fn: (args: unknown) => Promise<any>) => {
      tools.set(name, (args: unknown) => fn(def.inputSchema.parse(args)));
    },
  };
  gongTools.register(server as never, ctx, new Set(gongTools.names));
  const run = async (name: string, args: object = {}) => {
    const r = await tools.get(name)!(args);
    const text = r.content[0].text as string;
    return {
      isError: Boolean(r.isError),
      text,
      body: r.isError ? null : JSON.parse(text),
    };
  };
  return { run, requests, audit };
}

const ids = (r: { body: { calls: { id: string }[] } }) => r.body.calls.map((c) => c.id);

test("every connector capability's tools are Gong tools, and every Gong tool is in one", () => {
  const gong = loadConnectors().find((c) => c.id === "gong")!;
  assert.deepEqual(gong.capabilities.flatMap((c) => c.tools).sort(), [...gongTools.names].sort());
});

test("gong_list_calls lists trimmed calls in a window, never private ones", async () => {
  const { run, requests } = setup();
  const r = await run("gong_list_calls", {
    from: "2026-10-01",
    to: "2026-10-27",
  });
  assert.deepEqual(ids(r), ["1", "2"]);
  assert.deepEqual(r.body.calls[0], {
    id: "1",
    title: "Call with Acme",
    started: "2026-10-20T15:00:00Z",
    duration_seconds: 1800,
    direction: "Conference",
    scope: "External",
    workspace_id: "10",
    host_user_id: "900",
    url: "https://app.gong.io/call?id=1",
    participants: [
      {
        name: "Rita Rep",
        email: "rita@seller.com",
        title: "VP",
        affiliation: "Internal",
        gong_user_id: "900",
      },
      {
        name: "Cora Acme",
        email: "cora@acme.com",
        title: "VP",
        affiliation: "External",
        gong_user_id: null,
      },
    ],
    more_participants: false,
    crm: [
      { system: "Salesforce", type: "Account", id: "001ACME", name: "Acme" },
      {
        system: "Salesforce",
        type: "Opportunity",
        id: "0061",
        name: "Acme renewal",
      },
    ],
  });
  assert.equal(r.body.next, null);
  assert.ok(!r.text.includes("555") && !r.text.includes("zoom.us"));
  const sent = requests[0].body as any;
  assert.equal(requests[0].path, "/v2/calls/extensive");
  assert.deepEqual(sent.filter, {
    fromDateTime: "2026-10-01T00:00:00.000Z",
    toDateTime: "2026-10-27T00:00:00.000Z",
  });
  assert.equal(sent.contentSelector.exposedFields.parties, true);
  assert.equal(sent.contentSelector.exposedFields.content, undefined);
});

test("gong_list_calls defaults to the last 7 days and filters by participant and host", async () => {
  const { run, requests } = setup();
  const r = await run("gong_list_calls", {
    participant: "Globex.com",
    host_user_id: "900",
  });
  assert.deepEqual(ids(r), ["2"]);
  const f = (requests[0].body as any).filter;
  assert.equal(Date.parse(f.toDateTime) - Date.parse(f.fromDateTime), 7 * 86400_000);
  assert.deepEqual(f.primaryUserIds, ["900"]);
  assert.deepEqual(ids(await run("gong_list_calls", { participant: "cora@acme.com" })), ["1"]);
  assert.deepEqual(ids(await run("gong_list_calls", { participant: "nobody@acme.com" })), []);
});

test("gong_list_calls treats Gong's 404 for nothing found as an empty list", async () => {
  const { run } = setup({ workspaces: ["30"] });
  const r = await run("gong_list_calls", {});
  assert.equal(r.isError, false);
  assert.deepEqual(r.body.calls, []);
});

test("gong_list_calls pages with a cursor that resumes inside Gong's page", async () => {
  const { run, requests } = setup({}, { pageSize: 2 });
  const window = { from: "2026-10-01", to: "2026-10-27" };
  const first = await run("gong_list_calls", { ...window, max_results: 1 });
  assert.deepEqual(ids(first), ["1"]);
  assert.ok(first.body.next);
  // The window rides in the cursor: a later page reads the same one without being told.
  const second = await run("gong_list_calls", {
    max_results: 1,
    cursor: first.body.next,
  });
  assert.deepEqual(ids(second), ["2"]);
  assert.equal((requests[1].body as any).filter.fromDateTime, "2026-10-01T00:00:00.000Z");
  assert.equal((requests[1].body as any).cursor, undefined);
  // Call 2 ended Gong's first page; the next one holds only the private call.
  const third = await run("gong_list_calls", {
    max_results: 1,
    cursor: second.body.next,
  });
  assert.equal((requests.at(-1)!.body as any).cursor, "2");
  assert.deepEqual(ids(third), []);
  assert.equal(third.body.next, null);
  assert.equal((await run("gong_list_calls", { cursor: "junk" })).isError, true);
});

test("a filtered listing reads at most MAX_PAGES of Gong's pages per call", async () => {
  const { run, requests } = setup({ domains: ["nowhere.com"] }, { pageSize: 1 });
  const all = Array.from({ length: MAX_PAGES + 2 }, () => CALLS[1]);
  CALLS.push(...all);
  try {
    const r = await run("gong_list_calls", {});
    assert.deepEqual(r.body.calls, []);
    assert.equal(requests.length, MAX_PAGES);
    assert.ok(r.body.next);
  } finally {
    CALLS.splice(3);
  }
});

test("a workspaces limit is sent to Gong when it is one, and checked on every row", async () => {
  const one = setup({ workspaces: ["10"] });
  assert.deepEqual(ids(await one.run("gong_list_calls", {})), ["1"]);
  assert.equal((one.requests[0].body as any).filter.workspaceId, "10");
  const refused = await one.run("gong_list_calls", { workspace_id: "20" });
  assert.equal(refused.isError, true);
  assert.equal(one.audit.at(-1)!.decision, "denied");
  assert.equal(one.requests.length, 1);

  const two = setup({ workspaces: ["10", "30"] });
  assert.deepEqual(ids(await two.run("gong_list_calls", {})), ["1"]);
  assert.equal((two.requests[0].body as any).filter.workspaceId, undefined);
  assert.equal((await two.run("gong_get_call", { call_id: "2" })).isError, true);
});

test("customer limits keep calls to a domain or a CRM account", async () => {
  const domain = setup({ domains: ["acme.com"] });
  assert.deepEqual(ids(await domain.run("gong_list_calls", {})), ["1"]);
  assert.deepEqual(ids(await domain.run("gong_list_calls", { participant: "globex.com" })), []);
  const account = setup({ crm_accounts: ["001GLOBEX"] });
  assert.deepEqual(ids(await account.run("gong_list_calls", {})), ["2"]);
  const either = setup({ domains: ["acme.com"], crm_accounts: ["001GLOBEX"] });
  assert.deepEqual(ids(await either.run("gong_list_calls", {})), ["1", "2"]);
  // A subdomain counts; a look-alike does not.
  CALLS[1].parties[1].emailAddress = "cora@eu.globex.com";
  try {
    assert.deepEqual(ids(await setup({ domains: ["globex.com"] }).run("gong_list_calls", {})), [
      "2",
    ]);
    assert.deepEqual(ids(await setup({ domains: ["lobex.com"] }).run("gong_list_calls", {})), []);
  } finally {
    CALLS[1].parties[1].emailAddress = "cora@globex.com";
  }
});

test("gong_get_call reads one call with Gong's summary and CRM fields", async () => {
  const { run, requests } = setup();
  const r = await run("gong_get_call", { call_id: "1" });
  assert.equal(r.body.purpose, "Renewal");
  assert.equal(r.body.brief, "They want to renew with more seats.");
  assert.deepEqual(r.body.key_points, ["Budget approved"]);
  assert.deepEqual(r.body.highlights, [{ title: "Next steps", items: ["Send quote"] }]);
  assert.deepEqual(r.body.outline, [{ section: "Pricing", start_seconds: 60, items: ["Seats"] }]);
  assert.deepEqual(r.body.topics, [{ name: "Pricing", seconds: 300 }]);
  assert.equal(r.body.outcome, "Moving forward");
  assert.deepEqual(r.body.crm[1], {
    system: "Salesforce",
    type: "Opportunity",
    id: "0061",
    fields: { Name: "Acme renewal", Amount: 50000 },
  });
  assert.ok(!r.text.includes("Competitors"));
  const sent = requests[0].body as any;
  assert.deepEqual(sent.filter, { callIds: ["1"] });
  assert.equal(sent.contentSelector.exposedFields.content.brief, true);
});

test("a call outside the limits, private or missing gets the same answer", async () => {
  const { run, audit } = setup({ domains: ["acme.com"] });
  const answers = [];
  for (const call_id of ["2", "3", "404"]) {
    const r = await run("gong_get_call", { call_id });
    assert.equal(r.isError, true, call_id);
    answers.push(r.text);
  }
  assert.equal(new Set(answers).size, 1);
  assert.equal(answers[0], "No call with that id that this agent may read");
  assert.deepEqual(
    audit.map((a) => a.decision),
    ["denied", "denied", "failed"],
  );
});

test("a malformed call id never reaches Gong", async () => {
  const { run, requests } = setup();
  for (const call_id of ["", "1 OR 1", "../users", "abc"])
    assert.equal((await run("gong_get_call", { call_id })).isError, true, call_id);
  assert.equal(requests.length, 0);
});

test("gong_get_transcript names speakers and narrows by time and phrase", async () => {
  const { run, requests } = setup();
  const r = await run("gong_get_transcript", { call_id: "1" });
  assert.deepEqual(r.body.call, {
    id: "1",
    title: "Call with Acme",
    started: "2026-10-20T15:00:00Z",
  });
  assert.deepEqual(r.body.turns, [
    {
      speaker: "Rita Rep",
      affiliation: "Internal",
      start_seconds: 1,
      end_seconds: 9,
      text: "Thanks for joining. Let us talk pricing.",
    },
    {
      speaker: "Cora Acme",
      affiliation: "External",
      start_seconds: 70,
      end_seconds: 80,
      text: "Our budget is approved for PRICING next year.",
    },
    {
      speaker: "Unknown speaker",
      affiliation: null,
      start_seconds: 120,
      end_seconds: 121,
      text: "Bye.",
    },
  ]);
  assert.equal(r.body.next, null);
  assert.deepEqual(requests.at(-1)!.body, { filter: { callIds: ["1"] } });

  const window = await run("gong_get_transcript", {
    call_id: "1",
    from_seconds: 60,
    to_seconds: 100,
  });
  assert.deepEqual(
    window.body.turns.map((t: { speaker: string }) => t.speaker),
    ["Cora Acme"],
  );
  const phrase = await run("gong_get_transcript", {
    call_id: "1",
    contains: "pricing",
  });
  assert.deepEqual(
    phrase.body.turns.map((t: { text: string }) => t.text),
    ["Let us talk pricing.", "Our budget is approved for PRICING next year."],
  );
});

test("gong_get_transcript pages by turn", async () => {
  const { run } = setup();
  const first = await run("gong_get_transcript", {
    call_id: "1",
    max_turns: 2,
  });
  assert.equal(first.body.turns.length, 2);
  assert.equal(first.body.total_turns, 3);
  assert.equal(first.body.next, "2");
  const rest = await run("gong_get_transcript", {
    call_id: "1",
    max_turns: 2,
    cursor: "2",
  });
  assert.deepEqual(
    rest.body.turns.map((t: { text: string }) => t.text),
    ["Bye."],
  );
  assert.equal(rest.body.next, null);
  assert.equal((await run("gong_get_transcript", { call_id: "1", cursor: "-1" })).isError, true);
});

test("a transcript outside the limits is refused before Gong is asked for it", async () => {
  const { run, requests } = setup({ crm_accounts: ["001ACME"] });
  assert.equal((await run("gong_get_transcript", { call_id: "1" })).isError, false);
  for (const call_id of ["2", "3"]) {
    const r = await run("gong_get_transcript", { call_id });
    assert.equal(r.text, "No call with that id that this agent may read");
  }
  assert.equal(requests.filter((q) => q.path === "/v2/calls/transcript").length, 1);
});

test("gong_list_workspaces lists workspaces, only the key's own under a limit", async () => {
  assert.deepEqual((await setup().run("gong_list_workspaces")).body.workspaces, [
    { id: "10", name: "Sales", description: "Field" },
    { id: "20", name: "Partners", description: null },
  ]);
  assert.deepEqual(
    (await setup({ workspaces: ["20"] }).run("gong_list_workspaces")).body.workspaces,
    [{ id: "20", name: "Partners", description: null }],
  );
});

test("gong_list_users lists trimmed users, searches by name or email, and pages", async () => {
  const { run, requests } = setup({}, { pageSize: 2 });
  const r = await run("gong_list_users", { max_results: 2 });
  assert.deepEqual(r.body.users[0], {
    id: "900",
    email: "user0@seller.com",
    name: "Rita Rep",
    title: "AE",
    active: true,
    manager_id: "800",
  });
  assert.ok(!r.text.includes("555") && !r.text.includes("webConferences"));
  const rest = await run("gong_list_users", {
    max_results: 2,
    cursor: r.body.next,
  });
  assert.deepEqual(
    rest.body.users.map((u: { id: string }) => u.id),
    ["902"],
  );
  assert.equal(requests.at(-1)!.path, "/v2/users?cursor=2");
  const found = await run("gong_list_users", { query: "tom" });
  assert.deepEqual(
    found.body.users.map((u: { id: string }) => u.id),
    ["902"],
  );
});

test("gong_get_user finds a user by id or by email or alias", async () => {
  const { run } = setup({}, { pageSize: 1 });
  assert.equal((await run("gong_get_user", { user: "901" })).body.name, "Sam Rep");
  assert.equal((await run("gong_get_user", { user: "USER1@seller.com" })).body.id, "901");
  assert.equal((await run("gong_get_user", { user: "alias@seller.com" })).body.id, "902");
  const missing = await run("gong_get_user", { user: "ghost@seller.com" });
  assert.equal(missing.text, "No Gong user with that email address");
  assert.equal(
    (await run("gong_get_user", { user: "999" })).text,
    "Gong answered 404: User not found",
  );
  assert.equal((await run("gong_get_user", { user: "not an email" })).isError, true);
});

test("Gong's refusals and rate limits reach the agent in words and the audit", async () => {
  const { run, audit } = setup(
    {},
    {
      forced: {
        "POST /v2/calls/extensive": { status: 429, data: { retryAfter: 3600 } },
        "GET /v2/workspaces": {
          status: 401,
          data: { requestId: "r", errors: ["Invalid access key"] },
        },
      },
    },
  );
  const limited = await run("gong_list_calls", {});
  assert.equal(
    limited.text,
    "Gong is limiting how often the app may call it. Try again in 3600 seconds.",
  );
  assert.deepEqual(audit.at(-1), {
    tool: "gong_list_calls",
    target: null,
    decision: "failed",
    reason: limited.text,
    vendorStatus: 429,
  });
  assert.equal((await run("gong_get_call", { call_id: "1" })).isError, true);
  const refused = await run("gong_list_workspaces");
  assert.equal(refused.text, "Gong answered 401: Invalid access key");
});
