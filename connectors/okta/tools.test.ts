import assert from "node:assert/strict";
import { test } from "node:test";
import { oktaTools } from "./tools.ts";
import { loadConnectors } from "../../packages/connector-kit/src/connector.ts";
import type { ToolContext } from "../../packages/connector-kit/src/common.ts";
import type { VendorRequest, VendorResponse } from "../../packages/connector-kit/src/types.ts";

// The Okta tools against a fake Okta (answers shaped as the vendor returns them: lists as
// {items, after}). Ada is in Sales, Bob in Engineering; a key kept to Sales reaches only Ada.

const id = (prefix: string, n: number) => `${prefix}${String(n).padStart(17, "0")}`;
const ADA = id("00u", 1);
const BOB = id("00u", 2);
const SALES = id("00g", 1);
const ENG = id("00g", 2);
const GONE = id("00g", 9);
const FACTOR = "ufs00000000000000001";

const user = (uid: string, login: string, first: string, status = "ACTIVE") => ({
  id: uid,
  status,
  lastLogin: "2026-10-26T08:00:00.000Z",
  passwordChanged: "2026-09-01T00:00:00.000Z",
  statusChanged: null,
  created: "2025-01-01T00:00:00.000Z",
  profile: {
    login,
    email: login,
    firstName: first,
    lastName: "Lovelace",
    title: "Engineer",
    mobilePhone: "+1 555",
  },
  credentials: { password: {} },
  _links: { self: { href: "https://x" } },
});
const USERS: Record<string, ReturnType<typeof user>> = {
  [ADA]: user(ADA, "ada@acme.com", "Ada", "LOCKED_OUT"),
  [BOB]: user(BOB, "bob@acme.com", "Bob"),
};
const group = (gid: string, name: string) => ({
  id: gid,
  type: "OKTA_GROUP",
  profile: { name, description: `${name} team` },
  _links: {},
});
const GROUPS: Record<string, ReturnType<typeof group>> = {
  [SALES]: group(SALES, "Sales"),
  [ENG]: group(ENG, "Engineering"),
};
const MEMBERS: Record<string, string[]> = { [SALES]: [ADA], [ENG]: [BOB] };

type Options = { status?: Record<string, VendorResponse> };

function setup(limits: Record<string, unknown> = {}, options: Options = {}) {
  const requests: VendorRequest[] = [];
  const audit: { tool: string; target: string | null; decision: string; reason: string | null }[] =
    [];
  const list = (items: object[], after: string | null = null) => ({
    status: 200,
    data: { items, after },
  });
  const call = async (req: VendorRequest): Promise<VendorResponse> => {
    requests.push(req);
    const [path, query = ""] = req.path.split("?");
    const params = new URLSearchParams(query);
    const forced = options.status?.[`${req.method} ${path}`];
    if (forced) return forced;
    const byRef = (ref: string) =>
      Object.values(USERS).find((u) => u.id === ref || u.profile.login === ref);
    let m;
    if ((m = /^\/api\/v1\/users\/([^/]+)$/.exec(path)) && req.method === "GET") {
      const u = byRef(decodeURIComponent(m[1]));
      return u
        ? { status: 200, data: u }
        : {
            status: 404,
            data: { errorCode: "E0000007", errorSummary: "Not found: Resource not found" },
          };
    }
    if ((m = /^\/api\/v1\/users\/(\w+)\/groups$/.exec(path)))
      return list(
        Object.keys(MEMBERS)
          .filter((g) => MEMBERS[g].includes(m![1]))
          .map((g) => GROUPS[g]),
      );
    if ((m = /^\/api\/v1\/users\/(\w+)\/factors$/.exec(path)))
      return list([
        {
          id: FACTOR,
          factorType: "push",
          provider: "OKTA",
          vendorName: "OKTA",
          status: "ACTIVE",
          created: "2026-01-01T00:00:00.000Z",
          lastUpdated: "2026-01-02T00:00:00.000Z",
          profile: { deviceToken: "secret-device" },
        },
      ]);
    if (path === "/api/v1/users") {
      const all = Object.values(USERS);
      return params.get("after") ? list(all.slice(1)) : list(all.slice(0, 1), "page2");
    }
    if ((m = /^\/api\/v1\/groups\/(\w+)$/.exec(path)))
      return GROUPS[m[1]]
        ? { status: 200, data: GROUPS[m[1]] }
        : { status: 404, data: { errorSummary: "Not found" } };
    if (path === "/api/v1/groups") return list(Object.values(GROUPS), "g2");
    if ((m = /^\/api\/v1\/groups\/(\w+)\/users$/.exec(path)))
      return list(MEMBERS[m[1]].map((u) => USERS[u]));
    if (/^\/api\/v1\/groups\/\w+\/users\/\w+$/.test(path)) return { status: 204, data: null };
    if (path === "/api/v1/apps")
      return list(
        [
          {
            id: "0oa1",
            label: "Slack",
            name: "slack",
            status: "ACTIVE",
            signOnMode: "SAML_2_0",
            settings: { x: 1 },
          },
        ],
        "a2",
      );
    if (path === "/api/v1/logs")
      return list(
        [
          {
            published: "2026-10-26T08:00:00.000Z",
            eventType: "user.session.start",
            displayMessage: "User login to Okta",
            outcome: { result: "SUCCESS" },
            actor: { id: ADA, type: "User", displayName: "Ada Lovelace" },
            client: {
              ipAddress: "10.0.0.1",
              geographicalContext: { city: "Paris", state: "UNKNOWN", country: "France" },
              userAgent: { browser: "CHROME", os: "Mac OS X", rawUserAgent: "raw" },
            },
          },
          {
            published: "2026-10-25T08:00:00.000Z",
            eventType: "user.account.reset_password",
            displayMessage: "Reset password",
            outcome: { result: "SUCCESS" },
            actor: { id: BOB, type: "User", displayName: "Bob Admin" },
            client: {},
          },
        ],
        "log-2",
      );
    if (/^\/api\/v1\/users\/\w+\/(lifecycle\/\w+|sessions|factors\/\w+)$/.test(path))
      return { status: 200, data: {} };
    return { status: 500, data: {} };
  };
  const ctx = {
    limits,
    config: { org: "acme.okta.com" },
    call,
    forbidden: { okta_unlock_user: "The app needs okta.users.manage." },
    async audit(e) {
      audit.push(e);
    },
  } as ToolContext;
  const tools = new Map<string, (args: unknown) => Promise<any>>();
  const server = {
    registerTool: (name: string, def: any, fn: (args: unknown) => Promise<any>) => {
      // Parsed as the MCP server would, so defaults apply.
      tools.set(name, (args: unknown) => fn(def.inputSchema.parse(args)));
    },
  };
  oktaTools.register(server as never, ctx, new Set(oktaTools.names));
  const run = async (name: string, args: object) => {
    const r = await tools.get(name)!(args);
    const text = r.content[0].text as string;
    return { isError: Boolean(r.isError), text, body: r.isError ? null : JSON.parse(text) };
  };
  const writes = () => requests.filter((r) => r.method !== "GET");
  return { run, requests, audit, writes };
}

const sales = { groups: [SALES] };

test("every connector capability's tools are Okta tools, and every Okta tool is in one", () => {
  const okta = loadConnectors().find((c) => c.id === "okta")!;
  assert.deepEqual(okta.capabilities.flatMap((c) => c.tools).sort(), [...oktaTools.names].sort());
});

test("okta_get_user answers the trimmed person with status and groups", async () => {
  const { run, requests } = setup();
  const r = await run("okta_get_user", { user: "ada@acme.com" });
  assert.deepEqual(r.body, {
    id: ADA,
    login: "ada@acme.com",
    email: "ada@acme.com",
    name: "Ada Lovelace",
    title: "Engineer",
    department: null,
    status: "LOCKED_OUT",
    locked_out: true,
    last_login: "2026-10-26T08:00:00.000Z",
    password_changed: "2026-09-01T00:00:00.000Z",
    status_changed: null,
    created: "2025-01-01T00:00:00.000Z",
    groups: [{ id: SALES, name: "Sales", description: "Sales team", type: "OKTA_GROUP" }],
    more_groups: false,
  });
  assert.ok(!r.text.includes("555") && !r.text.includes("_links"));
  assert.equal(requests[0].path, "/api/v1/users/ada%40acme.com");
  assert.equal(requests[1].path, `/api/v1/users/${ADA}/groups?limit=200`);
});

test("a groups limit keeps every person tool to its members, and hides who exists", async () => {
  const { run, audit, requests } = setup(sales);
  assert.equal((await run("okta_get_user", { user: ADA })).isError, false);
  for (const ref of ["bob@acme.com", "ghost@acme.com"]) {
    const r = await run("okta_get_user", { user: ref });
    assert.equal(r.isError, true);
    assert.equal(r.text, "This agent may not look up or act on that person");
    assert.equal(audit.at(-1)!.decision, "denied");
  }
  for (const tool of ["okta_list_factors", "okta_get_user_activity", "okta_list_apps"])
    assert.equal((await run(tool, { user: BOB })).isError, true, tool);
  assert.ok(
    !requests.some((r) => r.path.includes(`${BOB}/factors`) || r.path.startsWith("/api/v1/logs")),
  );
  assert.ok(!requests.some((r) => r.path.startsWith("/api/v1/apps")));
});

test("a refused person or reference never reaches Okta", async () => {
  const { run, requests } = setup();
  for (const bad of ["", "a b", "x/../y", "a?b=c", 'a"b'])
    assert.equal((await run("okta_get_user", { user: bad })).isError, true, bad);
  assert.equal(requests.length, 0);
});

test("okta_search_users builds Okta's search and pages with a cursor", async () => {
  const { run, requests } = setup();
  const r = await run("okta_search_users", {
    query: 'ad"a',
    status: "LOCKED_OUT",
    max_results: 10,
  });
  assert.deepEqual(
    r.body.people.map((p: { id: string }) => p.id),
    [ADA],
  );
  assert.equal(r.body.more, true);
  assert.equal(r.body.next, "page2");
  const sent = new URLSearchParams(requests[0].path.split("?")[1]);
  assert.equal(
    sent.get("search"),
    '(profile.login sw "ada" or profile.email sw "ada" or profile.firstName sw "ada" or profile.lastName sw "ada") and status eq "LOCKED_OUT"',
  );
  assert.equal(sent.get("limit"), "10");
  const again = await run("okta_search_users", { query: "a", cursor: "page2" });
  assert.deepEqual(
    again.body.people.map((p: { id: string }) => p.id),
    [BOB],
  );
  assert.match(requests.at(-1)!.path, /&after=page2$/);
  assert.equal((await run("okta_search_users", {})).isError, true);
});

test("okta_search_users under a groups limit keeps members and reads on", async () => {
  const { run } = setup(sales);
  const r = await run("okta_search_users", { query: "a", max_results: 5 });
  // Ada on the first page, Bob on the second (dropped), then the listing ends.
  assert.deepEqual(
    r.body.people.map((p: { id: string }) => p.id),
    [ADA],
  );
  assert.equal(r.body.more, false);
});

test("okta_list_factors gives factor ids without device details", async () => {
  const { run } = setup();
  const r = await run("okta_list_factors", { user: ADA });
  assert.deepEqual(r.body.factors, [
    {
      id: FACTOR,
      type: "push",
      provider: "OKTA",
      name: "OKTA",
      status: "ACTIVE",
      enrolled: "2026-01-01T00:00:00.000Z",
      last_updated: "2026-01-02T00:00:00.000Z",
    },
  ]);
  assert.ok(!r.text.includes("secret-device"));
});

test("okta_search_groups searches every group, or only the key's own", async () => {
  const open = setup();
  const all = await open.run("okta_search_groups", { query: "Eng" });
  assert.equal(all.body.groups.length, 2);
  assert.equal(all.body.next, "g2");
  assert.equal(open.requests[0].path, "/api/v1/groups?q=Eng&limit=25");

  const kept = setup({ groups: [SALES, GONE] });
  const r = await kept.run("okta_search_groups", { query: "sa" });
  assert.deepEqual(r.body, {
    groups: [{ id: SALES, name: "Sales", description: "Sales team", type: "OKTA_GROUP" }],
    more: false,
    next: null,
  });
  assert.equal((await kept.run("okta_search_groups", { query: "eng" })).body.groups.length, 0);
  assert.ok(!kept.requests.some((q) => q.path.startsWith("/api/v1/groups?")));
});

test("okta_list_group_members lists a group, refusing one outside the key's groups", async () => {
  const { run, requests } = setup(sales);
  const r = await run("okta_list_group_members", { group_id: SALES });
  assert.deepEqual(
    r.body.people.map((p: { id: string }) => p.id),
    [ADA],
  );
  const before = requests.length;
  for (const group_id of [ENG, "everyone", "00g'x"]) {
    const refused = await run("okta_list_group_members", { group_id });
    assert.equal(refused.isError, true);
  }
  assert.equal(requests.length, before);
});

test("okta_list_apps lists apps, or a person's, trimmed", async () => {
  const { run, requests } = setup();
  const r = await run("okta_list_apps", { user: "ada@acme.com", query: "sl" });
  assert.deepEqual(r.body, {
    apps: [
      { id: "0oa1", label: "Slack", app: "slack", status: "ACTIVE", sign_on_mode: "SAML_2_0" },
    ],
    more: true,
    next: "a2",
  });
  const sent = new URLSearchParams(requests.at(-1)!.path.split("?")[1]);
  assert.equal(sent.get("filter"), `user.id eq "${ADA}"`);
  assert.equal(sent.get("q"), "sl");
});

test("okta_get_user_activity reads one person's log in a fixed window", async () => {
  const { run, requests } = setup(sales);
  const r = await run("okta_get_user_activity", {
    user: ADA,
    hours: 48,
    event_type: "user.session.start",
  });
  assert.equal(r.body.last_login, "2026-10-26T08:00:00.000Z");
  assert.deepEqual(r.body.events[0], {
    time: "2026-10-26T08:00:00.000Z",
    event_type: "user.session.start",
    message: "User login to Okta",
    outcome: "SUCCESS",
    reason: null,
    by: "this person",
    ip: "10.0.0.1",
    location: "Paris, France",
    client: "CHROME on Mac OS X",
  });
  // Another person's name is left out under a groups limit.
  assert.equal(r.body.events[1].by, "another person");
  assert.ok(!r.text.includes("Bob Admin") && !r.text.includes("raw"));
  const sent = new URLSearchParams(requests.at(-1)!.path.split("?")[1]);
  assert.equal(
    sent.get("filter"),
    `(actor.id eq "${ADA}" or target.id eq "${ADA}") and eventType eq "user.session.start"`,
  );
  const until = sent.get("until")!;
  assert.equal(Date.parse(until) - Date.parse(sent.get("since")!), 48 * 3600_000);
  assert.equal(sent.get("sortOrder"), "DESCENDING");
  assert.equal(r.body.next, `${until}~log-2`);

  // The next page reads the same window.
  await run("okta_get_user_activity", { user: ADA, hours: 48, cursor: r.body.next });
  const second = new URLSearchParams(requests.at(-1)!.path.split("?")[1]);
  assert.equal(second.get("until"), until);
  assert.equal(second.get("after"), "log-2");

  const open = setup();
  const named = await open.run("okta_get_user_activity", { user: ADA });
  assert.equal(named.body.events[1].by, "Bob Admin");
});

test("okta_get_user_activity refuses a bad event type or cursor before any call", async () => {
  const { run, requests } = setup();
  for (const args of [
    { user: ADA, event_type: 'x" or true' },
    { user: ADA, cursor: "garbage" },
    { user: ADA, cursor: "2026-10-27T00:00:00Z~" },
  ])
    assert.equal((await run("okta_get_user_activity", args)).isError, true);
  assert.equal(requests.length, 0);
});

test("group membership changes check the person and the group before writing", async () => {
  const { run, writes, audit } = setup(sales);
  const added = await run("okta_add_group_member", { user: ADA, group_id: SALES });
  assert.equal(added.body.member, true);
  assert.deepEqual(writes().at(-1), {
    method: "PUT",
    path: `/api/v1/groups/${SALES}/users/${ADA}`,
  });
  const removed = await run("okta_remove_group_member", { user: "ada@acme.com", group_id: SALES });
  assert.equal(removed.body.member, false);
  assert.deepEqual(writes().at(-1), {
    method: "DELETE",
    path: `/api/v1/groups/${SALES}/users/${ADA}`,
  });
  assert.equal(audit.at(-1)!.target, `ada@acme.com from ${SALES}`);

  // Another group, or a person outside the key's groups: refused, nothing written.
  const before = writes().length;
  assert.equal((await run("okta_add_group_member", { user: ADA, group_id: ENG })).isError, true);
  assert.equal((await run("okta_add_group_member", { user: BOB, group_id: SALES })).isError, true);
  assert.equal(
    (await run("okta_remove_group_member", { user: BOB, group_id: SALES })).isError,
    true,
  );
  assert.equal(writes().length, before);
  assert.equal(audit.at(-1)!.decision, "denied");
});

const ACTIONS: [string, string, string][] = [
  ["okta_unlock_user", "POST", `/api/v1/users/${ADA}/lifecycle/unlock`],
  [
    "okta_send_password_reset",
    "POST",
    `/api/v1/users/${ADA}/lifecycle/reset_password?sendEmail=true`,
  ],
  ["okta_suspend_user", "POST", `/api/v1/users/${ADA}/lifecycle/suspend`],
  ["okta_unsuspend_user", "POST", `/api/v1/users/${ADA}/lifecycle/unsuspend`],
  ["okta_clear_sessions", "DELETE", `/api/v1/users/${ADA}/sessions?oauthTokens=true`],
  ["okta_deactivate_user", "POST", `/api/v1/users/${ADA}/lifecycle/deactivate?sendEmail=false`],
];

test("each action reads the person, then calls on their id", async () => {
  for (const [tool, method, path] of ACTIONS) {
    const { run, writes, audit } = setup(sales);
    const r = await run(tool, { user: "ada@acme.com" });
    assert.equal(r.isError, false, tool);
    assert.equal(r.body.user.id, ADA);
    assert.deepEqual(writes(), [{ method, path }], tool);
    assert.equal(audit.at(-1)!.decision, "allowed");
  }
});

test("each action refuses a person outside the key's groups before any write", async () => {
  for (const [tool] of ACTIONS) {
    const { run, writes, audit } = setup(sales);
    for (const user of [BOB, "ghost@acme.com"]) {
      const r = await run(tool, { user });
      assert.equal(r.isError, true, tool);
      assert.equal(audit.at(-1)!.decision, "denied");
    }
    assert.equal(writes().length, 0, tool);
  }
});

test("okta_send_password_reset never answers a reset link", async () => {
  const { run } = setup();
  const r = await run("okta_send_password_reset", { user: ADA });
  assert.ok(!/resetPasswordUrl|http/.test(r.text));
});

test("okta_reset_factor removes one factor of a person in scope", async () => {
  const { run, writes } = setup(sales);
  const r = await run("okta_reset_factor", { user: ADA, factor_id: FACTOR });
  assert.equal(r.body.factor_id, FACTOR);
  assert.deepEqual(writes(), [
    { method: "DELETE", path: `/api/v1/users/${ADA}/factors/${FACTOR}` },
  ]);
  assert.equal((await run("okta_reset_factor", { user: BOB, factor_id: FACTOR })).isError, true);
  assert.equal(
    (await run("okta_reset_factor", { user: ADA, factor_id: "../sessions" })).isError,
    true,
  );
  assert.equal(writes().length, 1);
});

test("Okta's refusals reach the agent with its summary and the capability's hint", async () => {
  const { run, audit } = setup(
    {},
    {
      status: {
        [`POST /api/v1/users/${ADA}/lifecycle/unlock`]: {
          status: 403,
          data: {
            errorCode: "E0000006",
            errorSummary: "You do not have permission to perform the requested action",
            errorCauses: [],
          },
        },
        [`POST /api/v1/users/${ADA}/lifecycle/suspend`]: {
          status: 400,
          data: {
            errorSummary: "Api validation failed: user",
            errorCauses: [{ errorSummary: "Cannot suspend a user that is not active" }],
          },
        },
      },
    },
  );
  const forbidden = await run("okta_unlock_user", { user: ADA });
  assert.equal(
    forbidden.text,
    "Okta answered 403: You do not have permission to perform the requested action. The app needs okta.users.manage.",
  );
  assert.deepEqual(
    [audit.at(-1)!.decision, audit.at(-1)!.reason?.slice(0, 16)],
    ["failed", "Okta answered 40"],
  );
  const invalid = await run("okta_suspend_user", { user: ADA });
  assert.equal(
    invalid.text,
    "Okta answered 400: Api validation failed: user; Cannot suspend a user that is not active",
  );
});

test("destructive tools sit in destructive capabilities", () => {
  const okta = loadConnectors().find((c) => c.id === "okta")!;
  const destructive = okta.capabilities.filter((c) => c.destructive).flatMap((c) => c.tools);
  for (const tool of [
    "okta_reset_factor",
    "okta_suspend_user",
    "okta_clear_sessions",
    "okta_deactivate_user",
  ])
    assert.ok(destructive.includes(tool), tool);
  assert.ok(okta.capabilities.filter((c) => c.destructive).every((c) => !c.required));
});
