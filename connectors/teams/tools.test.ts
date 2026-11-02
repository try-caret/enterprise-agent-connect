import assert from "node:assert/strict";
import { test } from "node:test";
import type { ToolContext } from "../../packages/connector-kit/src/common.ts";
import type { TeamsAccess } from "../../packages/server/src/teams/runtime.ts";
import type { ConversationRow } from "../../packages/server/src/teams/store.ts";
import { teamsTools, type TeamsToolContext } from "./tools.ts";

// The Teams tools with a fake bot (TeamsAccess): one team with one channel.

const TEAM = "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa";
const CHANNEL = "19:general@thread.tacv2";
const SERVICE_URL = "https://smba.trafficmanager.net/amer/";
const PERSON = "44444444-4444-4444-4444-444444444444";

const teamPlace: ConversationRow = {
  botId: "b",
  conversationId: "19:team@thread.tacv2",
  kind: "team",
  serviceUrl: SERVICE_URL,
  teamId: "19:team@thread.tacv2",
  teamGroupId: TEAM,
  name: "Ops",
  userAadId: null,
  removedAt: null,
  limitNoticeAt: null,
};

type Handler = (input: any) => Promise<{ isError?: boolean; content: { text: string }[] }>;

/** The registered tools' handlers, called directly (schema defaults are not applied). */
function tools(teams: Partial<TeamsAccess>, extra: Partial<ToolContext> = {}) {
  const handlers = new Map<string, Handler>();
  const audit: Parameters<ToolContext["audit"]>[0][] = [];
  const ctx: TeamsToolContext = {
    limits: {},
    config: {},
    call: async () => {
      throw new Error("no Microsoft connection here");
    },
    teams: {
      places: async () => [teamPlace],
      listChannels: async () => ({ status: 200, data: { conversations: [{ id: CHANNEL }] } }),
      ...teams,
    } as TeamsAccess,
    audit: async (entry) => {
      audit.push(entry);
    },
    ...extra,
  };
  const server = {
    registerTool: (name: string, _def: unknown, handler: Handler) => handlers.set(name, handler),
  };
  teamsTools.register(server as never, ctx, new Set(teamsTools.names));
  const call = async (name: string, input: unknown) => {
    const r = await handlers.get(name)!(input);
    return { isError: Boolean(r.isError), text: r.content[0].text };
  };
  return { call, audit };
}

const graphMessages = [
  {
    id: "1",
    messageType: "message",
    from: { user: { id: PERSON.toUpperCase(), displayName: "Ann", userIdentityType: "aadUser" } },
    body: { contentType: "text", content: "Please approve" },
  },
  {
    id: "2",
    messageType: "message",
    from: { application: { id: "app-id", displayName: "Helper bot" } },
    body: { contentType: "text", content: "Noted" },
  },
  {
    id: "3",
    messageType: "message",
    // Someone outside the directory: the id names nothing in this organization.
    from: {
      user: {
        id: "55555555-5555-5555-5555-555555555555",
        displayName: "Ann",
        userIdentityType: "anonymousGuest",
      },
    },
    body: { contentType: "text", content: "Me too" },
  },
  { id: "4", messageType: "message", from: null, body: { content: "deleted" } },
];

test("teams_read_messages says who sent each message, not only their display name", async () => {
  const { call } = tools({
    graph: async () => ({ status: 200, data: { value: graphMessages } }),
  });
  const r = await call("teams_read_messages", { team: TEAM, channel: CHANNEL, max_results: 20 });
  assert.equal(r.isError, false);
  const { messages } = JSON.parse(r.text);
  assert.deepEqual(
    messages.map((m: Record<string, unknown>) => [m.from, m.from_id, m.from_kind]),
    [
      ["Ann", PERSON, "user"],
      ["Helper bot", null, "application"],
      ["Ann", null, "user"],
      ["Someone", null, null],
    ],
  );
});

test("teams_read_thread gives the sender's id on the post and on each reply", async () => {
  const { call } = tools({
    graph: async (path) =>
      path.endsWith("/replies?$top=50")
        ? { status: 200, data: { value: [graphMessages[1], graphMessages[0]] } }
        : { status: 200, data: graphMessages[0] },
  });
  const r = await call("teams_read_thread", { team: TEAM, channel: CHANNEL, message_id: "1" });
  const thread = JSON.parse(r.text);
  assert.equal(thread.post.from_id, PERSON);
  assert.equal(thread.post.from_kind, "user");
  assert.deepEqual(
    thread.replies.map((m: Record<string, unknown>) => [m.from_id, m.from_kind]),
    [
      [PERSON, "user"],
      [null, "application"],
    ],
  );
});

// ---- approval cards ----

function approvalBot(post: { status: number; data: unknown }) {
  const log: string[] = [];
  const rows: any[] = [];
  const teams: Partial<TeamsAccess> = {
    createApproval: async (input) => {
      const row = {
        ...input,
        id: "77777777-7777-7777-7777-777777777777",
        agentId: "a",
        botId: "b",
        threadId: null,
        activityId: null,
        status: "pending" as const,
        decidedByAadId: null,
        decidedByName: null,
        decidedAt: null,
        comment: null,
        createdAt: "2026-10-25T10:00:00.000Z",
        expiresAt: "2026-11-01T10:00:00.000Z",
        expired: false,
      };
      rows.push(row);
      return row;
    },
    newThreadCard: async (_place, channel, card) => {
      log.push(`thread ${channel} ${JSON.stringify(card).includes("Action.Execute")}`);
      return post;
    },
    approvalPosted: async (id, thread, activity) => {
      log.push(`posted ${id} ${thread} ${activity}`);
    },
    discardApproval: async (id) => {
      log.push(`discarded ${id}`);
    },
  };
  return { teams, log, rows };
}

const directory: ToolContext["call"] = async (req) =>
  req.path === `/v1.0/users/${encodeURIComponent("ann@contoso.example")}?$select=id,displayName`
    ? { status: 200, data: { id: PERSON, displayName: "Ann" } }
    : { status: 404, data: { error: { message: "Resource not found" } } };

const ask = {
  team: TEAM,
  channel: CHANNEL,
  approver: "ann@contoso.example",
  title: "Laptop",
  text: "Approve?",
};

test("teams_list_teams says whose a direct chat is, and lists no chats for a key held to teams", async () => {
  const direct: ConversationRow = {
    ...teamPlace,
    conversationId: "a:1direct",
    kind: "personal",
    teamId: null,
    teamGroupId: null,
    name: null,
    userAadId: PERSON,
  };
  const group: ConversationRow = { ...direct, conversationId: "19:group@thread.v2", kind: "groupChat", name: "Ops chat", userAadId: null };
  const places = async () => [teamPlace, direct, group];
  const open = await tools({ places }).call("teams_list_teams", {});
  assert.deepEqual(JSON.parse(open.text).chats, [
    { chat: "a:1direct", kind: "direct", name: null, user: PERSON },
    { chat: "19:group@thread.v2", kind: "group", name: "Ops chat", user: null },
  ]);
  const held = await tools({ places }, { limits: { teams: [TEAM] } }).call("teams_list_teams", {});
  assert.deepEqual(JSON.parse(held.text).chats, []);
});

test("teams_request_approval posts a new channel thread and keeps where it went", async () => {
  const bot = approvalBot({
    status: 201,
    data: { id: `${CHANNEL};messageid=900`, activityId: "900" },
  });
  const { call } = tools(bot.teams, { call: directory });
  const r = await call("teams_request_approval", ask);
  assert.equal(r.isError, false, r.text);
  const out = JSON.parse(r.text);
  assert.match(out.approval_id, /^appr_/);
  assert.equal(out.message_id, "900");
  assert.equal(out.expires, "2026-11-01T10:00:00.000Z");
  assert.deepEqual(bot.log, [
    `thread ${CHANNEL} true`,
    `posted 77777777-7777-7777-7777-777777777777 ${CHANNEL};messageid=900 900`,
  ]);
  assert.equal(bot.rows[0].approverAadId, PERSON);
  assert.notEqual(bot.rows[0].tokenHash, out.approval_id);
});

test("teams_request_approval forgets the approval when the card cannot be posted", async () => {
  const bot = approvalBot({ status: 403, data: { message: "MessageWritesBlocked" } });
  const { call } = tools(bot.teams, { call: directory });
  const r = await call("teams_request_approval", ask);
  assert.equal(r.isError, true);
  assert.deepEqual(bot.log, [
    `thread ${CHANNEL} true`,
    "discarded 77777777-7777-7777-7777-777777777777",
  ]);
});

test("teams_request_approval keeps to the key's channels and needs a real approver", async () => {
  const bot = approvalBot({ status: 201, data: {} });
  const looked: string[] = [];
  const { call } = tools(bot.teams, {
    limits: { channels: ["19:other@thread.tacv2"] },
    call: async (req) => {
      looked.push(req.path);
      return directory(req);
    },
  });
  assert.match((await call("teams_request_approval", ask)).text, /cannot reach that channel/);
  const open = tools(bot.teams, { call: directory });
  assert.match(
    (await open.call("teams_request_approval", { ...ask, approver: "Ann; drop" })).text,
    /sign-in name/,
  );
  assert.deepEqual(looked, []);
  assert.equal(bot.rows.length, 0);
});

test("teams_get_approval refuses an id that is not one", async () => {
  const { call } = tools({ approval: async () => null });
  assert.match((await call("teams_get_approval", { approval_id: "nope" })).text, /not valid/);
  assert.match(
    (await call("teams_get_approval", { approval_id: `appr_${"a".repeat(43)}` })).text,
    /No approval of this agent/,
  );
});
