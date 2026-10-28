import assert from "node:assert/strict";
import { test } from "node:test";
import { plainText, slackTools } from "./tools.ts";
import { slackVendor } from "./vendor.ts";
import { loadConnectors } from "../../packages/connector-kit/src/connector.ts";
import type { ToolContext } from "../../packages/connector-kit/src/common.ts";
import type { VendorConnection, VendorRequest } from "../../packages/connector-kit/src/types.ts";
import { assertToolsImplemented } from "../../packages/server/src/service.ts";

// The Slack tools against a fake Slack Web API, reached through the real vendor client (so
// ok:false answers, rate limits and the upload host are exercised too). The bot is in GENERAL
// and SECRET; OTHER exists but the bot is not in it.

const GENERAL = "C0GENERAL";
const SECRET = "C0SECRET";
const OTHER = "C0OTHER";
const TOKEN = "xoxb-test";
const UPLOAD = "https://files.slack.com/upload/v1/abc";

type Reply = { status?: number; body?: unknown; headers?: Record<string, string> };
type Sent = { url: string; method: string; auth: string | null; body: unknown };

const USERS: Record<string, object> = {
  U1: { id: "U1", name: "ada", real_name: "Ada Lovelace", profile: { display_name: "ada" } },
  U2: { id: "U2", name: "bob", real_name: "Bob Ross", profile: { display_name: "" } },
};

const MESSAGES = [
  {
    ts: "1730000002.000200",
    user: "U1",
    text: "ping <@U2> in <#C0SECRET|secret> &lt;see <https://x.example|the doc>&gt; <!here>",
    reply_count: 2,
    thread_ts: "1730000002.000200",
    reactions: [{ name: "eyes", count: 1, users: ["U2"] }],
    files: [{ name: "a.txt", title: "Notes", url_private: "https://files.slack.com/x" }],
    blocks: [{ type: "rich_text" }],
    team: "T1",
  },
  {
    ts: "1730000001.000100",
    bot_id: "B1",
    subtype: "bot_message",
    bot_profile: { name: "Deploys" },
    text: "",
    attachments: [{ fallback: "deployed v2" }],
  },
];

function fakeSlack(overrides: Record<string, Reply | ((url: URL) => Reply)> = {}) {
  const sent: Sent[] = [];
  const fetchImpl = (async (input: string, init: RequestInit) => {
    const url = new URL(input);
    const headers = new Headers(init.headers);
    const body =
      typeof init.body === "string" && headers.get("content-type")?.startsWith("application/json")
        ? JSON.parse(init.body)
        : init.body;
    sent.push({ url: input, method: init.method!, auth: headers.get("authorization"), body });
    const method = url.host === "files.slack.com" ? "upload" : url.pathname.slice(5);
    const override = overrides[method];
    const reply: Reply = override
      ? typeof override === "function"
        ? override(url)
        : override
      : answer(method, url, body as any);
    return new Response(
      typeof reply.body === "string" ? reply.body : JSON.stringify(reply.body ?? { ok: true }),
      { status: reply.status ?? 200, headers: reply.headers },
    );
  }) as unknown as typeof fetch;
  return { sent, fetchImpl };
}

const member = (channel: string | null) =>
  channel === OTHER ? { body: { ok: false, error: "not_in_channel" } } : null;

function answer(method: string, url: URL, body: any): Reply {
  const q = url.searchParams;
  switch (method) {
    case "auth.test":
      return {
        body: {
          ok: true,
          url: "https://acme.slack.com/",
          team: "Acme",
          team_id: "T1",
          user_id: "UBOT",
        },
        headers: { "x-oauth-scopes": "channels:read,groups:read,channels:history" },
      };
    case "users.conversations":
      return q.get("cursor")
        ? { body: { ok: true, channels: [{ id: SECRET, name: "secret", is_private: true }] } }
        : {
            body: {
              ok: true,
              channels: [
                { id: GENERAL, name: "general", topic: { value: "Company news" } },
                { id: "C0RANDOM", name: "random" },
              ],
              response_metadata: { next_cursor: "page2" },
            },
          };
    case "conversations.history":
      return (
        member(q.get("channel")) ?? {
          body: { ok: true, messages: MESSAGES, response_metadata: { next_cursor: "more" } },
        }
      );
    case "conversations.replies":
      return (
        member(q.get("channel")) ?? {
          body: {
            ok: true,
            messages: [MESSAGES[0], { ts: "1730000003.000300", user: "U2", text: "hi" }],
          },
        }
      );
    case "users.info": {
      const user = USERS[q.get("user")!];
      return { body: user ? { ok: true, user } : { ok: false, error: "user_not_found" } };
    }
    case "chat.getPermalink":
      return {
        body: { ok: true, permalink: `https://acme.slack.com/archives/${q.get("channel")}/p1` },
      };
    case "chat.postMessage":
      return (
        member(body.channel) ?? {
          body: { ok: true, channel: body.channel, ts: "1730000009.000900" },
        }
      );
    case "reactions.add":
      return { body: { ok: true } };
    case "files.getUploadURLExternal":
      return { body: { ok: true, upload_url: UPLOAD, file_id: "F1" } };
    case "upload":
      return { body: "OK - 5" };
    case "files.completeUploadExternal":
      return { body: { ok: true, files: [{ id: "F1" }] } };
    case "users.lookupByEmail":
      return q.get("email") === "ada@acme.example"
        ? {
            body: {
              ok: true,
              user: {
                ...USERS.U1,
                tz: "Europe/London",
                profile: { display_name: "ada", title: "Engineer", email: "ada@acme.example" },
              },
            },
          }
        : { body: { ok: false, error: "users_not_found" } };
    case "chat.delete":
      return body.ts === "1730000002.000200"
        ? { body: { ok: false, error: "cant_delete_message" } }
        : { body: { ok: true, channel: body.channel, ts: body.ts } };
    default:
      return { status: 404, body: "not found" };
  }
}

const connection: VendorConnection = { config: { teamId: "T1" }, secret: { bot_token: TOKEN } };

function setup(
  limits: Record<string, unknown> = {},
  overrides: Parameters<typeof fakeSlack>[0] = {},
) {
  const slack = fakeSlack(overrides);
  const waits: number[] = [];
  const vendor = slackVendor(slack.fetchImpl, async (s) => void waits.push(s));
  const audit: { tool: string; target: string | null; decision: string; reason: string | null }[] =
    [];
  const ctx = {
    limits,
    config: {},
    call: (req: VendorRequest) => vendor.call(connection, req),
    forbidden: { slack_post_message: "Add the bot scope chat:write." },
    async audit(e) {
      audit.push(e);
    },
  } as ToolContext;
  const tools = new Map<string, (args: unknown) => Promise<any>>();
  const server = { registerTool: (name: string, _def: unknown, fn: never) => tools.set(name, fn) };
  slackTools.register(server as never, ctx, new Set(slackTools.names));
  const run = async (name: string, args: object) => {
    const r = await tools.get(name)!(args);
    const text = r.content[0].text as string;
    return { isError: Boolean(r.isError), text, body: r.isError ? null : JSON.parse(text) };
  };
  const calls = () => slack.sent.map((s) => new URL(s.url).pathname.replace("/api/", ""));
  return { run, sent: slack.sent, calls, audit, waits, vendor };
}

test("the connector's capabilities, tools and checks line up; deleting is destructive", () => {
  const connector = loadConnectors().find((c) => c.id === "slack")!;
  assertToolsImplemented([{ connector, toolset: slackTools, vendor: slackVendor() }]);
  const destructive = connector.capabilities.filter((c) => c.destructive).map((c) => c.id);
  assert.deepEqual(destructive, ["delete"]);
  assert.deepEqual(connector.capabilities.find((c) => c.required)?.id, "read");
});

test("slack_list_channels lists the bot's channels within the key's limit and pages", async () => {
  const { run, sent } = setup();
  const all = await run("slack_list_channels", {});
  assert.deepEqual(
    all.body.channels.map((c: { id: string }) => c.id),
    [GENERAL, "C0RANDOM"],
  );
  assert.deepEqual(all.body.channels[0], {
    id: GENERAL,
    name: "general",
    private: false,
    topic: "Company news",
    purpose: null,
  });
  assert.equal(all.body.next, "page2");
  assert.match(sent[0].url, /types=public_channel%2Cprivate_channel&exclude_archived=true/);
  assert.equal(sent[0].auth, `Bearer ${TOKEN}`);

  const limited = setup({ channels: [SECRET] });
  const first = await limited.run("slack_list_channels", {});
  assert.deepEqual(first.body.channels, []);
  const second = await limited.run("slack_list_channels", { cursor: "page2" });
  assert.deepEqual(
    second.body.channels.map((c: { id: string }) => c.id),
    [SECRET],
  );
});

test("slack_read_channel trims messages, resolves names and converts markup", async () => {
  const { run, sent, calls } = setup();
  const r = await run("slack_read_channel", {
    channel: GENERAL,
    limit: 20,
    oldest: "2024-10-27T00:00:00Z",
  });
  assert.equal(r.isError, false);
  assert.equal(r.body.next, "more");
  const [human, bot] = r.body.messages;
  assert.deepEqual(human, {
    ts: "1730000002.000200",
    at: "2024-10-27T03:33:22.000Z",
    user_id: "U1",
    user: "ada",
    text: "ping @Bob Ross in #secret <see the doc (https://x.example)> @here",
    thread_ts: "1730000002.000200",
    reply_count: 2,
    reactions: [{ emoji: "eyes", count: 1 }],
    files: ["Notes"],
  });
  assert.deepEqual(bot, {
    ts: "1730000001.000100",
    at: "2024-10-27T03:33:21.000Z",
    bot: "Deploys",
    subtype: "bot_message",
    text: "deployed v2",
  });
  const history = new URL(sent[0].url);
  assert.equal(history.searchParams.get("oldest"), "1729987200.000000");
  assert.equal(history.searchParams.get("limit"), "20");
  // One users.info per person, author and mention alike.
  assert.deepEqual(calls().slice(1), ["users.info", "users.info"]);
  // Nothing of the raw payload beyond what is mapped.
  assert.ok(!r.text.includes("url_private") && !r.text.includes("blocks"));
});

test("reads refuse channels outside the key's limit before calling Slack", async () => {
  const { run, sent, audit } = setup({ channels: [GENERAL] });
  for (const [tool, args] of [
    ["slack_read_channel", { channel: SECRET, limit: 5 }],
    ["slack_read_thread", { channel: SECRET, thread_ts: "1730000002.000200", limit: 5 }],
    ["slack_get_permalink", { channel: SECRET, ts: "1730000002.000200" }],
    ["slack_post_message", { channel: SECRET, text: "hi" }],
    ["slack_add_reaction", { channel: SECRET, ts: "1730000002.000200", emoji: "eyes" }],
    ["slack_upload_text_file", { channel: SECRET, filename: "a.txt", content: "x" }],
    ["slack_delete_message", { channel: SECRET, ts: "1730000009.000900" }],
  ] as const) {
    const r = await run(tool, args);
    assert.equal(r.isError, true, tool);
    assert.match(r.text, /may not use that channel/);
  }
  assert.equal(sent.length, 0);
  assert.ok(audit.every((a) => a.decision === "denied"));
  // Inside the limit it goes through.
  assert.equal(
    (await run("slack_get_permalink", { channel: GENERAL, ts: "1730000002.000200" })).isError,
    false,
  );
});

test("bad channel ids and timestamps are refused before calling Slack", async () => {
  const { run, sent } = setup();
  for (const [tool, args] of [
    ["slack_read_channel", { channel: "general", limit: 5 }],
    ["slack_read_channel", { channel: "D0DIRECT", limit: 5 }],
    ["slack_read_channel", { channel: GENERAL, limit: 5, latest: "yesterday" }],
    ["slack_read_thread", { channel: GENERAL, thread_ts: "123", limit: 5 }],
    ["slack_post_message", { channel: GENERAL, text: "hi", thread_ts: "x" }],
    ["slack_upload_text_file", { channel: GENERAL, filename: "../a.txt", content: "x" }],
  ] as const)
    assert.equal((await run(tool, args)).isError, true, JSON.stringify(args));
  assert.equal(sent.length, 0);
});

test("slack_read_thread reads the thread and a channel the bot is not in gets Slack's reason", async () => {
  const { run, sent } = setup();
  const r = await run("slack_read_thread", {
    channel: GENERAL,
    thread_ts: "1730000002.000200",
    limit: 50,
  });
  assert.deepEqual(
    r.body.messages.map((m: { text: string; user: string }) => [m.user, m.text]),
    [
      ["ada", "ping @Bob Ross in #secret <see the doc (https://x.example)> @here"],
      ["Bob Ross", "hi"],
    ],
  );
  assert.equal(new URL(sent[0].url).searchParams.get("ts"), "1730000002.000200");
  const outside = await run("slack_read_channel", { channel: OTHER, limit: 5 });
  assert.equal(outside.isError, true);
  assert.match(outside.text, /not_in_channel: the app is not a member of that channel/);
});

test("slack_post_message posts as JSON, replies in a thread, and 403 carries the scope hint", async () => {
  const { run, sent, audit } = setup();
  const r = await run("slack_post_message", {
    channel: GENERAL,
    text: "Done",
    thread_ts: "1730000002.000200",
    also_send_to_channel: false,
  });
  assert.deepEqual(r.body, {
    channel: GENERAL,
    ts: "1730000009.000900",
    thread_ts: "1730000002.000200",
  });
  assert.equal(sent[0].method, "POST");
  assert.deepEqual(sent[0].body, {
    channel: GENERAL,
    text: "Done",
    thread_ts: "1730000002.000200",
    reply_broadcast: false,
  });
  assert.equal(audit.at(-1)!.decision, "allowed");

  const scoped = setup(
    {},
    { "chat.postMessage": { body: { ok: false, error: "missing_scope", needed: "chat:write" } } },
  );
  const refused = await scoped.run("slack_post_message", { channel: GENERAL, text: "x" });
  assert.equal(refused.isError, true);
  assert.equal(
    refused.text,
    "Slack answered missing_scope: the Slack app lacks a permission this needs (chat:write). Add the bot scope chat:write.",
  );
  assert.equal(scoped.audit.at(-1)!.decision, "failed");
});

test("slack_add_reaction strips colons and treats already_reacted as done", async () => {
  const { run, sent } = setup();
  assert.deepEqual(
    (
      await run("slack_add_reaction", {
        channel: GENERAL,
        ts: "1730000002.000200",
        emoji: ":eyes:",
      })
    ).body,
    { reacted: true, already: false },
  );
  assert.deepEqual(sent[0].body, {
    channel: GENERAL,
    timestamp: "1730000002.000200",
    name: "eyes",
  });
  const again = setup({}, { "reactions.add": { body: { ok: false, error: "already_reacted" } } });
  assert.deepEqual(
    (
      await again.run("slack_add_reaction", {
        channel: GENERAL,
        ts: "1730000002.000200",
        emoji: "eyes",
      })
    ).body,
    { reacted: true, already: true },
  );
  const bad = setup({}, { "reactions.add": { body: { ok: false, error: "invalid_name" } } });
  const r = await bad.run("slack_add_reaction", {
    channel: GENERAL,
    ts: "1730000002.000200",
    emoji: "nope",
  });
  assert.equal(r.isError, true);
  assert.match(r.text, /no emoji with that name/);
});

test("slack_upload_text_file runs Slack's three-step upload, the bytes only to files.slack.com", async () => {
  const { run, sent } = setup();
  const r = await run("slack_upload_text_file", {
    channel: GENERAL,
    filename: "report.csv",
    content: "héllo",
    comment: "Weekly report",
    thread_ts: "1730000002.000200",
  });
  assert.deepEqual(r.body, { file_id: "F1", channel: GENERAL, thread_ts: "1730000002.000200" });
  const [slot, bytes, complete] = sent;
  assert.equal(new URL(slot.url).searchParams.get("length"), "6");
  assert.equal(bytes.url, UPLOAD);
  assert.equal(bytes.auth, null, "the token never goes to the upload address");
  assert.equal(bytes.body, "héllo");
  assert.deepEqual(complete.body, {
    files: [{ id: "F1", title: "report.csv" }],
    channel_id: GENERAL,
    initial_comment: "Weekly report",
    thread_ts: "1730000002.000200",
  });

  const elsewhere = setup(
    {},
    {
      "files.getUploadURLExternal": {
        body: { ok: true, upload_url: "https://evil.example/u", file_id: "F2" },
      },
    },
  );
  const refused = await elsewhere.run("slack_upload_text_file", {
    channel: GENERAL,
    filename: "a.txt",
    content: "x",
  });
  assert.equal(refused.isError, true);
  assert.ok(!elsewhere.sent.some((s) => s.url.startsWith("https://evil.example")));

  const big = await run("slack_upload_text_file", {
    channel: GENERAL,
    filename: "a.txt",
    content: "x".repeat(1_000_001),
  });
  assert.match(big.text, /over 1 MB/);
});

test("slack_find_user_by_email answers a trimmed person, or found: false", async () => {
  const { run } = setup();
  assert.deepEqual((await run("slack_find_user_by_email", { email: "ada@acme.example" })).body, {
    found: true,
    id: "U1",
    name: "ada",
    real_name: "Ada Lovelace",
    title: "Engineer",
    time_zone: "Europe/London",
    is_bot: false,
    deactivated: false,
  });
  assert.deepEqual((await run("slack_find_user_by_email", { email: "x@acme.example" })).body, {
    found: false,
  });
});

test("slack_delete_message deletes the app's own message and passes on Slack's refusal", async () => {
  const { run, sent } = setup();
  assert.deepEqual(
    (await run("slack_delete_message", { channel: GENERAL, ts: "1730000009.000900" })).body,
    { deleted: true, channel: GENERAL, ts: "1730000009.000900" },
  );
  assert.deepEqual(sent[0].body, { channel: GENERAL, ts: "1730000009.000900" });
  const r = await run("slack_delete_message", { channel: GENERAL, ts: "1730000002.000200" });
  assert.equal(r.isError, true);
  assert.match(r.text, /can delete only its own messages/);
});

test("a rate limit is waited out when short, and answered with when to retry when long", async () => {
  let n = 0;
  const short = setup(
    {},
    {
      "chat.getPermalink": () =>
        ++n < 3
          ? { status: 429, headers: { "retry-after": "2" }, body: "" }
          : { body: { ok: true, permalink: "https://acme.slack.com/archives/C/p1" } },
    },
  );
  const r = await short.run("slack_get_permalink", { channel: GENERAL, ts: "1730000002.000200" });
  assert.equal(r.isError, false);
  assert.deepEqual(short.waits, [2, 2]);

  const long = setup(
    {},
    { "chat.getPermalink": { status: 429, headers: { "retry-after": "60" }, body: "" } },
  );
  const refused = await long.run("slack_get_permalink", {
    channel: GENERAL,
    ts: "1730000002.000200",
  });
  assert.equal(refused.isError, true);
  assert.match(refused.text, /Try again in 60 seconds/);
  assert.deepEqual(long.waits, []);
  assert.equal(long.sent.length, 1);
});

test("ok: false answers become statuses the service understands", async () => {
  const cases: [string, number][] = [
    ["invalid_auth", 401],
    ["missing_scope", 403],
    ["channel_not_found", 404],
    ["not_in_channel", 400],
    ["something_new", 400],
  ];
  for (const [error, status] of cases) {
    const { vendor } = setup({}, { "chat.getPermalink": { body: { ok: false, error } } });
    const res = await vendor.call(connection, { method: "GET", path: "/chat.getPermalink" });
    assert.equal(res.status, status, error);
  }
});

test("connect takes a bot token, proves it with auth.test and names the workspace", async () => {
  const { vendor, sent } = setup();
  const conn = await vendor.connect({ bot_token: ` ${TOKEN} ` });
  assert.deepEqual(conn, {
    config: { teamId: "T1", team: "Acme", url: "https://acme.slack.com/", botUserId: "UBOT" },
    secret: { bot_token: TOKEN },
    accountLabel: "Acme (acme.slack.com)",
  });
  assert.equal(sent[0].auth, `Bearer ${TOKEN}`);
  await assert.rejects(vendor.connect({ bot_token: "xoxp-user-token" }), /starts with xoxb-/);
  const refused = setup({}, { "auth.test": { body: { ok: false, error: "invalid_auth" } } });
  await assert.rejects(refused.vendor.connect({ bot_token: TOKEN }), /did not accept the token/);
});

test("the workspace check warns about missing scopes and discover lists the bot's channels", async () => {
  const { vendor } = setup();
  const checks = vendor.checks!;
  assert.deepEqual(
    await checks.workspace(connection, {
      expected: ["channels:read", "chat:write", "files:write"],
    }),
    {
      values: ["Acme (acme.slack.com)"],
      warning:
        "The Slack app is missing these bot scopes: chat:write, files:write. Add them under OAuth & Permissions, reinstall the app to the workspace, then check again.",
    },
  );
  assert.deepEqual(await checks.workspace(connection, { expected: ["channels:read"] }), {
    values: ["Acme (acme.slack.com)"],
  });
  assert.deepEqual((await checks.channels(connection, { expected: [] })).values, [
    GENERAL,
    "C0RANDOM",
    SECRET,
  ]);
});

test("plainText leaves plain text alone and resolves what it can", () => {
  assert.equal(plainText("a & b", new Map()), "a & b");
  assert.equal(plainText("<@U9> <@U9|old> <#C1>", new Map()), "@U9 @old #C1");
  assert.equal(
    plainText("<!subteam^S1|@oncall> <mailto:a@b.c|a@b.c>", new Map()),
    "@oncall a@b.c (mailto:a@b.c)",
  );
});
