import assert from "node:assert/strict";
import { randomUUID } from "node:crypto";
import { after, before, test } from "node:test";
import { z } from "zod";
import { schemas } from "../src/api.ts";
import { createApp } from "../src/app.ts";
import { migrate } from "../src/migrate.ts";
import { CredentialKey } from "../src/secrets.ts";
import { connectDatabase, createStore } from "../src/store.ts";
import { hashReplyToken } from "../src/teams/activity.ts";
import { createTeamsAdmin } from "../src/teams/admin.ts";
import type { BotClient } from "../src/teams/bot.ts";
import type { BotAuth } from "../src/teams/jwt.ts";
import type { RunInput, TriggerRunner } from "../src/teams/runner.ts";
import { botAad, createTeamsRuntime } from "../src/teams/runtime.ts";
import { createTeamsStore } from "../src/teams/store.ts";
import { ADMIN_TOKEN, teamsConnector } from "./helpers.ts";
import { teamsTools, type TeamsToolContext } from "../../../connectors/teams/tools.ts";

// The Teams runtime and trigger admin with fake runners, against the scratch database
// (TEST_DATABASE_URL, see store.test.ts). Skipped without one.
const url = process.env.TEST_DATABASE_URL;
const sql = url ? connectDatabase(url) : null;
const SCHEMA = "agent_connect";
before(async () => {
  if (sql) await migrate(sql, SCHEMA, () => {});
});
after(() => sql?.end());

const key = new CredentialKey(Buffer.alloc(32, 7).toString("base64"));
const TENANT = "33333333-3333-3333-3333-333333333333";
const PERSON = "44444444-4444-4444-4444-444444444444";
const SERVICE_URL = "https://smba.trafficmanager.net/amer/";
const ok = { status: 200, data: { value: [] }, headers: {} };
const auth = { token: async () => () => {} } as unknown as BotAuth;
type Sent = { method: string; path: string; body?: any };
// A Bot Connector that records what was sent. A new channel post answers its thread and
// activity id, any other post its activity id.
function fakeBots(sent: Sent[]) {
  return {
    connector: async (_bot: unknown, _url: string, req: Sent) => {
      sent.push(req);
      if (req.method === "POST" && req.path === "/v3/conversations")
        return {
          status: 201,
          data: { id: "19:general@thread.tacv2;messageid=900", activityId: "900" },
        };
      return { status: 200, data: { id: "900", value: [] } };
    },
    graph: async () => ok,
  } as unknown as BotClient;
}

// A runner whose start can reply through teams_reply's own calls before it returns, as a
// webhook receiver answering inside its handler does.
function fakeRunner(id: string, onStart?: (input: RunInput) => Promise<void>) {
  const calls: string[] = [];
  const runner: TriggerRunner = {
    id,
    settings: z.strictObject({}),
    unready: async () => null,
    verify: async () => {},
    async start(input) {
      calls.push(`start ${input.runId}`);
      await onStart?.(input);
      return `${id}:${input.runId}`;
    },
    async continue(session) {
      calls.push(`continue ${session}`);
    },
    async state(run) {
      return { status: run.replyCount > 0 ? "idle" : "running", costCents: 0, atBudget: false };
    },
    async stop(session) {
      calls.push(`stop ${session}`);
    },
  };
  return { runner, calls };
}

async function setup(runners: TriggerRunner[]) {
  const sent: Sent[] = [];
  const bots = fakeBots(sent);
  const store = createStore(sql!, SCHEMA);
  const teams = createTeamsStore(sql!, SCHEMA);
  const org = randomUUID();
  await store.ensureOrganization(org);
  const [{ id: agentId }] = await sql!`
    insert into agent_connect.agents (organization_id, name, key_hash, status)
    values (${org}, 'Triage', ${randomUUID().replaceAll("-", "").repeat(2)}, 'active')
    returning id::text`;
  await sql!`insert into agent_connect.connections
               (organization_id, connector_id, credential, credential_key_id, status, capabilities)
             values (${org}, 'teams', ${Buffer.from("x")}, 'k', 'connected', ${sql!.array(["respond"])})`;
  await sql!`insert into agent_connect.grants (agent_id, connector_id, capabilities)
             values (${agentId}, 'teams', ${sql!.array(["respond"])})`;
  const botId = randomUUID();
  const botAppId = randomUUID();
  const sealed = key.seal({ private_key: "pk" }, botAad(org, botId));
  await sql!`insert into agent_connect.teams_bots
               (id, organization_id, agent_id, bot_app_id, tenant_id, display_name, status, config,
                credential, credential_key_id)
             values (${botId}, ${org}, ${agentId}, ${botAppId}, ${TENANT}, 'Triage', 'connected',
                     ${sql!.json({ certificate: "cert" })}, ${sealed.ciphertext}, ${sealed.keyId})`;
  const byId = new Map(runners.map((r) => [r.id, r]));
  const publicUrl = "https://svc.example/agent-connect";
  const runtime = createTeamsRuntime({
    store,
    teams,
    key,
    auth,
    bots,
    runners: byId,
    connector: teamsConnector,
    publicUrl,
  });
  const admin = createTeamsAdmin({
    store,
    teams,
    key,
    bots,
    claude: () => {
      throw new Error("no Claude here");
    },
    runners: byId,
    publicUrl,
    microsoft: {} as never,
  });
  const setRunner = (runner: string) =>
    admin.setTrigger(org, agentId, {
      enabled: true,
      runner,
      budgetCents: 50,
      allowedPeople: [PERSON],
      by: "admin",
    });
  let n = 0;
  // A direct message from PERSON, through the bot's inbound route; `message` also waits for it.
  async function send(text: string) {
    const activity = {
      type: "message",
      id: `act-${++n}`,
      serviceUrl: SERVICE_URL,
      channelId: "msteams",
      recipient: { id: `28:${botAppId}` },
      from: { id: "29:person", aadObjectId: PERSON, name: "Ann" },
      conversation: { id: "a:chat", conversationType: "personal" },
      channelData: { tenant: { id: TENANT } },
      text,
    };
    const res = await runtime.inbound(
      botAppId,
      new Request(`${publicUrl}/teams/bots/${botAppId}/messages`, {
        method: "POST",
        headers: { authorization: "Bearer x", "content-type": "application/json" },
        body: JSON.stringify(activity),
      }),
    );
    assert.equal(res.status, 200);
  }
  async function message(text: string) {
    await send(text);
    await runtime.idle();
  }
  const runs = async () => teams.recentRuns(org, agentId, 10);
  const app = createApp({
    service: {} as never,
    store,
    adminToken: ADMIN_TOKEN,
    teams: { admin, runtime },
    validateResponses: true,
  });
  return {
    org,
    agentId,
    admin,
    teams,
    runtime,
    runners: byId,
    app,
    botId,
    setRunner,
    send,
    message,
    runs,
    sent,
    botAppId,
    store,
  };
}

test(
  "a runner can reply through teams_reply before its start returns",
  { skip: !sql },
  async () => {
    const replied: boolean[] = [];
    let reply: (input: RunInput) => Promise<void> = async () => {};
    const { runner } = fakeRunner("fake-hook", (input) => reply(input));
    const t = await setup([runner]);
    const bot = (await t.teams.botById(t.botId))!;
    reply = async (input) => {
      const access = t.runtime.access(bot);
      const run = await access.runForReply(hashReplyToken(input.replyToken));
      replied.push(Boolean(run) && (await access.claimReply(run!.id)));
    };
    await t.setRunner("fake-hook");
    await t.message("hello");
    assert.deepEqual(replied, [true]);
    const [run] = await t.runs();
    assert.equal(run.state, "running");
    assert.equal(run.replyCount, 1);
    assert.equal(run.sessionId, `fake-hook:${run.id}`);
    await t.runtime.watch();
    assert.equal((await t.runs())[0].state, "idle");
  },
);

test(
  "a session stays with the runner that started it when the trigger changes runner",
  { skip: !sql },
  async () => {
    const a = fakeRunner("fake-a");
    const b = fakeRunner("fake-b");
    const t = await setup([a.runner, b.runner]);

    // An answered, idle session of runner a is not continued by runner b.
    await t.setRunner("fake-a");
    await t.message("one");
    const [first] = await t.runs();
    await sql!`update agent_connect.teams_runs set reply_count = 1 where id = ${first.id}`;
    await t.runtime.watch();
    assert.equal((await t.runs())[0].state, "idle");
    await t.setRunner("fake-b");
    await t.message("two");
    const [second] = await t.runs();
    assert.notEqual(second.id, first.id);
    assert.equal(second.runner, "fake-b");
    assert.deepEqual(a.calls, [`start ${first.id}`]);
    assert.deepEqual(b.calls, [`start ${second.id}`]);

    // A live session of runner b is stopped by runner b when the trigger moves back to a.
    assert.equal(second.state, "running");
    await t.setRunner("fake-a");
    assert.deepEqual(b.calls, [`start ${second.id}`, `stop fake-b:${second.id}`]);
    const stopped = (await t.runs()).find((r) => r.id === second.id)!;
    assert.equal(stopped.state, "failed");
    assert.deepEqual(a.calls, [`start ${first.id}`]);
  },
);

test(
  "a run whose runner is no longer offered fails saying so, not as a removed trigger",
  { skip: !sql },
  async () => {
    const a = fakeRunner("fake-a");
    const t = await setup([a.runner]);
    await t.setRunner("fake-a");
    await t.message("one");
    t.runners.delete("fake-a");
    await t.runtime.watch();
    const [run] = await t.runs();
    assert.equal(run.state, "failed");
    assert.equal(run.error, "Runner fake-a is not offered by this deployment");
  },
);

test(
  "the old name of madeByService is answered on the deprecated prefix only",
  { skip: !sql },
  async () => {
    const t = await setup([]);
    // No certificate to describe: the fake one is not PEM.
    await sql!`update agent_connect.teams_bots set config = '{}' where id = ${t.botId}`;
    const bots = async (prefix: string) => {
      const res = await t.app.fetch(
        new Request(`http://svc/agent-connect/admin${prefix}/orgs/${t.org}/teams/bots`, {
          headers: { authorization: `Bearer ${ADMIN_TOKEN}` },
        }),
      );
      assert.equal(res.status, 200);
      return (await res.json()) as Record<string, unknown>[];
    };
    const [v1] = await bots("/v1");
    assert.equal(v1.madeByService, false);
    assert.ok(!("madeByArbium" in v1));
    const [legacy] = await bots("");
    assert.deepEqual(legacy, { ...v1, madeByArbium: false });

    // Also where the bot is nested in the answer. Adding to Teams needs a tenant, so the
    // admin answers with the bot it has.
    const record = Object.fromEntries(
      Object.keys(schemas.TeamsBotRecord.shape).map((k) => [k, v1[k]]),
    );
    t.admin.addToTeams = async () => ({ steps: [], bot: record as never });
    const add = async (prefix: string) => {
      const res = await t.app.fetch(
        new Request(
          `http://svc/agent-connect/admin${prefix}/orgs/${t.org}/agents/${t.agentId}/teams-bot/add`,
          {
            method: "POST",
            headers: { authorization: `Bearer ${ADMIN_TOKEN}`, "content-type": "application/json" },
            body: JSON.stringify({ displayName: "Triage", by: "admin" }),
          },
        ),
      );
      assert.equal(res.status, 200);
      return (await res.json()) as { bot: Record<string, unknown> };
    };
    assert.ok(!("madeByArbium" in (await add("/v1")).bot));
    assert.deepEqual((await add("")).bot, { ...record, madeByArbium: false });
  },
);

test("stop waits for the event in hand before returning", { skip: !sql }, async () => {
  let started!: () => void;
  const running = new Promise<void>((r) => (started = r));
  let release!: () => void;
  const held = new Promise<void>((r) => (release = r));
  const { runner } = fakeRunner("fake-held", async () => {
    started();
    await held;
  });
  const t = await setup([runner]);
  await t.setRunner("fake-held");
  await t.send("hello");
  await running;
  let stopped = false;
  const stopping = t.runtime.stop().then(() => (stopped = true));
  await new Promise((r) => setTimeout(r, 50));
  assert.equal(stopped, false);
  release();
  await stopping;
  const [event] =
    await sql!`select state from agent_connect.teams_events where bot_id = ${t.botId}`;
  assert.notEqual(event.state, "processing");
});

// ---- approval cards ----

const OTHER = "66666666-6666-6666-6666-666666666666";
const CHAT = "19:ops@thread.v2";

// The agent's Teams tools on this bot, the directory answering for PERSON (Ann) and OTHER.
async function approvals(t: Awaited<ReturnType<typeof setup>>) {
  await t.teams.upsertConversation({
    botId: t.botId,
    conversationId: CHAT,
    kind: "groupChat",
    serviceUrl: SERVICE_URL,
    teamId: null,
    teamGroupId: null,
    name: "Ops",
    userAadId: null,
  });
  const bot = (await t.teams.botById(t.botId))!;
  const lookups: string[] = [];
  const handlers = new Map<string, (input: unknown) => Promise<any>>();
  const ctx: TeamsToolContext = {
    limits: {},
    config: {},
    call: async (req) => {
      lookups.push(req.path);
      if (req.path.startsWith(`/v1.0/users/${encodeURIComponent("ann@contoso.example")}?`))
        return { status: 200, data: { id: PERSON.toUpperCase(), displayName: "Ann" } };
      return { status: 404, data: { error: { message: "Resource not found" } } };
    },
    teams: t.runtime.access(bot),
    audit: (entry) =>
      t.store.audit({
        organizationId: t.org,
        agentId: t.agentId,
        agentName: "Triage",
        connectorId: "teams",
        ...entry,
      }),
  };
  teamsTools.register(
    { registerTool: (name: string, _d: unknown, h: any) => handlers.set(name, h) } as never,
    ctx,
    new Set(teamsTools.names),
  );
  const call = async (name: string, input: unknown) => {
    const r = await handlers.get(name)!(input);
    return { isError: Boolean(r.isError), text: r.content[0].text as string };
  };
  async function request() {
    const r = await call("teams_request_approval", {
      chat: CHAT,
      approver: "ann@contoso.example",
      title: "Laptop for Bob",
      text: "Approve a replacement laptop?",
      facts: [{ name: "Cost", value: "1,200 USD" }],
    });
    assert.equal(r.isError, false, r.text);
    return JSON.parse(r.text) as { approval_id: string; message_id: string; expires: string };
  }
  const read = async (approvalId: string) =>
    JSON.parse((await call("teams_get_approval", { approval_id: approvalId })).text);
  let n = 0;
  // A click on the card's button, as Teams sends it to a bot's address.
  async function click(
    rowId: string,
    opts: {
      from?: string;
      verb?: string;
      comment?: string;
      botAppId?: string;
      conversation?: { id: string; conversationType: string };
    } = {},
  ) {
    const botAppId = opts.botAppId ?? t.botAppId;
    const activity = {
      type: "invoke",
      name: "adaptiveCard/action",
      id: `inv-${++n}`,
      serviceUrl: SERVICE_URL,
      channelId: "msteams",
      recipient: { id: `28:${botAppId}` },
      from: { id: "29:clicker", aadObjectId: opts.from ?? PERSON, name: "Ann Example" },
      conversation: opts.conversation ?? { id: CHAT, conversationType: "groupChat" },
      channelData: { tenant: { id: TENANT } },
      replyToId: "900",
      value: {
        action: {
          type: "Action.Execute",
          verb: opts.verb ?? "approve",
          data: { approval: rowId, ...(opts.comment ? { comment: opts.comment } : {}) },
        },
        trigger: "manual",
      },
    };
    const res = await t.runtime.inbound(
      botAppId,
      new Request(`https://svc.example/agent-connect/teams/bots/${botAppId}/messages`, {
        method: "POST",
        headers: { authorization: "Bearer x", "content-type": "application/json" },
        body: JSON.stringify(activity),
      }),
    );
    assert.equal(res.status, 200);
    return (await res.json()) as { statusCode: number; type: string; value: any };
  }
  const rowId = async () => {
    const [row] = await sql!`select id::text from agent_connect.approvals
                              where bot_id = ${t.botId} order by created_at desc limit 1`;
    return row.id as string;
  };
  const audits = async () =>
    (await sql!`select tool, decision, reason from agent_connect.audit
                 where agent_id = ${t.agentId} and tool like 'teams%approval' order by id`) as unknown as {
      tool: string;
      decision: string;
      reason: string;
    }[];
  return { call, request, read, click, rowId, audits, lookups };
}

const cardText = (card: any) => JSON.stringify(card);

test("an approval card is posted as the bot and read back pending", { skip: !sql }, async () => {
  const t = await setup([]);
  const a = await approvals(t);
  const asked = await a.request();
  assert.match(asked.approval_id, /^appr_[A-Za-z0-9_-]{43}$/);
  assert.equal(asked.message_id, "900");
  const days = (Date.parse(asked.expires) - Date.now()) / 86400_000;
  assert.ok(days > 6.9 && days <= 7, `expires in ${days} days`);
  // Only the hash of what the agent holds is stored.
  const [stored] = await sql!`select token_hash, approver_aad_id::text, thread_id, activity_id
                                from agent_connect.approvals where bot_id = ${t.botId}`;
  assert.notEqual(stored.token_hash, asked.approval_id);
  assert.equal(stored.approver_aad_id, PERSON);
  assert.deepEqual([stored.thread_id, stored.activity_id], [CHAT, "900"]);
  const post = t.sent.find((s) => s.method === "POST")!;
  assert.equal(post.path, `/v3/conversations/${encodeURIComponent(CHAT)}/activities`);
  const [attachment] = post.body.attachments;
  assert.equal(attachment.contentType, "application/vnd.microsoft.card.adaptive");
  assert.equal(post.body.text, undefined);
  const actions = attachment.content.body.find((b: any) => b.type === "ActionSet").actions;
  assert.deepEqual(
    actions.map((x: any) => [x.type, x.verb, x.data.approval]),
    [
      ["Action.Execute", "approve", await a.rowId()],
      ["Action.Execute", "reject", await a.rowId()],
    ],
  );
  assert.match(cardText(attachment.content), /1,200 USD/);
  const status = await a.read(asked.approval_id);
  assert.equal(status.status, "pending");
  assert.deepEqual(status.approver, { id: PERSON, name: "Ann" });
  assert.equal(status.decided_by, null);
  const rows = await a.audits();
  assert.deepEqual(
    rows.map((r) => [r.tool, r.decision]),
    [
      ["teams_request_approval", "allowed"],
      ["teams_get_approval", "allowed"],
    ],
  );
});

test("the approver's click is recorded once and the card shows it", { skip: !sql }, async () => {
  const t = await setup([]);
  const a = await approvals(t);
  const asked = await a.request();
  const id = await a.rowId();
  const answer = await a.click(id, { verb: "reject", comment: "Too expensive" });
  assert.equal(answer.statusCode, 200);
  assert.equal(answer.type, "application/vnd.microsoft.card.adaptive");
  assert.match(cardText(answer.value), /Rejected by Ann Example/);
  assert.ok(!cardText(answer.value).includes("Action.Execute"));
  // Everyone's card is edited in place.
  const edit = t.sent.find((s) => s.method === "PUT")!;
  assert.equal(edit.path, `/v3/conversations/${encodeURIComponent(CHAT)}/activities/900`);
  assert.match(cardText(edit.body), /Rejected by Ann Example/);
  const status = await a.read(asked.approval_id);
  assert.equal(status.status, "rejected");
  assert.deepEqual(status.decided_by, { id: PERSON, name: "Ann Example" });
  assert.equal(status.comment, "Too expensive");
  assert.ok(status.decided_at);

  // A second click changes nothing.
  const again = await a.click(id, { verb: "approve" });
  assert.match(cardText(again.value), /Rejected by Ann Example/);
  assert.equal((await a.read(asked.approval_id)).status, "rejected");
  const clicks = (await a.audits()).filter((r) => r.tool === "teams:approval");
  assert.deepEqual(
    clicks.map((r) => r.decision),
    ["allowed", "denied"],
  );
  assert.match(clicks[1].reason, /already rejected/);
});

test(
  "a click by anyone but the approver records nothing but an audit row",
  { skip: !sql },
  async () => {
    const t = await setup([]);
    const a = await approvals(t);
    const asked = await a.request();
    const answer = await a.click(await a.rowId(), { from: OTHER });
    assert.equal(answer.type, "application/vnd.microsoft.activity.message");
    assert.equal(answer.value, "Only Ann can answer this request.");
    assert.ok(!t.sent.some((s) => s.method === "PUT"));
    assert.equal((await a.read(asked.approval_id)).status, "pending");
    const [click] = (await a.audits()).filter((r) => r.tool === "teams:approval");
    assert.equal(click.decision, "denied");
    assert.match(click.reason, new RegExp(`clicked by ${OTHER}, not the approver`));
  },
);

test("an expired approval takes no answer", { skip: !sql }, async () => {
  const t = await setup([]);
  const a = await approvals(t);
  const asked = await a.request();
  const id = await a.rowId();
  await sql!`update agent_connect.approvals set expires_at = now() - interval '1 minute'
              where id = ${id}`;
  const answer = await a.click(id);
  assert.match(cardText(answer.value), /Expired/);
  assert.match(cardText(t.sent.find((s) => s.method === "PUT")!.body), /Expired/);
  const status = await a.read(asked.approval_id);
  assert.equal(status.status, "expired");
  assert.equal(status.decided_by, null);
  assert.equal((await a.audits()).find((r) => r.tool === "teams:approval")!.decision, "denied");
});

test("another bot's approval is not answered through this bot", { skip: !sql }, async () => {
  const t = await setup([]);
  const other = await setup([]);
  const a = await approvals(t);
  const asked = await a.request();
  const answer = await a.click(await a.rowId(), { botAppId: other.botAppId });
  assert.equal(answer.type, "application/vnd.microsoft.activity.message");
  assert.equal(answer.value, "This request is not known to this agent.");
  assert.equal((await a.read(asked.approval_id)).status, "pending");
  // Audited under the bot that was clicked.
  const [row] = await sql!`select decision from agent_connect.audit
                            where agent_id = ${other.agentId} and tool = 'teams:approval'`;
  assert.equal(row.decision, "denied");
  // And the other agent cannot read it either.
  const b = await approvals(other);
  const r = await b.call("teams_get_approval", { approval_id: asked.approval_id });
  assert.equal(r.isError, true);
});

test(
  "a click is answered with real-time responses off and starts no run",
  { skip: !sql },
  async () => {
    const t = await setup([]);
    assert.equal(await t.teams.trigger(t.agentId), null);
    const a = await approvals(t);
    const asked = await a.request();
    await a.click(await a.rowId());
    await t.runtime.idle();
    assert.equal((await a.read(asked.approval_id)).status, "approved");
    const [events] = await sql!`select count(*)::int as n from agent_connect.teams_events
                               where bot_id = ${t.botId}`;
    assert.equal(events.n, 0);
    assert.equal((await t.runs()).length, 0);
  },
);

test(
  "an approver the directory does not know is refused before anything is posted",
  { skip: !sql },
  async () => {
    const t = await setup([]);
    const a = await approvals(t);
    const r = await a.call("teams_request_approval", {
      chat: CHAT,
      approver: "nobody@contoso.example",
      title: "x",
      text: "y",
    });
    assert.equal(r.isError, true);
    assert.match(r.text, /No one in this organization signs in as nobody@contoso.example/);
    assert.ok(!t.sent.some((s) => s.method === "POST"));
    const [n] = await sql!`select count(*)::int as n from agent_connect.approvals
                          where bot_id = ${t.botId}`;
    assert.equal(n.n, 0);
  },
);

test(
  "a click from another conversation than the card's does not decide it",
  { skip: !sql },
  async () => {
    const t = await setup([]);
    const a = await approvals(t);
    const asked = await a.request();
    const id = await a.rowId();
    // The approver's own click, but sent from a 1:1 chat with the bot (a forwarded card or a
    // crafted invoke).
    const answer = await a.click(id, {
      conversation: { id: "a:ann-1on1", conversationType: "personal" },
    });
    assert.equal(answer.type, "application/vnd.microsoft.activity.message");
    assert.equal(answer.value, "This request is not known to this agent.");
    assert.ok(!t.sent.some((s) => s.method === "PUT"));
    assert.equal((await a.read(asked.approval_id)).status, "pending");
    const [click] = (await a.audits()).filter((r) => r.tool === "teams:approval");
    assert.equal(click.decision, "denied");
    assert.match(click.reason, /in a:ann-1on1: no such approval for this bot there/);
    // From the card's own chat it still counts.
    await a.click(id);
    assert.equal((await a.read(asked.approval_id)).status, "approved");
  },
);

// decideApproval is the one statement that writes a decision; each of its conditions is
// checked here on its own, without the runtime's checks in front of it.
test("the store records a decision only when every condition holds", { skip: !sql }, async () => {
  const t = await setup([]);
  const other = await setup([]);
  const make = () =>
    t.teams.createApproval({
      agentId: t.agentId,
      botId: t.botId,
      tokenHash: randomUUID().replaceAll("-", "").repeat(2),
      conversationId: CHAT,
      serviceUrl: SERVICE_URL,
      approverAadId: PERSON,
      approverName: "Ann",
      title: "x",
      body: "y",
      facts: [],
      days: 7,
    });
  const by = (aad: string) => ({ status: "approved" as const, byAadId: aad, byName: "Ann" });
  const status = async (id: string) =>
    (await sql!`select status from agent_connect.approvals where id = ${id}`)[0].status;

  // Not the approver.
  const a = await make();
  assert.equal(await t.teams.decideApproval(t.botId, a.id, CHAT, by(OTHER), null), null);
  assert.equal(await status(a.id), "pending");
  // Expired.
  const b = await make();
  await sql!`update agent_connect.approvals set expires_at = now() - interval '1 second'
              where id = ${b.id}`;
  assert.equal(await t.teams.decideApproval(t.botId, b.id, CHAT, by(PERSON), null), null);
  assert.equal(await status(b.id), "pending");
  // Another conversation.
  const c = await make();
  assert.equal(await t.teams.decideApproval(t.botId, c.id, "a:other", by(PERSON), null), null);
  assert.equal(await status(c.id), "pending");
  // Another bot.
  const d = await make();
  assert.equal(await t.teams.decideApproval(other.botId, d.id, CHAT, by(PERSON), null), null);
  assert.equal(await status(d.id), "pending");
  // All of them hold: recorded, once.
  const e = await make();
  const decided = await t.teams.decideApproval(t.botId, e.id, CHAT, by(PERSON), "ok");
  assert.equal(decided?.status, "approved");
  assert.equal(decided?.decidedByAadId, PERSON);
  const again = { status: "rejected" as const, byAadId: PERSON, byName: "Ann" };
  assert.equal(await t.teams.decideApproval(t.botId, e.id, CHAT, again, null), null);
  assert.equal(await status(e.id), "approved");
  // The lookup a click starts with keeps to the bot and the conversation too.
  assert.equal(await t.teams.approvalForBot(t.botId, a.id, "a:other"), null);
  assert.equal(await t.teams.approvalForBot(other.botId, a.id, CHAT), null);
  assert.equal((await t.teams.approvalForBot(t.botId, a.id, CHAT))?.id, a.id);
});

test("two clicks at once record one answer", { skip: !sql }, async () => {
  const t = await setup([]);
  const a = await approvals(t);
  const asked = await a.request();
  const id = await a.rowId();
  const answers = await Promise.all([
    a.click(id, { verb: "approve" }),
    a.click(id, { verb: "reject" }),
    a.click(id, { verb: "approve" }),
  ]);
  const status = (await a.read(asked.approval_id)).status;
  assert.ok(status === "approved" || status === "rejected", status);
  const shown = new RegExp(status === "approved" ? "Approved by" : "Rejected by");
  for (const answer of answers) assert.match(cardText(answer.value), shown);
  const clicks = (await a.audits()).filter((r) => r.tool === "teams:approval");
  assert.deepEqual(clicks.map((r) => r.decision).sort(), ["allowed", "denied", "denied"]);
  // Only the one decision edited the card.
  assert.equal(t.sent.filter((s) => s.method === "PUT").length, 1);
});

test(
  "a click that lands before the card's id is saved is edited in once it is",
  { skip: !sql },
  async () => {
    const t = await setup([]);
    const a = await approvals(t);
    await a.request();
    const id = await a.rowId();
    // As if the click came between the post and saving the card's id.
    await sql!`update agent_connect.approvals set thread_id = null, activity_id = null
                where id = ${id}`;
    await a.click(id);
    assert.ok(!t.sent.some((s) => s.method === "PUT"));
    const access = t.runtime.access((await t.teams.botById(t.botId))!);
    await access.approvalPosted(id, CHAT, "900");
    const edit = t.sent.find((s) => s.method === "PUT")!;
    assert.equal(edit.path, `/v3/conversations/${encodeURIComponent(CHAT)}/activities/900`);
    assert.match(cardText(edit.body), /Approved by Ann Example/);
  },
);
