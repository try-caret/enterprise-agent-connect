import assert from "node:assert/strict";
import { randomBytes, randomUUID } from "node:crypto";
import { readFileSync } from "node:fs";
import { after, before, test } from "node:test";
import { z } from "zod";
import { parseConnector } from "../../connector-kit/src/connector.ts";
import type { VendorRequest, VendorResponse } from "../../connector-kit/src/types.ts";
import { createEventsRuntime, subscriptionAad } from "../src/events/runtime.ts";
import { createEventsStore, type EventsStore } from "../src/events/store.ts";
import { migrate } from "../src/migrate.ts";
import { CredentialKey } from "../src/secrets.ts";
import { connectDatabase, createStore } from "../src/store.ts";
import { hashReplyToken } from "../src/teams/activity.ts";
import type { BotClient } from "../src/teams/bot.ts";
import type { BotAuth } from "../src/teams/jwt.ts";
import type { Caps, EventRunInput, RunInput, TriggerRunner } from "../src/teams/runner.ts";
import { botAad, createTeamsRuntime } from "../src/teams/runtime.ts";
import { createTeamsStore } from "../src/teams/store.ts";
import { teamsConnector } from "./helpers.ts";

// The service as two processes on one database: each test drives two runtimes against the
// scratch database (TEST_DATABASE_URL), each process's step held open where the other's matters.
// In a schema of its own: the other files' runtimes claim whatever is due in theirs.
const url = process.env.TEST_DATABASE_URL;
const sql = url ? connectDatabase(url) : null;
const SCHEMA = `two_replicas_${randomBytes(4).toString("hex")}`;
before(async () => {
  if (sql) await migrate(sql, SCHEMA, () => {});
});
after(async () => {
  await sql?.begin(async (tx) => {
    await tx`set local client_min_messages = warning`;
    await tx.unsafe(`drop schema ${SCHEMA} cascade`);
  });
  await sql?.end();
});

const key = new CredentialKey(Buffer.alloc(32, 7).toString("base64"));
const PUBLIC = "https://svc.example/agent-connect";
const MAILBOX = "helper@example.com";
const SECRET = "s".repeat(40);
const name = (table: string) => sql!(`${SCHEMA}.${table}`);
const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));
const lapseLeases = () => sql!`update ${name("leases")} set expires_at = now() - interval '1 second'`;

// Holds whoever calls `pass` until `open`; `entered` says one has arrived.
function gate() {
  let open!: () => void;
  let enter!: () => void;
  const opened = new Promise<void>((r) => (open = r));
  const entered = new Promise<void>((r) => (enter = r));
  return { open, entered, pass: () => (enter(), opened) };
}

// Holds each caller until `n` have called, or for 200 ms: one holding a lock the others want
// cannot be joined by them.
function barrier(n: number) {
  let arrived = 0;
  let open!: () => void;
  const all = new Promise<void>((r) => (open = r));
  return async () => {
    if (++arrived >= n) open();
    await Promise.race([all, sleep(200)]);
  };
}

const running = { status: "running" as const, costCents: 0, atBudget: false };

function fakeRunner(over: Partial<TriggerRunner> = {}) {
  const started: (RunInput | EventRunInput)[] = [];
  const runner: TriggerRunner = {
    id: "fake",
    settings: z.strictObject({}),
    unready: async () => null,
    verify: async () => {},
    async start(input) {
      started.push(input);
      return `s-${input.runId}`;
    },
    async startEvent(input) {
      started.push(input);
      return `s-${input.runId}`;
    },
    continue: async () => {},
    state: async () => running,
    stop: async () => {},
    ...over,
  };
  return { runner, started };
}

// ---- app events ----

const app = parseConnector(
  {
    ...JSON.parse(
      readFileSync(
        new URL("../../../connectors/outlook-mail/connector.json", import.meta.url),
        "utf8",
      ),
    ),
    id: "app",
    name: "App",
    events: {
      types: [{ id: "received", label: "Received" }],
      fields: { mailbox: "mailbox", id: "message.id", from: "message.from" },
      limits: { mailbox: "mailbox" },
      poll: {
        request: { method: "GET", path: "/items" },
        items: "items",
        cursor: "message.updated",
        deliveryId: ["message.id"],
        types: [{ type: "received" }],
      },
      graph: {
        per: "mailbox",
        resource: "users/{value}/messages",
        changeType: "created",
        type: "received",
        maxMinutes: 10070,
        read: { method: "GET", path: "/v1.0/users/{value}/messages/{id}" },
        delta: { method: "GET", path: "/v1.0/users/{value}/delta?since={since}" },
        wrap: "message",
        deliveryId: "message.id",
        received: "message.received",
        own: ["from"],
        only: {},
      },
    },
  },
  "app",
);

const item = (id: string) => ({
  mailbox: MAILBOX,
  message: { id, updated: new Date(Date.now() + 1000).toISOString() },
});

async function eventsSetup(
  source: "poll" | "graph",
  over: Partial<TriggerRunner> = {},
  caps: Partial<Caps> = {},
) {
  const store = createStore(sql!, SCHEMA);
  const events = createEventsStore(sql!, SCHEMA);
  const org = randomUUID();
  await store.ensureOrganization(org);
  const [{ id: agentId }] = await sql!`
    insert into ${name("agents")} (organization_id, name, key_hash, status)
    values (${org}, 'Helper', ${randomUUID().replaceAll("-", "").repeat(2)}, 'active')
    returning id::text`;
  await sql!`insert into ${name("connections")}
               (organization_id, connector_id, credential, credential_key_id, status, capabilities)
             values (${org}, 'app', ${Buffer.from("x")}, 'k', 'connected', ${sql!.array(["read"])})`;
  await sql!`insert into ${name("grants")} (agent_id, connector_id, capabilities, limits)
             values (${agentId}, 'app', ${sql!.array(["read"])}, ${sql!.json({ mailbox: MAILBOX })})`;
  // A pass and a poll round cover every organization: earlier tests' are finished with.
  await sql!`update ${name("event_runs")} set state = 'ended' where state in ('starting', 'running')`;
  await sql!`update ${name("event_subscriptions")} set polled_at = now()`;
  const id = randomUUID();
  const sealed = source === "graph" ? key.seal({ secret: SECRET }, subscriptionAad(org, id)) : null;
  const sub = (await events.createSubscription({
    id,
    organizationId: org,
    connectorId: "app",
    source,
    secret: sealed?.ciphertext ?? null,
    secretKeyId: sealed?.keyId ?? null,
    filter: "",
    intervalSeconds: 300,
    cursor: new Date(Date.now() - 60_000).toISOString(),
    createdBy: "ann",
  }))!;
  const { runner, started } = fakeRunner(over);
  await events.setTrigger(org, {
    agentId,
    connectorId: "app",
    enabled: true,
    eventTypes: ["received"],
    conditions: [],
    prompt: "Answer.",
    runner: runner.id,
    settings: {},
    secret: null,
    secretKeyId: null,
    managedAgentId: null,
    environmentId: null,
    vaultIds: [],
    budgetCents: 50,
    maxRunsPerHour: 100,
    maxCentsPerDay: 100_000,
    maxConcurrent: 10,
    updatedBy: "ann",
    ...caps,
  });
  const calls: VendorRequest[] = [];
  const items: unknown[] = [];
  const subscribed = new Set<string>();
  // What a test holds a call on, before the app answers.
  const before = { call: async (_req: VendorRequest) => {} };
  // Slow enough for the other process to reach the same step.
  async function appCall(_org: string, _app: string, req: VendorRequest): Promise<VendorResponse> {
    calls.push(req);
    await before.call(req);
    await sleep(30);
    if (req.method === "POST" && req.path === "/v1.0/subscriptions") {
      const body = req.body as { resource: string; expirationDateTime: string };
      if (subscribed.has(body.resource))
        return { status: 409, data: { error: { code: "Conflict" } } };
      subscribed.add(body.resource);
      return {
        status: 201,
        data: { id: `g-${subscribed.size}`, expirationDateTime: body.expirationDateTime },
      };
    }
    if (req.path.includes("/delta"))
      return {
        status: 200,
        data: { value: [], "@odata.deltaLink": "https://graph.example/v1.0/delta?t=done" },
      };
    if (req.path === "/items") return { status: 200, data: { items } };
    return { status: 404, data: null };
  }
  // One process of the service.
  const pod = (store2: EventsStore = events) =>
    createEventsRuntime({
      store,
      events: store2,
      key,
      connectors: new Map([["app", app]]),
      appCall,
      runners: new Map([[runner.id, runner]]),
      publicUrl: PUBLIC,
    });
  return { org, sub, store, events, calls, items, started, pod, before };
}

test(
  "a lease is its holder's until dropped or lapsed, and only its holder renews or drops it",
  { skip: !sql },
  async () => {
    const events = createEventsStore(sql!, SCHEMA);
    const lease = `test:${randomUUID()}`;
    const first = (await events.takeLease(lease, 60))!;
    assert.ok(first);
    assert.equal(await events.takeLease(lease, 60), null, "held");
    assert.equal(await events.takeLease(lease, 60, first), first, "renewed by its holder");
    await lapseLeases();
    const second = (await events.takeLease(lease, 60))!;
    assert.ok(second, "a lapsed lease is free");
    await events.dropLease(lease, first);
    assert.equal(await events.takeLease(lease, 60), null, "the old holder's drop does nothing");
    await events.dropLease(lease, second);
    assert.ok(await events.takeLease(lease, 60), "dropped");
  },
);

test(
  "two processes' syncs of a subscription make one Graph subscription for its watch",
  { skip: !sql },
  async () => {
    const t = await eventsSetup("graph");
    await Promise.all([t.pod().syncApp(t.org, "app"), t.pod().syncApp(t.org, "app")]);
    const posts = t.calls.filter((c) => c.method === "POST" && c.path === "/v1.0/subscriptions");
    assert.equal(posts.length, 1);
    const [w] = await t.events.watches(t.sub.id);
    assert.ok(w.remoteId, "the refused one did not clear its id");
    assert.equal(w.lastError, null);
  },
);

test(
  "a lifecycle notification waits for another process's work on the subscription",
  { skip: !sql },
  async () => {
    const t = await eventsSetup("graph");
    const a = t.pod();
    await a.syncApp(t.org, "app");
    const [w] = await t.events.watches(t.sub.id);
    t.calls.length = 0;
    const lease = `graph-sync:${t.sub.id}`;
    const token = (await t.events.takeLease(lease, 60))!;
    const b = t.pod();
    const res = await b.inbound(
      t.sub.id,
      new Request(`${PUBLIC}/events/${t.sub.id}`, {
        method: "POST",
        body: JSON.stringify({
          value: [
            {
              subscriptionId: w.remoteId,
              clientState: SECRET,
              lifecycleEvent: "reauthorizationRequired",
            },
          ],
        }),
      }),
    );
    assert.equal(res.status, 202);
    await sleep(100);
    assert.equal(t.calls.length, 0, "nothing while another process holds it");
    await t.events.dropLease(lease, token);
    await b.settle();
    assert.equal(t.calls[0]?.method, "PATCH");
  },
);

test(
  "a sync that loses its lease to another process makes no further Graph call",
  { skip: !sql },
  async (ctx) => {
    ctx.mock.timers.enable({ apis: ["setInterval"] });
    const t = await eventsSetup("graph");
    const posting = gate();
    t.before.call = async (req) => {
      if (req.method === "POST") await posting.pass();
    };
    let renewed!: () => void;
    const renewal = new Promise<void>((r) => (renewed = r));
    const renewing = {
      ...t.events,
      async takeLease(...args: Parameters<EventsStore["takeLease"]>) {
        const token = await t.events.takeLease(...args);
        if (args[2]) renewed();
        return token;
      },
    };
    const syncing = t.pod(renewing).syncApp(t.org, "app");
    await posting.entered;
    // Its Graph call stalled past the lease's term, and another process took the lease.
    await lapseLeases();
    assert.ok(await t.events.takeLease(`graph-sync:${t.sub.id}`, 60));
    ctx.mock.timers.tick(20_000);
    await renewal;
    posting.open();
    await syncing;
    assert.deepEqual(
      t.calls.map((c) => c.method),
      ["POST"],
      "no catch-up after its renewal found the lease taken",
    );
  },
);

test(
  "a finished sync leaves its lease free, with a renewal on its way or not",
  { skip: !sql },
  async (ctx) => {
    ctx.mock.timers.enable({ apis: ["setInterval"] });
    const t = await eventsSetup("graph");
    const posting = gate();
    t.before.call = async (req) => {
      if (req.method === "POST") await posting.pass();
    };
    const renewing = gate();
    let dropped!: () => void;
    const drop = new Promise<void>((r) => (dropped = r));
    let renewed!: () => void;
    const renewal = new Promise<void>((r) => (renewed = r));
    const slow = {
      ...t.events,
      async takeLease(...args: Parameters<EventsStore["takeLease"]>) {
        if (!args[2]) return t.events.takeLease(...args);
        await renewing.pass();
        const token = await t.events.takeLease(...args);
        renewed();
        return token;
      },
      async dropLease(...args: Parameters<EventsStore["dropLease"]>) {
        await t.events.dropLease(...args);
        dropped();
      },
    };
    const syncing = t.pod(slow).syncApp(t.org, "app");
    await posting.entered;
    ctx.mock.timers.tick(20_000);
    // A renewal is on its way to the database when the sync finishes.
    await renewing.entered;
    posting.open();
    await Promise.race([drop, sleep(200)]);
    renewing.open();
    await Promise.all([syncing, renewal]);
    assert.ok(await t.events.takeLease(`graph-sync:${t.sub.id}`, 60), "free for the next sync");
  },
);

test("one process makes each due poll", { skip: !sql }, async () => {
  const t = await eventsSetup("poll");
  const first = t.pod().pollDue();
  while (!t.calls.length) await sleep(5);
  // The first process is waiting on the app.
  await t.pod().pollDue();
  await first;
  assert.equal(t.calls.filter((c) => c.path === "/items").length, 1);
});

test("two processes cannot both take a trigger's last run slot", { skip: !sql }, async () => {
  const t = await eventsSetup("poll", {}, { maxConcurrent: 1 });
  t.items.push(item("i1"), item("i2"));
  // A stopped process queues both without working on them.
  const queueing = t.pod();
  await queueing.stop();
  await queueing.poll(t.sub);
  // Each process's count waits for the other's, so without the lock both say there is room.
  const counted = barrier(2);
  const racing = {
    ...t.events,
    runCounts: async (...args: Parameters<EventsStore["runCounts"]>) => {
      const counts = await t.events.runCounts(...args);
      await counted();
      return counts;
    },
  };
  const [a, b] = [t.pod(racing), t.pod(racing)];
  a.kick();
  b.kick();
  await Promise.all([a.idle(), b.idle()]);
  assert.equal(t.started.length, 1);
  const audit = await t.store.recentAudit(t.org, 50, { agentId: null, connectorId: null });
  assert.ok(audit.some((e) => e.reason === "Capped: runs at once reached"));
});

test("one process makes each pass over an app's running runs", { skip: !sql }, async () => {
  let reads = 0;
  const t = await eventsSetup("poll", {
    state: async () => {
      reads++;
      return running;
    },
  });
  t.items.push(item("i1"));
  const [a, b] = [t.pod(), t.pod()];
  await a.poll(t.sub);
  await a.idle();
  await Promise.all([a.watchTurn(), b.watchTurn()]);
  assert.equal(reads, 1);
  await lapseLeases();
  await b.watchTurn();
  assert.equal(reads, 2, "the next tick's pass, by the other process");
});

// ---- Teams ----

const TENANT = "33333333-3333-3333-3333-333333333333";
const PERSON = "44444444-4444-4444-4444-444444444444";
const SERVICE_URL = "https://smba.trafficmanager.net/amer/";
const auth = { token: async () => () => {} } as unknown as BotAuth;
type Sent = { method: string; path: string; body?: { type?: string; text?: string } };

async function teamsSetup(over: Partial<TriggerRunner> = {}, caps: Partial<Caps> = {}) {
  const store = createStore(sql!, SCHEMA);
  const teams = createTeamsStore(sql!, SCHEMA);
  // A pass covers every organization: earlier tests' sessions are finished with.
  await sql!`update ${name("teams_runs")} set state = 'ended' where state in ('starting', 'running')`;
  const org = randomUUID();
  await store.ensureOrganization(org);
  const [{ id: agentId }] = await sql!`
    insert into ${name("agents")} (organization_id, name, key_hash, status)
    values (${org}, 'Triage', ${randomUUID().replaceAll("-", "").repeat(2)}, 'active')
    returning id::text`;
  await sql!`insert into ${name("connections")}
               (organization_id, connector_id, credential, credential_key_id, status, capabilities)
             values (${org}, 'teams', ${Buffer.from("x")}, 'k', 'connected', ${sql!.array(["respond"])})`;
  await sql!`insert into ${name("grants")} (agent_id, connector_id, capabilities)
             values (${agentId}, 'teams', ${sql!.array(["respond"])})`;
  const botId = randomUUID();
  const botAppId = randomUUID();
  const sealed = key.seal({ private_key: "pk" }, botAad(org, botId));
  await sql!`insert into ${name("teams_bots")}
               (id, organization_id, agent_id, bot_app_id, tenant_id, display_name, status, config,
                credential, credential_key_id)
             values (${botId}, ${org}, ${agentId}, ${botAppId}, ${TENANT}, 'Triage', 'connected',
                     ${sql!.json({ certificate: "cert" })}, ${sealed.ciphertext}, ${sealed.keyId})`;
  const { runner, started } = fakeRunner(over);
  await teams.setTrigger(org, {
    agentId,
    enabled: true,
    runner: runner.id,
    settings: {},
    secret: null,
    secretKeyId: null,
    managedAgentId: null,
    environmentId: null,
    vaultIds: [],
    budgetCents: 50,
    maxRunsPerHour: 100,
    maxCentsPerDay: 100_000,
    maxConcurrent: 10,
    replyMode: "all",
    allowedTeams: [],
    allowedPeople: [PERSON],
    anyoneInOrganization: false,
    updatedBy: "ann",
    ...caps,
  });
  const sent: Sent[] = [];
  const bots = {
    connector: async (_bot: unknown, _url: string, req: Sent) => {
      sent.push(req);
      return { status: 200, data: { id: "900", value: [] } };
    },
    graph: async () => ({ status: 200, data: { value: [] }, headers: {} }),
  } as unknown as BotClient;
  const pod = () =>
    createTeamsRuntime({
      store,
      teams,
      key,
      auth,
      bots,
      runners: new Map([[runner.id, runner]]),
      connector: teamsConnector,
      publicUrl: PUBLIC,
    });
  type Pod = ReturnType<typeof pod>;
  let n = 0;
  // A direct message from PERSON, each in its own chat so none is held behind another.
  async function send(to: Pod, text: string) {
    n++;
    const res = await to.inbound(
      botAppId,
      new Request(`${PUBLIC}/teams/bots/${botAppId}/messages`, {
        method: "POST",
        headers: { authorization: "Bearer x", "content-type": "application/json" },
        body: JSON.stringify({
          type: "message",
          id: `act-${n}`,
          serviceUrl: SERVICE_URL,
          channelId: "msteams",
          recipient: { id: `28:${botAppId}` },
          from: { id: "29:person", aadObjectId: PERSON, name: "Ann" },
          conversation: { id: `a:chat-${n}`, conversationType: "personal" },
          channelData: { tenant: { id: TENANT } },
          text,
        }),
      }),
    );
    assert.equal(res.status, 200);
  }
  const told = () => sent.filter((s) => s.body?.type === "message").map((s) => s.body!.text);
  return { org, agentId, botId, teams, sent, started, pod, send, told };
}

test("two processes cannot both take an agent's last run slot", { skip: !sql }, async () => {
  // The first process has read its counts and is held in the checks after them.
  const held = gate();
  let first = true;
  const t = await teamsSetup(
    {
      unready: async () => {
        if (first) {
          first = false;
          await held.pass();
        }
        return null;
      },
    },
    { maxConcurrent: 1 },
  );
  const [a, b] = [t.pod(), t.pod()];
  await t.send(a, "one");
  await held.entered;
  await t.send(b, "two");
  await b.idle();
  held.open();
  await a.idle();
  assert.equal(t.started.length, 1);
  assert.deepEqual(t.told(), [
    "I'm busy in other conversations right now. Please try again in a few minutes.",
  ]);
});

test("an agent's last run slot goes to one process when both admit at once", { skip: !sql }, async () => {
  const t = await teamsSetup();
  // Each process's count waits for the other's, so without the lock both say there is room.
  const counted = barrier(2);
  const racing = {
    ...t.teams,
    runCounts: async (...args: Parameters<typeof t.teams.runCounts>) => {
      const counts = await t.teams.runCounts(...args);
      await counted();
      return counts;
    },
  };
  const admit = (i: number) =>
    racing.admit(
      {
        agentId: t.agentId,
        botId: t.botId,
        conversationId: `a:chat-${i}`,
        threadId: `a:chat-${i}`,
        triggerActivityId: `act-${i}`,
        replyTokenHash: hashReplyToken(`token-${i}`),
        requestedBy: PERSON,
        runner: "fake",
      },
      null,
      { budgetCents: 50, maxRunsPerHour: 100, maxCentsPerDay: 100_000, maxConcurrent: 1 },
    );
  const answers = await Promise.all([admit(1), admit(2)]);
  assert.deepEqual(answers.map((a) => ("cap" in a ? a.cap : "started")).sort(), [
    "concurrent",
    "started",
  ]);
});

test(
  "a message another process is still handing over counts towards the hourly limit",
  { skip: !sql },
  async () => {
    const held = gate();
    let first = true;
    const started: RunInput[] = [];
    const t = await teamsSetup(
      {
        async start(input) {
          started.push(input);
          if (first) {
            first = false;
            await held.pass();
          }
          return `s-${input.runId}`;
        },
      },
      { maxRunsPerHour: 1 },
    );
    const [a, b] = [t.pod(), t.pod()];
    await t.send(a, "one");
    await held.entered;
    await t.send(b, "two");
    await b.idle();
    assert.equal(started.length, 1, "the second is over the limit");
    held.open();
    await a.idle();
    assert.deepEqual(t.told(), ["I've reached my limit for this hour. Please try again later."]);
  },
);

test(
  "overlapping passes log a finished session's spend and tell the person once",
  { skip: !sql },
  async () => {
    const t = await teamsSetup({
      state: async () => {
        await sleep(30);
        return { status: "idle", costCents: 30, atBudget: false };
      },
    });
    const [a, b] = [t.pod(), t.pod()];
    await t.send(a, "hello");
    await a.idle();
    await Promise.all([a.watch(), b.watch()]);
    assert.deepEqual(t.told(), [
      "I could not answer this one. Ask your admin to check my Teams settings.",
    ]);
    const spend = await sql!`select cents from ${name("teams_spend")} where agent_id = ${t.agentId}`;
    assert.deepEqual(
      spend.map((s) => s.cents),
      [30],
    );
  },
);

test("one process makes each pass over the running sessions", { skip: !sql }, async () => {
  let reads = 0;
  const t = await teamsSetup({
    state: async () => {
      reads++;
      return running;
    },
  });
  const [a, b] = [t.pod(), t.pod()];
  await t.send(a, "hello");
  await a.idle();
  await Promise.all([a.watchTurn(), b.watchTurn()]);
  assert.equal(reads, 1);
  await lapseLeases();
  await b.watchTurn();
  assert.equal(reads, 2, "the next tick's pass, by the other process");
});
