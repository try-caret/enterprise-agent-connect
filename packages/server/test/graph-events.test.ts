import assert from "node:assert/strict";
import { randomUUID } from "node:crypto";
import { readFileSync } from "node:fs";
import { after, before, test } from "node:test";
import { z } from "zod";
import { parseConnector } from "../../connector-kit/src/connector.ts";
import type { VendorRequest, VendorResponse } from "../../connector-kit/src/types.ts";
import { createApp } from "../src/app.ts";
import { createEventsAdmin } from "../src/events/admin.ts";
import { createEventsRuntime } from "../src/events/runtime.ts";
import type { ReportTags } from "../src/report.ts";
import { createEventsStore } from "../src/events/store.ts";
import { migrate } from "../src/migrate.ts";
import { CredentialKey } from "../src/secrets.ts";
import { connectDatabase, createStore } from "../src/store.ts";
import {
  signedWebhook,
  type EventRunInput,
  type RunnerContext,
  type TriggerRunner,
} from "../src/teams/runner.ts";
import { ADMIN_TOKEN } from "./helpers.ts";

// Microsoft Graph change notifications as an event source, on a made-up mail app ("mailer")
// against a fake Graph, on the scratch database (TEST_DATABASE_URL). Skipped without one.
const url = process.env.TEST_DATABASE_URL;
const sql = url ? connectDatabase(url) : null;
const SCHEMA = "agent_connect";
before(async () => {
  if (sql) await migrate(sql, SCHEMA, () => {});
});
after(() => sql?.end());

const key = new CredentialKey(Buffer.alloc(32, 7).toString("base64"));
const PUBLIC = "https://svc.example/agent-connect";
const MAILBOX = "helper@example.com";
const base = JSON.parse(
  readFileSync(new URL("../../../connectors/outlook-mail/connector.json", import.meta.url), "utf8"),
);
const mailer = parseConnector(
  {
    ...base,
    id: "mailer",
    name: "Mailer",
    events: {
      types: [{ id: "received", label: "Mail received" }],
      fields: {
        mailbox: "mailbox",
        id: "message.id",
        conversation: "message.conversationId",
        from: "message.from.emailAddress.address",
        to: "message.toRecipients[].emailAddress.address",
        cc: "message.ccRecipients[].emailAddress.address",
        subject: "message.subject",
      },
      limits: { mailbox: "mailbox" },
      thread: "conversation",
      graph: {
        per: "mailbox",
        resource: "users/{value}/inbox/messages",
        changeType: "created",
        type: "received",
        maxMinutes: 10070,
        read: { method: "GET", path: "/v1.0/users/{value}/messages/{id}" },
        delta: { method: "GET", path: "/v1.0/users/{value}/inbox/delta?since={since}" },
        wrap: "message",
        deliveryId: "message.id",
        received: "message.receivedDateTime",
        own: ["from"],
        only: { only_me: ["to", "cc"] },
      },
    },
  },
  "mailer",
);

type Fake = { runner: TriggerRunner; started: EventRunInput[] };
function fakeRunner(): Fake {
  const started: EventRunInput[] = [];
  const runner: TriggerRunner = {
    id: "fake",
    settings: z.strictObject({}),
    unready: async () => null,
    verify: async () => {},
    start: async () => {
      throw new Error("not for Teams");
    },
    async startEvent(input) {
      started.push(input);
      return `s-${input.runId}`;
    },
    continue: async () => {},
    state: async () => ({ status: "running", costCents: 0, atBudget: false }),
    stop: async () => {},
  };
  return { runner, started };
}

const message = (id: string, over: Record<string, unknown> = {}) => ({
  id,
  conversationId: "conv-1",
  from: { emailAddress: { address: "ann@example.com" } },
  toRecipients: [{ emailAddress: { address: MAILBOX } }],
  ccRecipients: [],
  subject: `About ${id}`,
  receivedDateTime: new Date(Date.now() + 1000).toISOString(),
  ...over,
});

// A fake Graph: subscriptions made, renewed and deleted; messages read by id; delta pages.
function fakeGraph() {
  const calls: VendorRequest[] = [];
  const messages = new Map<string, unknown>();
  const subscriptions = new Map<string, Record<string, unknown>>();
  let made = 0;
  let renew = 200;
  let deltaPages: VendorResponse[] = [];
  async function appCall(_org: string, _app: string, req: VendorRequest) {
    calls.push(req);
    if (req.path === "/v1.0/subscriptions" && req.method === "POST") {
      const id = `g-${++made}`;
      const body = req.body as Record<string, unknown>;
      subscriptions.set(id, body);
      return { status: 201, data: { id, expirationDateTime: body.expirationDateTime } };
    }
    const sub = /^\/v1\.0\/subscriptions\/(.+)$/.exec(req.path);
    if (sub && req.method === "PATCH")
      return { status: subscriptions.has(sub[1]) ? renew : 404, data: req.body };
    if (sub && req.method === "DELETE")
      return { status: subscriptions.delete(sub[1]) ? 204 : 404, data: null };
    const read = /^\/v1\.0\/users\/[^/]+\/messages\/([^/?]+)$/.exec(req.path);
    if (read) {
      const m = messages.get(decodeURIComponent(read[1]));
      return m ? { status: 200, data: m } : { status: 404, data: null };
    }
    if (req.path.includes("/delta"))
      return (
        deltaPages.shift() ?? {
          status: 200,
          data: {
            value: [],
            "@odata.deltaLink": `https://graph.example/v1.0/users/x/inbox/delta?$deltatoken=done`,
          },
        }
      );
    return { status: 404, data: null };
  }
  return {
    appCall,
    calls,
    messages,
    subscriptions,
    setRenew: (status: number) => (renew = status),
    setDelta: (pages: VendorResponse[]) => (deltaPages = pages),
  };
}

async function setup(runner: TriggerRunner) {
  const store = createStore(sql!, SCHEMA);
  const events = createEventsStore(sql!, SCHEMA);
  const org = randomUUID();
  await store.ensureOrganization(org);
  const agent = async (name: string, mailbox: string) => {
    const [{ id }] = await sql!`
      insert into agent_connect.agents (organization_id, name, key_hash, status)
      values (${org}, ${name}, ${randomUUID().replaceAll("-", "").repeat(2)}, 'active')
      returning id::text`;
    await sql!`insert into agent_connect.grants (agent_id, connector_id, capabilities, limits)
               values (${id}, 'mailer', ${sql!.array(["read"])}, ${sql!.json({ mailbox })})`;
    return id as string;
  };
  await sql!`insert into agent_connect.connections
               (organization_id, connector_id, credential, credential_key_id, status, capabilities)
             values (${org}, 'mailer', ${Buffer.from("x")}, 'k', 'connected', ${sql!.array(["read"])})`;
  const agentId = await agent("Helper", MAILBOX);
  const graph = fakeGraph();
  const reports: { kind: string; tags?: ReportTags }[] = [];
  const deps = {
    store,
    events,
    key,
    connectors: new Map([["mailer", mailer]]),
    appCall: graph.appCall,
    runners: new Map([[runner.id, runner]]),
    publicUrl: PUBLIC,
    report: (kind: string, _error: unknown, tags?: ReportTags) => reports.push({ kind, tags }),
  };
  const runtime = createEventsRuntime(deps);
  const admin = createEventsAdmin({ ...deps, runtime });
  const app = createApp({
    service: {} as never,
    store,
    adminToken: ADMIN_TOKEN,
    events: { admin, runtime },
    validateResponses: true,
  });
  async function call(method: string, path: string, body?: unknown) {
    const res = await app.fetch(
      new Request(`http://svc/agent-connect/admin/v1/orgs/${org}${path}`, {
        method,
        headers: { authorization: `Bearer ${ADMIN_TOKEN}`, "content-type": "application/json" },
        body: body === undefined ? undefined : JSON.stringify(body),
      }),
    );
    return { status: res.status, body: (res.status === 204 ? null : await res.json()) as any };
  }
  async function subscribe() {
    const res = await call("POST", "/event-subscriptions", {
      app: "mailer",
      source: "graph",
      by: "ann",
    });
    assert.equal(res.status, 200, JSON.stringify(res.body));
    return res.body as { subscription: { id: string; url: string }; secret: null };
  }
  async function trigger(agent = agentId, over: Record<string, unknown> = {}) {
    const res = await call("PUT", `/agents/${agent}/event-triggers/mailer`, {
      enabled: true,
      eventTypes: ["received"],
      prompt: "Answer in the thread.",
      runner: runner.id,
      budgetCents: 50,
      by: "ann",
      ...over,
    });
    assert.equal(res.status, 200, JSON.stringify(res.body));
  }
  const watches = async (subId: string) =>
    (await call("GET", `/event-subscriptions/${subId}/watches`)).body as {
      value: string;
      subscribed: boolean;
      lastNotificationAt: string | null;
      lastError: string | null;
    }[];
  async function notify(subId: string, value: unknown[]) {
    const res = await app.fetch(
      new Request(`http://svc/agent-connect/events/${subId}`, {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ value }),
      }),
    );
    await runtime.settle();
    await runtime.idle();
    return res.status;
  }
  const audit = async () =>
    (await store.recentAudit(org, 200, { agentId: null, connectorId: null })).map(
      (a) => `${a.tool} ${a.decision} ${a.reason}`,
    );
  return {
    org,
    agentId,
    agent,
    events,
    runtime,
    app,
    call,
    subscribe,
    trigger,
    watches,
    notify,
    audit,
    graph,
    reports,
  };
}

// What Graph sends for a new message in a watched mailbox.
const created = (remote: string, clientState: string, id: string) => ({
  subscriptionId: remote,
  clientState,
  changeType: "created",
  resource: `Users/0000/Messages/${id}`,
  resourceData: { "@odata.type": "#Microsoft.Graph.Message", id },
});

test(
  "Graph's validation handshake is answered with the token as plain text",
  { skip: !sql },
  async () => {
    const t = await setup(fakeRunner().runner);
    const { subscription } = await t.subscribe();
    const res = await t.app.fetch(
      new Request(
        `http://svc/agent-connect/events/${subscription.id}?validationToken=${encodeURIComponent("Validation: <token> 1")}`,
        { method: "POST", headers: { "content-type": "text/plain" } },
      ),
    );
    assert.equal(res.status, 200);
    assert.equal(res.headers.get("content-type"), "text/plain; charset=utf-8");
    assert.equal(res.headers.get("x-content-type-options"), "nosniff");
    assert.equal(await res.text(), "Validation: <token> 1");
    // Only for a Graph subscription that exists here.
    const unknown = await t.app.fetch(
      new Request(`http://svc/agent-connect/events/${randomUUID()}?validationToken=x`, {
        method: "POST",
      }),
    );
    assert.equal(unknown.status, 404);
  },
);

test(
  "a subscription watches each mailbox with a trigger, and follows triggers on and off",
  { skip: !sql },
  async () => {
    const t = await setup(fakeRunner().runner);
    const { subscription, secret } = await t.subscribe();
    assert.equal(secret, null, "the clientState is never shown");
    assert.equal(subscription.url, `${PUBLIC}/events/${subscription.id}`);
    assert.deepEqual(await t.watches(subscription.id), [], "no trigger, nothing watched");

    await t.trigger();
    const [made] = t.graph.subscriptions.values();
    assert.equal(made.resource, `users/${MAILBOX}/inbox/messages`);
    assert.equal(made.changeType, "created");
    assert.equal(made.notificationUrl, subscription.url);
    assert.equal(made.lifecycleNotificationUrl, subscription.url);
    assert.match(String(made.clientState), /^[0-9a-f]{64}$/);
    const minutes = (Date.parse(String(made.expirationDateTime)) - Date.now()) / 60_000;
    assert.ok(minutes > 10060 && minutes <= 10070, `expires in ${minutes} minutes`);
    const [w] = await t.watches(subscription.id);
    assert.equal(w.value, MAILBOX);
    assert.equal(w.subscribed, true);

    // Two agents on one mailbox share its watch; a second mailbox gets its own.
    await t.trigger(await t.agent("Twin", MAILBOX));
    const other = await t.agent("Other", "other@example.com");
    await t.trigger(other);
    assert.deepEqual(
      (await t.watches(subscription.id)).map((x) => x.value),
      ["helper@example.com", "other@example.com"],
    );
    assert.equal(t.graph.subscriptions.size, 2);

    // Its trigger off: the mailbox's Graph subscription goes.
    await t.trigger(other, { enabled: false });
    assert.deepEqual(
      (await t.watches(subscription.id)).map((x) => x.value),
      [MAILBOX],
    );
    assert.equal(t.graph.subscriptions.size, 1);

    // Unsubscribing removes every Graph subscription first.
    assert.equal(
      (await t.call("DELETE", `/event-subscriptions/${subscription.id}`, { by: "ann" })).status,
      204,
    );
    assert.equal(t.graph.subscriptions.size, 0);
  },
);

test(
  "a notification batch starts runs once each: clientState checked, own and old mail skipped",
  { skip: !sql },
  async () => {
    const f = fakeRunner();
    const t = await setup(f.runner);
    const { subscription } = await t.subscribe();
    await t.trigger();
    const [[remote, made]] = t.graph.subscriptions.entries();
    const state = String(made.clientState);
    t.graph.messages.set(
      "m1",
      message("m1", { ccRecipients: [{ emailAddress: { address: "bob@example.com" } }] }),
    );
    t.graph.messages.set("m2", message("m2"));

    // One proven, one not: the proven one is worked on.
    assert.equal(
      await t.notify(subscription.id, [
        created(remote, state, "m1"),
        created(remote, "forged", "m2"),
      ]),
      202,
    );
    assert.equal(f.started.length, 1);
    assert.deepEqual(f.started[0].event.data, {
      mailbox: MAILBOX,
      id: "m1",
      conversation: "conv-1",
      from: "ann@example.com",
      to: [MAILBOX],
      cc: ["bob@example.com"],
      subject: "About m1",
      only_me: false,
    });
    assert.ok(
      t.graph.calls.some(
        (c) => c.path === `/v1.0/users/${encodeURIComponent(MAILBOX)}/messages/m1`,
      ),
    );
    // None proven: refused.
    assert.equal(await t.notify(subscription.id, [created(remote, "forged", "m2")]), 401);
    // Not a batch.
    assert.equal(await t.notify(subscription.id, {} as never), 400);

    // The same notification again (Graph retries): no second run.
    assert.equal(await t.notify(subscription.id, [created(remote, state, "m1")]), 202);
    assert.equal(f.started.length, 1);

    // Sent by the mailbox itself, received before the watch started, gone already, or for a
    // Graph subscription not ours: nothing starts.
    t.graph.messages.set(
      "own",
      message("own", { from: { emailAddress: { address: "Helper@Example.com" } } }),
    );
    t.graph.messages.set("old", message("old", { receivedDateTime: "2020-01-01T00:00:00Z" }));
    await t.notify(subscription.id, [
      created(remote, state, "own"),
      created(remote, state, "old"),
      created(remote, state, "gone"),
      created("someone-else", state, "m2"),
    ]);
    assert.equal(f.started.length, 1);

    // Only the mailbox addressed.
    await t.notify(subscription.id, [created(remote, state, "m2")]);
    assert.equal(f.started.length, 2);
    assert.equal(f.started[1].event.data.only_me, true);

    const log = await t.audit();
    assert.equal(log.filter((l) => l.includes("clientState did not match")).length, 1);
    assert.ok(log.includes(`events:delivery denied Ignored: sent by ${MAILBOX} itself`));
    assert.ok(log.includes("events:delivery denied Ignored: the item is gone"));
    const [w] = await t.watches(subscription.id);
    assert.ok(w.lastNotificationAt);
  },
);

test(
  "the catch-up reads the delta from its cursor; what notifications started is not started again",
  { skip: !sql },
  async () => {
    const f = fakeRunner();
    const t = await setup(f.runner);
    const { subscription } = await t.subscribe();
    await t.trigger();
    // The first round starts from when the watch started.
    const first = t.graph.calls.find((c) => c.path.includes("/delta"))!;
    assert.match(
      first.path,
      new RegExp(`^/v1\\.0/users/${encodeURIComponent(MAILBOX)}/inbox/delta\\?since=\\d{4}-`),
    );

    const [[remote, made]] = t.graph.subscriptions.entries();
    t.graph.messages.set("m1", message("m1"));
    await t.notify(subscription.id, [created(remote, String(made.clientState), "m1")]);
    assert.equal(f.started.length, 1);

    // Two pages: m1 again and m2 (missed), then an old one and a removal.
    t.graph.setDelta([
      {
        status: 200,
        data: {
          value: [message("m1"), message("m2")],
          "@odata.nextLink": "https://graph.example/v1.0/users/x/inbox/delta?$skiptoken=p2",
        },
      },
      {
        status: 200,
        data: {
          value: [
            message("old", { receivedDateTime: "2020-01-01T00:00:00Z" }),
            { id: "m3", "@removed": { reason: "deleted" } },
          ],
          "@odata.deltaLink": "https://graph.example/v1.0/users/x/inbox/delta?$deltatoken=t1",
        },
      },
    ]);
    await t.runtime.syncApp(t.org, "mailer");
    await t.runtime.idle();
    assert.deepEqual(
      f.started.map((s) => s.event.data.id),
      ["m1", "m2"],
    );
    const deltas = t.graph.calls.filter((c) => c.path.includes("/delta")).map((c) => c.path);
    assert.equal(
      deltas.at(-1),
      "/v1.0/users/x/inbox/delta?$skiptoken=p2",
      "links keep only path and query",
    );

    // The next round continues from the delta link.
    await t.runtime.syncApp(t.org, "mailer");
    assert.equal(
      t.graph.calls.filter((c) => c.path.includes("/delta")).at(-1)!.path,
      "/v1.0/users/x/inbox/delta?$deltatoken=t1",
    );

    // A delta link Graph no longer knows starts over from the watch's start.
    t.graph.setDelta([{ status: 410, data: null }]);
    await t.runtime.syncApp(t.org, "mailer");
    await t.runtime.syncApp(t.org, "mailer");
    assert.match(
      t.graph.calls.filter((c) => c.path.includes("/delta")).at(-1)!.path,
      /\/inbox\/delta\?since=/,
    );
    assert.equal(f.started.length, 2);
  },
);

test(
  "renewal past half its life, and lifecycle notifications: reauthorize, removed, missed",
  { skip: !sql },
  async () => {
    const f = fakeRunner();
    const t = await setup(f.runner);
    const { subscription } = await t.subscribe();
    await t.trigger();
    const [[remote, made]] = t.graph.subscriptions.entries();
    const state = String(made.clientState);
    const patches = () => t.graph.calls.filter((c) => c.method === "PATCH").length;

    // Not yet due: nothing renewed.
    await t.runtime.syncApp(t.org, "mailer");
    assert.equal(patches(), 0);
    // Past half its life: renewed in place.
    await t.events.watchSubscribed(
      (await t.events.watches(subscription.id))[0].id,
      remote,
      new Date(Date.now() + 3600_000).toISOString(),
      null,
    );
    await t.runtime.syncApp(t.org, "mailer");
    assert.equal(patches(), 1);
    const [w] = await t.watches(subscription.id);
    assert.equal(w.subscribed, true);

    // Graph asks for reauthorization: renewed again.
    await t.notify(subscription.id, [
      { subscriptionId: remote, clientState: state, lifecycleEvent: "reauthorizationRequired" },
    ]);
    assert.equal(patches(), 2);

    // A renewal Graph refuses is shown on the watch.
    t.graph.setRenew(403);
    await t.notify(subscription.id, [
      { subscriptionId: remote, clientState: state, lifecycleEvent: "reauthorizationRequired" },
    ]);
    assert.match(
      (await t.watches(subscription.id))[0].lastError ?? "",
      /^Renewing: Microsoft answered 403/,
    );
    t.graph.setRenew(200);

    // Missed notifications: caught up at once.
    t.graph.setDelta([
      {
        status: 200,
        data: {
          value: [message("m9")],
          "@odata.deltaLink": "https://graph.example/v1.0/users/x/inbox/delta?$deltatoken=x",
        },
      },
    ]);
    await t.notify(subscription.id, [
      { subscriptionId: remote, clientState: state, lifecycleEvent: "missed" },
    ]);
    assert.deepEqual(
      f.started.map((s) => s.event.data.id),
      ["m9"],
    );

    // Graph removed the subscription: a new one is made under the same watch.
    t.graph.subscriptions.delete(remote);
    await t.notify(subscription.id, [
      { subscriptionId: remote, clientState: state, lifecycleEvent: "subscriptionRemoved" },
    ]);
    assert.deepEqual([...t.graph.subscriptions.keys()], ["g-2"]);
    const [after] = await t.watches(subscription.id);
    assert.equal(after.subscribed, true);
    assert.equal(after.lastError, null);
    // A renewal Graph no longer knows makes a new one too.
    t.graph.subscriptions.clear();
    await t.events.watchSubscribed(
      (await t.events.watches(subscription.id))[0].id,
      "g-2",
      new Date(Date.now() + 60_000).toISOString(),
      null,
    );
    await t.runtime.syncApp(t.org, "mailer");
    assert.deepEqual([...t.graph.subscriptions.keys()], ["g-3"]);
    // Lifecycle notifications must be proven too.
    assert.equal(
      await t.notify(subscription.id, [
        { subscriptionId: "g-3", clientState: "x", lifecycleEvent: "missed" },
      ]),
      401,
    );
  },
);

test(
  "events of one thread carry one thread id per trigger, to the runner and the signed webhook",
  { skip: !sql },
  async () => {
    const got: string[] = [];
    const fetch = (async (_url: string, init: RequestInit) => {
      got.push(String(init.body));
      return new Response(null, { status: 200 });
    }) as never;
    const ctx = {
      teams: { claudeKey: async () => null } as unknown as RunnerContext["teams"],
      key,
      publicUrl: PUBLIC,
    };
    const t = await setup(signedWebhook(fetch)(ctx));
    const { subscription } = await t.subscribe();
    const hook = {
      runner: "webhook",
      settings: { url: "https://hooks.example/mail" },
      signingSecret: "r".repeat(40),
    };
    await t.trigger(t.agentId, hook);
    const twin = await t.agent("Twin", MAILBOX);
    await t.trigger(twin, hook);
    const [[remote, made]] = t.graph.subscriptions.entries();
    const state = String(made.clientState);
    t.graph.messages.set("a1", message("a1", { conversationId: "conv-A" }));
    t.graph.messages.set("a2", message("a2", { conversationId: "conv-A" }));
    t.graph.messages.set("b1", message("b1", { conversationId: "conv-B" }));
    // Each run ends (the webhook answered) before the next mail, so no cap holds one back.
    for (const id of ["a1", "a2", "b1"]) {
      await t.notify(subscription.id, [created(remote, state, id)]);
      await t.runtime.watch();
    }

    const sent = got
      .map((b) => JSON.parse(b))
      .filter((b) => b.type === "event")
      .map((b) => ({ agent: b.agent, id: b.data.id, thread: b.thread, run: b.run }));
    assert.equal(sent.length, 6);
    const of = (agent: string, id: string) => sent.find((s) => s.agent === agent && s.id === id)!;
    assert.match(of(t.agentId, "a1").thread, /^[0-9a-f]{32}$/);
    assert.equal(of(t.agentId, "a1").thread, of(t.agentId, "a2").thread, "one thread");
    assert.notEqual(of(t.agentId, "a1").run, of(t.agentId, "a2").run, "each event its own run");
    assert.notEqual(of(t.agentId, "a1").thread, of(t.agentId, "b1").thread, "another thread");
    assert.notEqual(of(t.agentId, "a1").thread, of(twin, "a1").thread, "another trigger");
    assert.ok(!sent.some((s) => s.thread.includes("conv")), "the app's own id is not shown");

    const runs = (await t.call("GET", `/agents/${t.agentId}/event-runs`)).body as {
      thread: string | null;
      deliveryId: string;
    }[];
    assert.equal(runs.find((r) => r.deliveryId === "a2")!.thread, of(t.agentId, "a2").thread);
  },
);

test("an app's graph block needs its limit, so no trigger hears another agent's mailbox", () => {
  const events = (mailer as unknown as { events: Record<string, unknown> }).events;
  assert.throws(
    () => parseConnector({ ...base, id: "mailer", events: { ...events, limits: {} } }, "mailer"),
    /event limit on `per`/,
  );
  assert.throws(
    () =>
      parseConnector({ ...base, id: "mailer", events: { ...events, thread: "nope" } }, "mailer"),
    /thread name fields/,
  );
});

test("a Graph catch-up Microsoft refuses is reported with its status", { skip: !sql }, async () => {
  const t = await setup(fakeRunner().runner);
  const { subscription } = await t.subscribe();
  await t.trigger();
  t.graph.setDelta([{ status: 403, data: null }]);
  await t.runtime.syncApp(t.org, "mailer");
  assert.deepEqual(
    t.reports.filter((r) => r.kind === "event_poll_failed").map((r) => r.tags),
    [{ app: "mailer", org: t.org, subscription: subscription.id, source: "graph", status: 403 }],
  );
});
