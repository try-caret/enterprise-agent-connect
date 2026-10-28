import assert from "node:assert/strict";
import { createHmac, randomUUID } from "node:crypto";
import { readFileSync } from "node:fs";
import { after, before, test } from "node:test";
import { z } from "zod";
import { parseConnector } from "../../connector-kit/src/connector.ts";
import type { VendorRequest, VendorResponse } from "../../connector-kit/src/types.ts";
import { createApp } from "../src/app.ts";
import { createEventsAdmin } from "../src/events/admin.ts";
import { createEventsRuntime } from "../src/events/runtime.ts";
import { createEventsStore } from "../src/events/store.ts";
import { migrate } from "../src/migrate.ts";
import { CredentialKey } from "../src/secrets.ts";
import { connectDatabase, createStore } from "../src/store.ts";
import {
  signedWebhook,
  SIGNATURE_HEADER,
  type EventRunInput,
  type RunnerContext,
  type TriggerRunner,
} from "../src/teams/runner.ts";
import { ADMIN_TOKEN } from "./helpers.ts";

// App events end to end on a made-up app ("tracker"), against the scratch database
// (TEST_DATABASE_URL, see store.test.ts). Skipped without one.
const url = process.env.TEST_DATABASE_URL;
const sql = url ? connectDatabase(url) : null;
const SCHEMA = "agent_connect";
before(async () => {
  if (sql) await migrate(sql, SCHEMA, () => {});
});
after(() => sql?.end());

const key = new CredentialKey(Buffer.alloc(32, 7).toString("base64"));
const PUBLIC = "https://svc.example/agent-connect";
const jiraJson = JSON.parse(
  readFileSync(new URL("../../../connectors/jira/connector.json", import.meta.url), "utf8"),
);
const tracker = parseConnector(
  {
    ...jiraJson,
    id: "tracker",
    name: "Tracker",
    events: {
      types: [
        { id: "created", label: "Item created" },
        { id: "updated", label: "Item updated" },
      ],
      fields: { key: "item.key", project: "item.project", type: "item.type", by: "actor" },
      limits: { projects: "project" },
      self: { request: { method: "GET", path: "/me" }, id: "id", fields: ["by"] },
      webhook: {
        verify: { type: "hmac-sha256", header: "X-Sig", prefix: "sha256=" },
        type: { field: "event" },
        deliveryId: { header: "X-Delivery" },
        timestamp: { at: { field: "ts" }, unit: "ms", toleranceSeconds: 300 },
        setup: [
          { title: "Add it", body: "Paste the address.", click: ["Settings > Hooks > {url}"] },
        ],
        register: {
          method: "POST",
          path: "/hooks",
          body: { url: "{url}", secret: "{secret}", events: "{events}", filter: "{filter}" },
          location: "self",
        },
      },
      poll: {
        request: { method: "POST", path: "/search", body: { q: "{filter}since -{minutes}m" } },
        filter: "({value}) and ",
        items: "items",
        wrap: "item",
        cursor: "item.updated",
        deliveryId: ["item.key", "item.updated"],
        types: [{ type: "created", since: "item.created" }, { type: "updated" }],
      },
    },
  },
  "tracker",
);

type Fake = { runner: TriggerRunner; started: EventRunInput[]; states: Map<string, string> };
// A runner that records each start; a run stays running until its state is set.
function fakeRunner(id = "fake"): Fake {
  const started: EventRunInput[] = [];
  const states = new Map<string, string>();
  const runner: TriggerRunner = {
    id,
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
    async state(run) {
      const s = states.get(run.sessionId!) ?? "running";
      return { status: s as "running", costCents: s === "idle" ? 30 : 0, atBudget: false };
    },
    stop: async () => {},
  };
  return { runner, started, states };
}

async function setup(
  runners: TriggerRunner[],
  limits: Record<string, unknown> = { projects: ["KAN"] },
) {
  const store = createStore(sql!, SCHEMA);
  const events = createEventsStore(sql!, SCHEMA);
  const org = randomUUID();
  await store.ensureOrganization(org);
  const [{ id: agentId }] = await sql!`
    insert into agent_connect.agents (organization_id, name, key_hash, status)
    values (${org}, 'Triage', ${randomUUID().replaceAll("-", "").repeat(2)}, 'active')
    returning id::text`;
  await sql!`insert into agent_connect.connections
               (organization_id, connector_id, credential, credential_key_id, status, capabilities)
             values (${org}, 'tracker', ${Buffer.from("x")}, 'k', 'connected', ${sql!.array(["read"])})`;
  await sql!`insert into agent_connect.grants (agent_id, connector_id, capabilities, limits)
             values (${agentId}, 'tracker', ${sql!.array(["read"])}, ${sql!.json(limits as never)})`;
  const calls: VendorRequest[] = [];
  const answers = new Map<string, VendorResponse>([
    ["/me", { status: 200, data: { id: "svc" } }],
    ["/hooks", { status: 403, data: null }],
  ]);
  const appCall = async (_org: string, _app: string, request: VendorRequest) => {
    calls.push(request);
    return answers.get(request.path) ?? { status: 404, data: null };
  };
  const deps = {
    store,
    events,
    key,
    connectors: new Map([["tracker", tracker]]),
    appCall,
    runners: new Map(runners.map((r) => [r.id, r])),
  };
  const runtime = createEventsRuntime(deps);
  const admin = createEventsAdmin({ ...deps, publicUrl: PUBLIC });
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
  async function subscribe(source: "webhook" | "poll" = "webhook", extra = {}) {
    const res = await call("POST", "/event-subscriptions", {
      app: "tracker",
      source,
      by: "ann",
      ...extra,
    });
    assert.equal(res.status, 200, JSON.stringify(res.body));
    return res.body as {
      subscription: { id: string; url: string | null };
      secret: string | null;
      setup: { click: string[] }[];
      registration: { ok: boolean; message: string } | null;
    };
  }
  async function trigger(over: Record<string, unknown> = {}) {
    const res = await call("PUT", `/agents/${agentId}/event-triggers/tracker`, {
      enabled: true,
      eventTypes: ["created"],
      prompt: "Triage it.",
      runner: runners[0].id,
      budgetCents: 50,
      by: "ann",
      ...over,
    });
    assert.equal(res.status, 200, JSON.stringify(res.body));
    return res.body;
  }
  let n = 0;
  async function deliver(
    subId: string,
    secret: string,
    body: Record<string, unknown>,
    opts: { delivery?: string; sig?: string } = {},
  ) {
    const raw = JSON.stringify({ ts: Date.now(), ...body });
    const res = await app.fetch(
      new Request(`${PUBLIC.replace("https://svc.example", "http://svc")}/events/${subId}`, {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-delivery": opts.delivery ?? `d${++n}`,
          "x-sig": opts.sig ?? `sha256=${createHmac("sha256", secret).update(raw).digest("hex")}`,
        },
        body: raw,
      }),
    );
    await runtime.idle();
    return res.status;
  }
  const audit = async () =>
    (await store.recentAudit(org, 200, { agentId: null, connectorId: null })).map(
      (a) => `${a.tool} ${a.decision} ${a.reason}`,
    );
  const runs = async () => events.recentRuns(org, agentId, 50);
  return {
    org,
    agentId,
    store,
    events,
    runtime,
    admin,
    app,
    call,
    subscribe,
    trigger,
    deliver,
    audit,
    runs,
    calls,
    answers,
  };
}

const item = (over: Record<string, unknown> = {}) => ({
  event: "created",
  item: { key: "KAN-1", project: "KAN", type: "Bug" },
  actor: "ann",
  ...over,
});

test(
  "a signed delivery starts the agent once; bad, replayed and repeated ones start nothing",
  { skip: !sql },
  async () => {
    const f = fakeRunner();
    const t = await setup([f.runner]);
    const { subscription, secret } = await t.subscribe();
    await t.trigger();

    assert.equal(await t.deliver(subscription.id, secret!, item(), { delivery: "one" }), 200);
    assert.equal(f.started.length, 1);
    assert.equal(f.started[0].event.type, "created");
    assert.deepEqual(f.started[0].event.data, {
      key: "KAN-1",
      project: "KAN",
      type: "Bug",
      by: "ann",
    });
    assert.match(f.started[0].text, /^Triage it\./);

    // The same delivery again (a retry, or a replay within the window).
    assert.equal(await t.deliver(subscription.id, secret!, item(), { delivery: "one" }), 200);
    // Signed with another secret, or not signed at all.
    assert.equal(await t.deliver(subscription.id, "b".repeat(64), item()), 401);
    assert.equal(await t.deliver(subscription.id, secret!, item(), { sig: "sha256=00" }), 401);
    // Signed, but stamped long ago.
    assert.equal(
      await t.deliver(subscription.id, secret!, item({ ts: Date.now() - 600_000 })),
      200,
    );
    // Not an event the app declares.
    assert.equal(await t.deliver(subscription.id, secret!, item({ event: "exploded" })), 200);
    assert.equal(f.started.length, 1);

    const log = await t.audit();
    assert.ok(log.includes("events:delivery allowed Received"));
    assert.ok(
      log.includes("events:delivery denied Ignored: already received (a retry or a replay)"),
    );
    // Two refusals within a minute: one audit row.
    assert.equal(log.filter((l) => l.includes("signature did not match")).length, 1);
    assert.ok(log.some((l) => l.includes("Refused: sent too long ago")));
    assert.ok(log.some((l) => l.includes("not an event type this app declares")));
    assert.ok(log.some((l) => l.startsWith("events:run allowed Started run")));
    // A subscription that does not exist, and a browser.
    assert.equal(await t.deliver(randomUUID(), secret!, item()), 404);
    const browser = await t.app.fetch(
      new Request(`http://svc/agent-connect/events/${subscription.id}`, {
        method: "POST",
        headers: { origin: "https://evil.example" },
        body: "{}",
      }),
    );
    assert.equal(browser.status, 403);
  },
);

test(
  "limits, access, the app's own changes and conditions decide who is started, each audited",
  { skip: !sql },
  async () => {
    const f = fakeRunner();
    const t = await setup([f.runner]);
    const { subscription, secret } = await t.subscribe();
    await t.trigger({
      eventTypes: ["created", "updated"],
      conditions: [{ field: "type", in: ["Bug", "Task"] }],
    });
    const send = (over: Record<string, unknown>) => t.deliver(subscription.id, secret!, item(over));

    await send({ item: { key: "HELP-1", project: "HELP", type: "Bug" } });
    await send({ item: { key: "X-1", type: "Bug" } });
    await send({ actor: "svc" });
    await send({ item: { key: "KAN-2", project: "KAN", type: "Epic" } });
    assert.equal(f.started.length, 0);
    await send({ event: "updated" });
    assert.equal(f.started.length, 1);

    const log = await t.audit();
    assert.ok(log.some((l) => /Out of limits: projects HELP/.test(l)));
    assert.ok(log.some((l) => /Out of limits: projects \(none\)/.test(l)));
    assert.ok(log.some((l) => /own connection/.test(l)));
    assert.ok(log.some((l) => /Filtered: type does not match/.test(l)));

    // Without the grant the agent is never started.
    await sql!`delete from agent_connect.grants where agent_id = ${t.agentId}`;
    await send({});
    assert.equal(f.started.length, 1);
    assert.ok((await t.audit()).some((l) => /no access to Tracker/.test(l)));
  },
);

test("a trigger's caps hold back runs past them", { skip: !sql }, async () => {
  const f = fakeRunner();
  const t = await setup([f.runner]);
  const { subscription, secret } = await t.subscribe();
  await t.trigger({ maxConcurrent: 1 });
  await t.deliver(subscription.id, secret!, item());
  await t.deliver(subscription.id, secret!, item());
  assert.equal(f.started.length, 1);
  assert.ok((await t.audit()).some((l) => /Capped: runs at once reached/.test(l)));

  // The first ends (with its cost), so the next one runs.
  f.states.set(`s-${f.started[0].runId}`, "idle");
  await t.runtime.watch();
  const [ended] = await t.runs();
  assert.equal(ended.state, "ended");
  assert.equal(ended.costCents, 30);
  await t.deliver(subscription.id, secret!, item());
  assert.equal(f.started.length, 2);
});

test(
  "a poll reads from its cursor, moves it, and starts each change once",
  { skip: !sql },
  async () => {
    const f = fakeRunner();
    const t = await setup([f.runner]);
    const { subscription, secret } = await t.subscribe("poll", { filter: "project = KAN" });
    assert.equal(secret, null);
    assert.equal(subscription.url, null);
    await t.trigger({ eventTypes: ["created", "updated"] });
    const later = (s: number) => new Date(Date.now() + s * 1000).toISOString();
    // KAN-1 is new: created after the subscription, its update stamped just after (as Jira
    // does). KAN-2 was made before and changed since.
    const a = { key: "KAN-1", project: "KAN", created: later(1), updated: later(1.2) };
    const b = { key: "KAN-2", project: "KAN", created: later(-3600), updated: later(2) };
    t.answers.set("/search", { status: 200, data: { items: [a, b] } });

    const sub = async () => (await t.events.subscription(subscription.id))!;
    await t.runtime.poll(await sub());
    await t.runtime.idle();
    assert.deepEqual(
      f.started.map((s) => [s.event.type, s.event.data.key]),
      [
        ["created", "KAN-1"],
        ["updated", "KAN-2"],
      ],
    );
    const search = t.calls.find((c) => c.path === "/search")!;
    assert.match(String((search.body as { q: string }).q), /^\(project = KAN\) and since -\d+m$/);
    assert.equal((await sub()).cursor, b.updated);

    // The overlap brings both back: nothing new starts and nothing is audited as a repeat.
    await t.runtime.poll(await sub());
    await t.runtime.idle();
    assert.equal(f.started.length, 2);
    assert.ok(!(await t.audit()).some((l) => l.includes("already received")));

    // A failing read keeps the cursor and says why.
    t.answers.set("/search", { status: 500, data: null });
    await t.runtime.poll(await sub());
    assert.equal((await sub()).cursor, b.updated);
    assert.equal((await sub()).lastError, "Tracker answered 500");
  },
);

test(
  "end to end: the signed webhook runner posts the event to the receiver, then the run ends",
  { skip: !sql },
  async () => {
    const got: { body: string; sig: string }[] = [];
    const fetch = (async (_url: string, init: RequestInit) => {
      got.push({
        body: String(init.body),
        sig: (init.headers as Record<string, string>)[SIGNATURE_HEADER],
      });
      return new Response(null, { status: 200 });
    }) as never;
    const ctx = {
      teams: { claudeKey: async () => null } as unknown as RunnerContext["teams"],
      key,
      publicUrl: PUBLIC,
    };
    const webhook = signedWebhook(fetch)(ctx);
    const t = await setup([webhook]);
    const { subscription, secret } = await t.subscribe();
    const receiverSecret = "r".repeat(40);
    const saved = await t.trigger({
      runner: "webhook",
      settings: { url: "https://hooks.example/events" },
      signingSecret: receiverSecret,
    });
    assert.equal((saved as { hasSigningSecret: boolean }).hasSigningSecret, true);
    assert.equal(JSON.parse(got[0].body).type, "ping");

    await t.deliver(subscription.id, secret!, item());
    const sent = JSON.parse(got[1].body);
    assert.equal(sent.type, "event");
    assert.equal(sent.app, "tracker");
    assert.equal(sent.event, "created");
    assert.deepEqual(sent.data, { key: "KAN-1", project: "KAN", type: "Bug", by: "ann" });
    assert.equal(sent.mcpUrl, `${PUBLIC}/mcp/${t.agentId}`);
    // What a receiver checks: the HMAC of "<t>.<body>" with its secret.
    const [, ts, v1] = /^t=(\d+),v1=([0-9a-f]{64})$/.exec(got[1].sig)!;
    assert.equal(
      createHmac("sha256", receiverSecret).update(`${ts}.${got[1].body}`).digest("hex"),
      v1,
    );

    await t.runtime.watch();
    const [run] = await t.runs();
    assert.equal(run.state, "ended");
    assert.equal(run.runner, "webhook");
  },
);

test(
  "admin: the secret is shown once, registering falls back to steps, triggers are checked",
  { skip: !sql },
  async () => {
    const f = fakeRunner();
    const t = await setup([f.runner]);
    const created = await t.subscribe("webhook", { register: true, filter: "project = KAN" });
    assert.equal(created.subscription.url, `${PUBLIC}/events/${created.subscription.id}`);
    assert.match(created.secret!, /^[0-9a-f]{64}$/);
    assert.deepEqual(created.setup[0].click, [`Settings > Hooks > ${created.subscription.url}`]);
    assert.equal(created.registration!.ok, false);
    const reg = t.calls.find((c) => c.path === "/hooks")!;
    assert.deepEqual(reg.body, {
      url: created.subscription.url,
      secret: created.secret,
      events: ["created", "updated"],
      filter: "project = KAN",
    });
    const listed = await t.call("GET", "/event-subscriptions");
    assert.ok(!JSON.stringify(listed.body).includes(created.secret!));
    // One per app.
    assert.equal(
      (await t.call("POST", "/event-subscriptions", { app: "tracker", source: "poll", by: "a" }))
        .status,
      400,
    );

    const rotated = await t.call("POST", `/event-subscriptions/${created.subscription.id}/secret`, {
      by: "ann",
    });
    assert.equal(rotated.status, 200);
    assert.notEqual(rotated.body.secret, created.secret);
    assert.equal(await t.deliver(created.subscription.id, created.secret!, item()), 401);

    const bad = (over: Record<string, unknown>) =>
      t.call("PUT", `/agents/${t.agentId}/event-triggers/tracker`, {
        enabled: true,
        eventTypes: ["created"],
        prompt: "x",
        runner: "fake",
        budgetCents: 50,
        by: "ann",
        ...over,
      });
    assert.match((await bad({ eventTypes: ["exploded"] })).body.error, /no event exploded/);
    assert.match(
      (await bad({ conditions: [{ field: "nope", in: ["a"] }] })).body.error,
      /no field nope/,
    );
    assert.match((await bad({ runner: "claude-managed-agents" })).body.error, /no runner/);
    assert.equal((await bad({ budgetCents: 600, maxCentsPerDay: 500 })).status, 400);
    assert.equal(
      (
        await t.call("PUT", `/agents/${randomUUID()}/event-triggers/tracker`, {
          enabled: true,
          eventTypes: ["created"],
          prompt: "x",
          runner: "fake",
          budgetCents: 50,
          by: "ann",
        })
      ).status,
      400,
    );

    await t.trigger();
    assert.equal((await t.call("GET", `/agents/${t.agentId}/event-triggers`)).body.length, 1);
    assert.equal((await t.call("GET", `/agents/${t.agentId}/event-runs`)).status, 200);
    assert.equal(
      (await t.call("DELETE", `/agents/${t.agentId}/event-triggers/tracker`, { by: "ann" })).status,
      204,
    );
    assert.equal(
      (await t.call("DELETE", `/event-subscriptions/${created.subscription.id}`, { by: "ann" }))
        .status,
      204,
    );
    assert.equal(
      (await t.call("DELETE", `/event-subscriptions/${created.subscription.id}`, { by: "ann" }))
        .status,
      404,
    );
    const log = await t.audit();
    for (const action of [
      "subscribe_events",
      "rotate_event_secret",
      "set_event_trigger",
      "remove_event_trigger",
      "unsubscribe_events",
    ])
      assert.ok(
        log.some((l) => l.startsWith(`admin:${action} allowed`)),
        action,
      );
  },
);

test(
  "a webhook registered through the app's API keeps where the app put it",
  { skip: !sql },
  async () => {
    const t = await setup([fakeRunner().runner]);
    t.answers.set("/hooks", { status: 201, data: { self: "https://tracker.example/hooks/7" } });
    const created = await t.subscribe("webhook", { register: true });
    assert.deepEqual(created.registration, { ok: true, message: "Registered in Tracker" });
    assert.equal(
      (created.subscription as { registered?: string }).registered,
      "https://tracker.example/hooks/7",
    );
  },
);
