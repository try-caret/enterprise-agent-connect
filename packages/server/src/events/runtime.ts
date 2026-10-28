import { effectiveCapabilities, type Connector } from "../../../connector-kit/src/connector.js";
import {
  ConnectError,
  type VendorRequest,
  type VendorResponse,
} from "../../../connector-kit/src/types.js";
import { drainer } from "../queue.js";
import { credentialAad, type CredentialKey } from "../secrets.js";
import type { AuditEntry, Store } from "../store.js";
import { capReached, type TriggerRunner } from "../teams/runner.js";
import { readCappedBytes } from "../teams/runtime.js";
import {
  bodyHash,
  eventPrompt,
  failedCondition,
  fill,
  fresh,
  locate,
  outsideLimits,
  pickFields,
  rawAt,
  signed,
  valueAt,
  type EventData,
} from "./deliver.js";
import type { EventTriggerRow, EventsStore, QueuedEvent, SubscriptionRow } from "./store.js";

// Events from connected apps (docs/triggers.md): deliveries at /events/<subscription id> and
// polls, one queue, and the runs they start on the same runners as Teams (teams/runner.ts).

// Larger than any one event an app sends (a Jira issue with its changelog), far below a flood.
const MAX_BODY = 1024 * 1024;
const STUCK_MS = 30 * 60_000;
const VERIFY_TTL_MS = 60_000;
const SELF_TTL_MS = 3600_000;
// One refused delivery per subscription and minute is audited: anyone can post refusals.
const REFUSED_AUDIT_MS = 60_000;

export const subscriptionAad = (org: string, id: string) =>
  credentialAad(org, `event-subscription:${id}`);
export const eventTriggerAad = (org: string, agentId: string, app: string) =>
  credentialAad(org, `event-trigger:${agentId}:${app}`);
const eventAad = (subscriptionId: string, deliveryId: string) =>
  `agent-connect:app-event:${subscriptionId}:${deliveryId}`;

const reply = (status: number) => new Response(null, { status });

export type EventsRuntimeDeps = {
  store: Store;
  events: EventsStore;
  key: CredentialKey;
  // Apps with `events`, by id.
  connectors: Map<string, Connector>;
  // A call with the organization's own connection; null when the app is not connected.
  appCall(org: string, app: string, request: VendorRequest): Promise<VendorResponse | null>;
  runners: Map<string, TriggerRunner>;
};

export function createEventsRuntime(deps: EventsRuntimeDeps) {
  const { events, key, connectors } = deps;

  const audit = (entry: Omit<AuditEntry, "vendorStatus"> & { vendorStatus?: number | null }) =>
    deps.store.audit({ vendorStatus: null, ...entry, reason: entry.reason?.slice(0, 500) ?? null });
  const delivery = (
    sub: SubscriptionRow,
    target: string,
    decision: AuditEntry["decision"],
    reason: string,
  ) =>
    audit({
      organizationId: sub.organizationId,
      agentId: null,
      agentName: null,
      connectorId: sub.connectorId,
      tool: "events:delivery",
      target,
      decision,
      reason,
    });

  const openSecret = (sub: SubscriptionRow) =>
    key.open(
      { ciphertext: sub.secret!, keyId: sub.secretKeyId! },
      subscriptionAad(sub.organizationId, sub.id),
    ).secret;

  // ---- inbound ----

  const refusedAt = new Map<string, number>();

  /** POST /events/:subscription. The proof is checked on the raw body before anything in it
   *  is read; accepted deliveries are queued and answered at once. */
  async function inbound(subscriptionId: string, req: Request): Promise<Response> {
    if (req.headers.has("origin")) return reply(403);
    const sub = await events.subscription(subscriptionId);
    const connector = sub ? connectors.get(sub.connectorId) : undefined;
    const hook = connector?.events?.webhook;
    if (!sub || sub.source !== "webhook" || !hook) return reply(404);
    const raw = await readCappedBytes(req, MAX_BODY);
    if (raw === null) return reply(413);
    if (!signed(hook.verify, openSecret(sub), raw, req.headers)) {
      if (Date.now() - (refusedAt.get(sub.id) ?? 0) > REFUSED_AUDIT_MS) {
        refusedAt.set(sub.id, Date.now());
        await delivery(sub, "unknown", "denied", "Refused: the signature did not match");
      }
      return reply(401);
    }
    let body: unknown;
    try {
      body = JSON.parse(raw.toString("utf8"));
    } catch {
      await delivery(sub, "unknown", "denied", "Refused: the body is not JSON");
      return reply(400);
    }
    const type = locate(hook.type, body, req.headers);
    const deliveryId =
      (hook.deliveryId && locate(hook.deliveryId, body, req.headers)) || bodyHash(raw);
    if (hook.timestamp && !fresh(hook.timestamp, body, req.headers)) {
      await delivery(
        sub,
        `${type} ${deliveryId}`,
        "denied",
        "Refused: sent too long ago (replayed?)",
      );
      return reply(200);
    }
    await accept(sub, connector!, type, deliveryId, body, false);
    return reply(200);
  }

  // Queues one event, once per delivery id. Poll overlaps repeat items by design, so their
  // duplicates are not audited.
  async function accept(
    sub: SubscriptionRow,
    connector: Connector,
    type: string | null,
    deliveryId: string,
    body: unknown,
    quietDuplicates: boolean,
  ) {
    const target = `${type ?? "unknown"} ${deliveryId}`;
    if (!type || !connector.events!.types.some((t) => t.id === type))
      return delivery(sub, target, "denied", "Ignored: not an event type this app declares");
    const data = pickFields(connector.events!.fields, body);
    const sealed = key.seal({ data: JSON.stringify(data) }, eventAad(sub.id, deliveryId));
    const id = await events.insertEvent({
      subscriptionId: sub.id,
      deliveryId,
      eventType: type,
      payload: sealed.ciphertext,
      payloadKeyId: sealed.keyId,
    });
    if (!id) {
      if (!quietDuplicates)
        await delivery(sub, target, "denied", "Ignored: already received (a retry or a replay)");
      return;
    }
    await delivery(sub, target, "allowed", "Received");
    queue.kick();
  }

  // ---- the queue ----

  const queue = drainer("events", async () => {
    const event = await events.claimEvent();
    if (!event) return false;
    try {
      await events.finishEvent(event.id, "done", await handle(event));
    } catch (error) {
      console.error(`[events] event ${event.id} failed`, error);
      await events.finishEvent(event.id, "failed", "Internal error");
    }
    return true;
  });

  async function handle(event: QueuedEvent): Promise<string> {
    const connector = connectors.get(event.connectorId);
    if (!connector?.events || !event.payload) return "app or payload gone";
    const data = JSON.parse(
      key.open(
        { ciphertext: event.payload, keyId: event.payloadKeyId! },
        eventAad(event.subscriptionId, event.deliveryId),
      ).data,
    ) as EventData;
    const triggers = await events.triggersFor(event.organizationId, event.connectorId);
    if (!triggers.length) return "no triggers";
    const self = await selfId(event.organizationId, connector);
    const fromSelf =
      self !== null && (connector.events.self?.fields ?? []).some((f) => data[f] === self);
    let started = 0;
    for (const trigger of triggers) {
      try {
        if (await consider(event, connector, data, trigger, fromSelf)) started++;
      } catch (error) {
        console.error(`[events] trigger ${trigger.id} on event ${event.id}`, error);
      }
    }
    return `${started} run(s) started`;
  }

  // The app connection's own account id (connector `self`), cached an hour.
  const selves = new Map<string, { id: string | null; at: number }>();
  async function selfId(org: string, connector: Connector): Promise<string | null> {
    const self = connector.events?.self;
    if (!self) return null;
    const k = `${org}:${connector.id}`;
    const hit = selves.get(k);
    if (hit && Date.now() - hit.at < SELF_TTL_MS) return hit.id;
    const res = await deps.appCall(org, connector.id, self.request as VendorRequest);
    const id = res && res.status >= 200 && res.status < 300 ? rawAt(res.data, self.id) : null;
    // Unknown, so nothing is mistaken for the app's own change; asked again next event.
    if (typeof id !== "string") return null;
    selves.set(k, { id, at: Date.now() });
    return id;
  }

  /** One trigger's decision on one event, audited whatever it is. True when a run started. */
  async function consider(
    event: QueuedEvent,
    connector: Connector,
    data: EventData,
    trigger: EventTriggerRow,
    fromSelf: boolean,
  ): Promise<boolean> {
    const state = await events.agentState(trigger.agentId, connector.id);
    const decide = async (decision: AuditEntry["decision"], reason: string) => {
      await audit({
        organizationId: event.organizationId,
        agentId: trigger.agentId,
        agentName: state?.agentName ?? "",
        connectorId: connector.id,
        tool: "events:run",
        target: `${event.eventType} ${event.deliveryId}`,
        decision,
        reason,
      });
      return false;
    };
    if (!trigger.eventTypes.includes(event.eventType))
      return decide("denied", "Filtered: not one of the trigger's event types");
    if (fromSelf)
      return decide("denied", `Filtered: made by ${connector.name}'s own connection (an agent)`);
    const capabilities =
      state?.agentStatus === "active" && state.grant && state.connection?.connected
        ? effectiveCapabilities(connector, state.grant.capabilities, state.connection.capabilities)
        : [];
    if (!capabilities.length)
      return decide("denied", `Not started: the agent has no access to ${connector.name}`);
    const outside = outsideLimits(connector.events!.limits, data, state!.grant!.limits);
    if (outside) return decide("denied", `Out of limits: ${outside}`);
    const failed = failedCondition(trigger.conditions, data);
    if (failed) return decide("denied", `Filtered: ${failed.field} does not match`);
    const runner = deps.runners.get(trigger.runner);
    if (!runner?.startEvent)
      return decide("failed", `Not started: this deployment cannot run ${trigger.runner}`);
    const unready = await runner.unready(event.organizationId);
    if (unready) return decide("failed", `Not started: ${unready}`);
    const cap = capReached(await events.runCounts(trigger.id), trigger);
    if (cap)
      return decide(
        "denied",
        `Capped: ${cap === "hour" ? "runs per hour" : cap === "day" ? "spending per day" : "runs at once"} reached`,
      );

    const runId = await events.startRun({
      agentId: trigger.agentId,
      triggerId: trigger.id,
      eventId: event.id,
      eventType: event.eventType,
      deliveryId: event.deliveryId,
      runner: trigger.runner,
    });
    try {
      const secret = trigger.secret
        ? key.open(
            { ciphertext: trigger.secret, keyId: trigger.secretKeyId! },
            eventTriggerAad(event.organizationId, trigger.agentId, connector.id),
          ).secret
        : null;
      await verified(runner, trigger, event.organizationId, secret);
      const type = connector.events!.types.find((t) => t.id === event.eventType)!;
      const sessionId = await runner.startEvent({
        org: event.organizationId,
        agentId: trigger.agentId,
        runId,
        text: eventPrompt({
          prompt: trigger.prompt,
          appName: connector.name,
          type,
          deliveryId: event.deliveryId,
          data,
        }),
        budgetCents: trigger.budgetCents,
        trigger,
        secret,
        event: { app: connector.id, type: type.id, deliveryId: event.deliveryId, data },
      });
      await events.updateRun(runId, { state: "running", sessionId });
      await decide("allowed", `Started run ${runId}`);
      return true;
    } catch (error) {
      const why = error instanceof Error ? error.message : String(error);
      if (!(error instanceof ConnectError)) console.error(`[events] run ${runId}`, error);
      await events.updateRun(runId, { state: "failed", error: why.slice(0, 500) });
      return decide("failed", `Run ${runId} could not start: ${why}`);
    }
  }

  // The runner's check of the trigger (for Claude: the agent uses this agent's own MCP address
  // and nothing else of this service), again before each run, trusted for a minute.
  const verifiedAt = new Map<string, number>();
  async function verified(
    runner: TriggerRunner,
    trigger: EventTriggerRow,
    org: string,
    secret: string | null,
  ) {
    const k = JSON.stringify([
      trigger.id,
      trigger.runner,
      trigger.settings,
      trigger.managedAgentId,
      trigger.vaultIds,
    ]);
    if (Date.now() - (verifiedAt.get(k) ?? 0) < VERIFY_TTL_MS) return;
    try {
      await runner.verify(trigger, org, secret, "answer");
    } catch (error) {
      if (error instanceof ConnectError)
        throw new ConnectError(`trigger check failed: ${error.message}`);
      throw error;
    }
    verifiedAt.set(k, Date.now());
  }

  // ---- polls ----

  /** One read of the app from the subscription's cursor; each new item becomes an event. */
  async function poll(sub: SubscriptionRow) {
    const connector = connectors.get(sub.connectorId);
    const p = connector?.events?.poll;
    if (!p) return events.polled(sub.id, null, `${sub.connectorId} cannot be polled`);
    const cursor = sub.cursor ?? new Date().toISOString();
    const since = Date.parse(cursor);
    const minutes =
      Math.max(
        1,
        Math.ceil((Date.now() - (Number.isFinite(since) ? since : Date.now())) / 60_000),
      ) + 1;
    const request = fill(p.request, {
      minutes: String(minutes),
      cursor,
      filter: sub.filter && p.filter ? p.filter.replace("{value}", sub.filter) : "",
    }) as VendorRequest;
    try {
      const res = await deps.appCall(sub.organizationId, sub.connectorId, request);
      if (!res) return events.polled(sub.id, null, `${connector!.name} is not connected`);
      if (res.status < 200 || res.status >= 300) {
        await audit({
          organizationId: sub.organizationId,
          agentId: null,
          agentName: null,
          connectorId: sub.connectorId,
          tool: "events:poll",
          target: request.path,
          decision: "failed",
          reason: `${connector!.name} answered ${res.status}`,
          vendorStatus: res.status,
        });
        return events.polled(sub.id, null, `${connector!.name} answered ${res.status}`);
      }
      // ponytail: one page per poll; the next continues from the newest item read. More changes
      // than a page within the overlap minute would be read again and again: page then.
      const items = rawAt(res.data, p.items);
      let newest = cursor;
      for (const item of Array.isArray(items) ? items : []) {
        const body = p.wrap ? { [p.wrap]: item } : item;
        const same = (a: string, b: string) =>
          JSON.stringify(valueAt(body, a)) === JSON.stringify(valueAt(body, b));
        const type = p.types.find((t) => !t.same || same(...t.same))?.type ?? null;
        const deliveryId = p.deliveryId.map((path) => String(valueAt(body, path) ?? "")).join("@");
        const at = valueAt(body, p.cursor);
        if (typeof at === "string" && Date.parse(at) > Date.parse(newest)) newest = at;
        await accept(sub, connector!, type, deliveryId, body, true);
      }
      await events.polled(sub.id, newest, null);
    } catch (error) {
      console.error(`[events] poll ${sub.id}`, error);
      await events.polled(
        sub.id,
        null,
        error instanceof ConnectError ? error.message : "Internal error",
      );
    }
  }

  let polling = false;
  async function pollDue() {
    if (polling) return;
    polling = true;
    try {
      for (const sub of await events.duePolls()) await poll(sub);
    } finally {
      polling = false;
    }
  }

  // ---- watching the runs ----

  async function watch() {
    const stale = await events.failStaleStarts();
    if (stale) console.error(`[events] ${stale} run(s) failed after not starting`);
    for (const run of await events.runningRuns()) {
      try {
        const runner = deps.runners.get(run.runner);
        if (!runner) {
          await events.updateRun(run.id, {
            state: "failed",
            error: `Runner ${run.runner} is not offered by this deployment`,
          });
          continue;
        }
        const s = await runner.state(
          { sessionId: run.sessionId, replyCount: 0, awaitsReply: false },
          run.organizationId,
        );
        if ("gone" in s) {
          await events.updateRun(run.id, { state: "failed", error: s.gone });
          continue;
        }
        const costCents = s.costCents ?? undefined;
        if (s.status === "idle" && s.atBudget)
          await events.updateRun(run.id, {
            state: "failed",
            error: "Stopped at its spending limit",
            costCents,
          });
        else if (s.status === "idle" || s.status === "terminated")
          await events.updateRun(run.id, { state: "ended", costCents });
        else if (Date.now() - Date.parse(run.startedAt) > STUCK_MS) {
          await runner
            .stop(run.sessionId!, run.organizationId)
            .catch((e) => console.error(`[events] stopping run ${run.id}`, e));
          await events.updateRun(run.id, {
            state: "failed",
            error: "Stopped after 30 minutes without finishing",
            costCents,
          });
        }
      } catch (error) {
        console.error(`[events] watching run ${run.id}`, error);
      }
    }
  }

  let timers: NodeJS.Timeout[] = [];
  return {
    inbound,
    poll,
    pollDue,
    watch,
    kick: queue.kick,
    idle: queue.idle,
    /** Drains the queue, watches runs and polls what is due, every few seconds. */
    start() {
      queue.start();
      timers = [
        setInterval(() => {
          queue.kick();
          watch().catch((error) => console.error("[events] watch", error));
        }, 5000),
        setInterval(
          () => pollDue().catch((error) => console.error("[events] poll", error)),
          15_000,
        ),
        setInterval(
          () => events.pruneEvents().catch((error) => console.error("[events] prune", error)),
          3600_000,
        ),
      ];
      for (const t of timers) t.unref();
    },
    async stop() {
      for (const t of timers) clearInterval(t);
      await queue.stop();
    },
  };
}

export type EventsRuntime = ReturnType<typeof createEventsRuntime>;
