import { randomUUID } from "node:crypto";
import { effectiveCapabilities, type Connector } from "../../../connector-kit/src/connector.js";
import {
  ConnectError,
  type VendorRequest,
  type VendorResponse,
} from "../../../connector-kit/src/types.js";
import { drainer } from "../queue.js";
import { credentialAad, type CredentialKey } from "../secrets.js";
import type { AuditEntry, Store } from "../store.js";
import {
  nextAttempt,
  RETRY_WINDOW_MS,
  RetryLater,
  type EventRunInput,
  type TriggerRunner,
} from "../teams/runner.js";
import { readCappedBytes } from "../teams/runtime.js";
import {
  bodyHash,
  eventPrompt,
  threadId,
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
import { createGraphSource, SubscriptionBusy } from "./graph.js";
import { noReport, type Report } from "../report.js";
import type {
  AgentAppState,
  EventRetryRow,
  EventTriggerRow,
  EventsStore,
  QueuedEvent,
  SubscriptionRow,
} from "./store.js";

// Events from connected apps (docs/triggers.md): deliveries at /events/<subscription id>, Graph
// notifications (graph.ts) and polls, one queue, and the runs they start on the same runners
// as Teams (teams/runner.ts).

// Larger than any one event an app sends (a Jira issue with its changelog), far below a flood.
const MAX_BODY = 1024 * 1024;
const STUCK_MS = 30 * 60_000;
const VERIFY_TTL_MS = 60_000;
const SELF_TTL_MS = 3600_000;
// A pass holds the watch, renewed while it lasts and dropped after; one that dies mid-pass holds
// the other processes' passes back this long at most.
const WATCH_LEASE_SECONDS = 15;
// One refused delivery per subscription and minute is audited, in each process: anyone can
// post refusals.
const REFUSED_AUDIT_MS = 60_000;

export const subscriptionAad = (org: string, id: string) =>
  credentialAad(org, `event-subscription:${id}`);
export const eventTriggerAad = (org: string, agentId: string, app: string) =>
  credentialAad(org, `event-trigger:${agentId}:${app}`);
const eventAad = (subscriptionId: string, deliveryId: string) =>
  `agent-connect:app-event:${subscriptionId}:${deliveryId}`;
const deliveryAad = (runId: string) => `agent-connect:event-delivery:${runId}`;

// What a run's next hand-over to its runner sends again, sealed on the run meanwhile.
type Delivery = Pick<EventRunInput, "deliveryId" | "text" | "event">;
// The run being handed over, as its audit rows name it.
type RunContext = {
  runId: string;
  org: string;
  agentName: string;
  connectorId: string;
  eventType: string;
  deliveryId: string;
};

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
  // Where Graph sends notifications.
  publicUrl: string | null;
  report?: Report;
  // How long a start the runner did not take is tried again (runner.ts nextAttempt).
  retryWindowMs?: number;
};

export function createEventsRuntime(deps: EventsRuntimeDeps) {
  const { events, key, connectors } = deps;
  const report = deps.report ?? noReport;
  const retryWindowMs = deps.retryWindowMs ?? RETRY_WINDOW_MS;
  // A subscription's poll or sync failing: reported when it starts failing, not on every retry.
  const polled = (
    sub: SubscriptionRow,
    cursor: string | null,
    error: string | null,
    cause: unknown = null,
    status?: number,
  ) => {
    if (error && !sub.lastError)
      report("event_poll_failed", cause, {
        app: sub.connectorId,
        org: sub.organizationId,
        subscription: sub.id,
        source: sub.source,
        status,
      });
    return events.polled(sub.id, cursor, error);
  };

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
  async function refused(sub: SubscriptionRow, reason: string) {
    if (Date.now() - (refusedAt.get(sub.id) ?? 0) <= REFUSED_AUDIT_MS) return;
    refusedAt.set(sub.id, Date.now());
    await delivery(sub, "unknown", "denied", reason);
  }

  /** POST /events/:subscription. The proof is checked on the raw body before anything in it
   *  is read; accepted deliveries are queued and answered at once. */
  async function inbound(subscriptionId: string, req: Request): Promise<Response> {
    if (req.headers.has("origin")) return reply(403);
    const sub = await events.subscription(subscriptionId);
    const connector = sub ? connectors.get(sub.connectorId) : undefined;
    if (sub?.source === "graph" && connector?.events?.graph)
      return graph.inbound(sub, connector, req);
    const hook = connector?.events?.webhook;
    if (!sub || sub.source !== "webhook" || !hook) return reply(404);
    const raw = await readCappedBytes(req, MAX_BODY);
    if (raw === null) return reply(413);
    if (!signed(hook.verify, openSecret(sub), raw, req.headers)) {
      await refused(sub, "Refused: the signature did not match");
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
    const handshake = hook.handshake;
    if (handshake && locate(handshake.at, body, req.headers) === handshake.equals) {
      const echo = valueAt(body, handshake.echo);
      if (typeof echo !== "string") return reply(400);
      await delivery(sub, handshake.equals, "allowed", "Answered the app's check of the address");
      return new Response(echo, { status: 200, headers: { "content-type": "text/plain" } });
    }
    await accept(
      sub,
      connector!,
      type,
      deliveryId,
      pickFields(connector!.events!.fields, body),
      false,
    );
    return reply(200);
  }

  // Queues one event, once per delivery id. Poll overlaps repeat items by design, so their
  // duplicates are not audited.
  async function accept(
    sub: SubscriptionRow,
    connector: Connector,
    type: string | null,
    deliveryId: string,
    data: EventData,
    quietDuplicates: boolean,
  ) {
    const target = `${type ?? "unknown"} ${deliveryId}`;
    if (!type || !connector.events!.types.some((t) => t.id === type))
      return delivery(sub, target, "denied", "Ignored: not an event type this app declares");
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

  const graph = createGraphSource({
    events,
    publicUrl: deps.publicUrl,
    appCall: deps.appCall,
    audit,
    delivery,
    refused,
    accept,
    openSecret,
  });

  // ---- the queue ----

  const queue = drainer("events", async () => {
    const event = await events.claimEvent();
    if (!event) return false;
    try {
      await events.finishEvent(event.id, "done", await handle(event));
    } catch (error) {
      console.error(`[events] event ${event.id} failed`, error);
      report("event_failed", error, {
        app: event.connectorId,
        org: event.organizationId,
        subscription: event.subscriptionId,
      });
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
    const fields = connector.events.self?.fields ?? [];
    // Asked only when the event says who made it (a poll's records do not).
    const self = fields.some((f) => data[f] != null)
      ? await selfId(event.organizationId, connector)
      : null;
    const fromSelf =
      self instanceof Error ? self : self !== null && fields.some((f) => data[f] === self);
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

  // The app connection's own account id (connector `self`), cached an hour. An Error when the
  // app does not say: no run starts then, since an agent's own change would start it again.
  const selves = new Map<string, { id: string; at: number }>();
  async function selfId(org: string, connector: Connector): Promise<string | Error> {
    const self = connector.events!.self!;
    const k = `${org}:${connector.id}`;
    const hit = selves.get(k);
    if (hit && Date.now() - hit.at < SELF_TTL_MS) return hit.id;
    const res = await deps.appCall(org, connector.id, self.request as VendorRequest);
    const id = res && res.status >= 200 && res.status < 300 ? rawAt(res.data, self.id) : null;
    if (typeof id !== "string")
      return new Error(
        `${connector.name} did not say which account the connection is (${res ? `it answered ${res.status}` : "not connected"}), so an agent's own change cannot be told apart. Check the connection has every scope its setup lists.`,
      );
    selves.set(k, { id, at: Date.now() });
    return id;
  }

  /** Why the event may not start a run of the trigger, as the agent's access and the trigger
   *  stand now; null when it may. */
  async function refusal(
    org: string,
    connector: Connector,
    eventType: string,
    data: EventData,
    trigger: EventTriggerRow,
    state: AgentAppState | null,
  ): Promise<[AuditEntry["decision"], string] | null> {
    if (!trigger.enabled) return ["denied", "Not started: the trigger is off"];
    if (!trigger.eventTypes.includes(eventType))
      return ["denied", "Filtered: not one of the trigger's event types"];
    const capabilities =
      state?.agentStatus === "active" && state.grant && state.connection?.connected
        ? effectiveCapabilities(connector, state.grant.capabilities, state.connection.capabilities)
        : [];
    if (!capabilities.length)
      return ["denied", `Not started: the agent has no access to ${connector.name}`];
    const outside = outsideLimits(connector.events!.limits, data, state!.grant!.limits);
    if (outside) return ["denied", `Out of limits: ${outside}`];
    const failed = failedCondition(trigger.conditions, data);
    if (failed) return ["denied", `Filtered: ${failed.field} does not match`];
    const runner = deps.runners.get(trigger.runner);
    if (!runner?.startEvent)
      return ["failed", `Not started: this deployment cannot run ${trigger.runner}`];
    const unready = await runner.unready(org);
    if (unready) return ["failed", `Not started: ${unready}`];
    return null;
  }

  /** One trigger's decision on one event, audited whatever it is. True when a run started. */
  async function consider(
    event: QueuedEvent,
    connector: Connector,
    data: EventData,
    trigger: EventTriggerRow,
    fromSelf: boolean | Error,
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
    if (fromSelf instanceof Error) return decide("failed", `Not started: ${fromSelf.message}`);
    if (fromSelf)
      return decide("denied", `Filtered: made by ${connector.name}'s own connection (an agent)`);
    const org = event.organizationId;
    const refused = await refusal(org, connector, event.eventType, data, trigger, state);
    if (refused) return decide(...refused);
    const runner = deps.runners.get(trigger.runner)!;

    const threadField = connector.events!.thread;
    const threadValue = threadField ? data[threadField] : null;
    const thread =
      threadValue === null || threadValue === undefined || Array.isArray(threadValue)
        ? null
        : threadId(trigger.id, String(threadValue));
    const runTags = {
      app: connector.id,
      org: event.organizationId,
      agent: trigger.agentId,
      runner: trigger.runner,
    };
    // A run whose row cannot even be written is reported here; any later failure, once, by
    // failStart.
    const admitted = await events
      .admitRun(
        {
          agentId: trigger.agentId,
          triggerId: trigger.id,
          eventId: event.id,
          eventType: event.eventType,
          deliveryId: event.deliveryId,
          thread,
          runner: trigger.runner,
        },
        trigger,
      )
      .catch((error: unknown) => {
        report("event_run_start_failed", error, runTags);
        throw error;
      });
    if ("cap" in admitted) {
      const { cap } = admitted;
      return decide(
        "denied",
        `Capped: ${cap === "hour" ? "runs per hour" : cap === "day" ? "spending per day" : "runs at once"} reached`,
      );
    }
    const { runId } = admitted;
    const type = connector.events!.types.find((t) => t.id === event.eventType)!;
    const run = {
      runId,
      org: event.organizationId,
      agentName: state?.agentName ?? "",
      connectorId: connector.id,
      eventType: event.eventType,
      deliveryId: event.deliveryId,
    };
    const sent: Delivery = {
      deliveryId: randomUUID(),
      text: eventPrompt({
        prompt: trigger.prompt,
        appName: connector.name,
        type,
        deliveryId: event.deliveryId,
        data,
      }),
      event: { app: connector.id, type: type.id, deliveryId: event.deliveryId, thread, data },
    };
    return handOver(run, runner, trigger, sent, 1, Date.now());
  }

  const runAudit =
    (run: RunContext, agentId: string) =>
    async (decision: AuditEntry["decision"], reason: string) =>
      audit({
        organizationId: run.org,
        agentId,
        agentName: run.agentName,
        connectorId: run.connectorId,
        tool: "events:run",
        target: `${run.eventType} ${run.deliveryId}`,
        decision,
        reason,
      });

  /** Hands the event to the runner: started (true), tried again later while the runner may
   *  yet take it (RetryLater, within the retry window), or failed. */
  async function handOver(
    run: RunContext,
    runner: TriggerRunner,
    trigger: EventTriggerRow,
    sent: Delivery,
    attempt: number,
    since: number,
  ): Promise<boolean> {
    const after = attempt > 1 ? ` after ${attempt} attempts` : "";
    try {
      const secret = trigger.secret
        ? key.open(
            { ciphertext: trigger.secret, keyId: trigger.secretKeyId! },
            eventTriggerAad(run.org, trigger.agentId, run.connectorId),
          ).secret
        : null;
      await verified(runner, trigger, run.org, secret);
      const sessionId = await runner.startEvent!({
        org: run.org,
        agentId: trigger.agentId,
        runId: run.runId,
        text: sent.text,
        budgetCents: trigger.budgetCents,
        trigger,
        secret,
        deliveryId: sent.deliveryId,
        event: sent.event,
      });
      await events.updateRun(run.runId, { state: "running", sessionId });
      await runAudit(run, trigger.agentId)("allowed", `Started run ${run.runId}${after}`);
      return true;
    } catch (error) {
      const at =
        error instanceof RetryLater
          ? nextAttempt(attempt, since, retryWindowMs, error.retryAfterMs)
          : null;
      if (at === null) {
        await failStart(run, trigger, error, attempt);
        return false;
      }
      console.warn(
        `[events] run ${run.runId} attempt ${attempt}: ${(error as Error).message}; again later`,
      );
      await events.scheduleRetry(
        run.runId,
        attempt,
        new Date(at),
        key.seal({ delivery: JSON.stringify(sent) }, deliveryAad(run.runId)),
      );
      return false;
    }
  }

  // A run that could not start: failed with the reason, audited and reported once.
  async function failStart(
    run: RunContext,
    trigger: Pick<EventTriggerRow, "agentId" | "runner">,
    error: unknown,
    attempts: number,
  ) {
    const why = error instanceof Error ? error.message : String(error);
    if (!(error instanceof ConnectError)) console.error(`[events] run ${run.runId}`, error);
    report("event_run_start_failed", error, {
      app: run.connectorId,
      org: run.org,
      agent: trigger.agentId,
      runner: trigger.runner,
    });
    await events.updateRun(run.runId, { state: "failed", error: why.slice(0, 500) });
    const after = attempts > 1 ? ` after ${attempts} attempts` : "";
    const reason = `Run ${run.runId} could not start: ${why}${after}`;
    await runAudit(run, trigger.agentId)("failed", reason);
  }

  const retries = drainer("events-retries", async () => {
    const row = await events.claimRetry();
    if (!row) return false;
    // Whatever stops a hand-over fails its run: left starting, it would be claimed again forever.
    await retry(row)
      .catch((error) => failStart(runOf(row, ""), row, error, row.attempts))
      .catch((error) => console.error(`[events] retrying run ${row.id}`, error));
    return true;
  });

  const runOf = (row: EventRetryRow, agentName: string): RunContext => ({
    runId: row.id,
    org: row.organizationId,
    agentName,
    connectorId: row.connectorId,
    eventType: row.eventType,
    deliveryId: row.deliveryId,
  });

  // The next hand-over of a run the runner did not take, as the trigger stands now.
  async function retry(row: EventRetryRow) {
    const sent = JSON.parse(
      key.open({ ciphertext: row.delivery, keyId: row.deliveryKeyId }, deliveryAad(row.id))
        .delivery,
    ) as Delivery;
    const [trigger, state] = await Promise.all([
      events.triggerById(row.triggerId),
      events.agentState(row.agentId, row.connectorId),
    ]);
    const run = runOf(row, state?.agentName ?? "");
    const fail = (why: string) => failStart(run, row, new ConnectError(why), row.attempts);
    const since = Date.parse(row.startedAt);
    if (Date.now() > since + retryWindowMs) return fail("The runner did not take it in time");
    const connector = connectors.get(row.connectorId);
    if (!connector?.events) return fail(`${row.connectorId} is no longer offered`);
    if (!trigger || trigger.runner !== row.runner)
      return fail("Its trigger was turned off or changed runner");
    // Access may have changed while it waited. The caps were checked when it first came, and it
    // has held its place among them since, so they are not counted again.
    const data = sent.event.data as EventData;
    const refused = await refusal(run.org, connector, row.eventType, data, trigger, state);
    if (refused) return fail(`No longer allowed: ${refused[1]}`);
    const runner = deps.runners.get(row.runner)!;
    await handOver(run, runner, trigger, sent, row.attempts + 1, since);
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
    if (!p) return polled(sub, null, `${sub.connectorId} cannot be polled`);
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
      if (!res) return polled(sub, null, `${connector!.name} is not connected`);
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
        return polled(sub, null, `${connector!.name} answered ${res.status}`, null, res.status);
      }
      // ponytail: one page per poll; the next continues from the newest item read. More changes
      // than a page within the overlap minute would be read again and again: page then.
      const items = rawAt(res.data, p.items);
      let newest = cursor;
      for (const item of Array.isArray(items) ? items : []) {
        const body = p.wrap ? { [p.wrap]: item } : item;
        const after = (path: string) => {
          const at = valueAt(body, path);
          return typeof at === "string" && Date.parse(at) > since;
        };
        const type = p.types.find((t) => !t.since || after(t.since))?.type ?? null;
        const deliveryId = p.deliveryId.map((path) => String(valueAt(body, path) ?? "")).join("@");
        const at = valueAt(body, p.cursor);
        if (typeof at === "string" && Date.parse(at) > Date.parse(newest)) newest = at;
        await accept(
          sub,
          connector!,
          type,
          deliveryId,
          pickFields(connector!.events!.fields, body),
          true,
        );
      }
      await events.polled(sub.id, newest, null);
    } catch (error) {
      console.error(`[events] poll ${sub.id}`, error);
      const why = error instanceof ConnectError ? error.message : "Internal error";
      await polled(sub, null, why, error);
    }
  }

  /** A Graph subscription's watches brought in line with the triggers, renewed and caught up.
   *  False, and nothing recorded, when another change to them outlasted the wait. */
  async function sync(sub: SubscriptionRow): Promise<boolean> {
    const connector = connectors.get(sub.connectorId);
    if (!connector?.events?.graph) {
      await polled(sub, null, `${sub.connectorId} sends no Graph notifications`);
      return true;
    }
    try {
      const { error, status } = await graph.sync(sub, connector);
      await polled(sub, null, error, null, status);
    } catch (error) {
      if (error instanceof SubscriptionBusy) return false;
      console.error(`[events] graph sync ${sub.id}`, error);
      const why = error instanceof ConnectError ? error.message : "Internal error";
      await polled(sub, null, why, error);
    }
    return true;
  }

  let polling = false;
  async function pollDue() {
    if (polling) return;
    polling = true;
    try {
      let sub;
      while ((sub = await events.claimDuePoll()))
        if (sub.source === "graph") await sync(sub);
        else await poll(sub);
    } finally {
      polling = false;
    }
  }

  // ---- watching the runs ----

  async function watch() {
    const stale = await events.failStaleStarts();
    if (stale) console.error(`[events] ${stale} run(s) failed after not starting`);
    for (const run of await events.runningRuns()) {
      // Passes can overlap (one whose lease another process took over): what one read is
      // written only if the run is still running.
      const finish = (patch: Parameters<typeof events.updateRun>[1]) =>
        events.updateRun(run.id, patch, "running");
      try {
        const runner = deps.runners.get(run.runner);
        if (!runner) {
          await finish({
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
          await finish({ state: "failed", error: s.gone });
          continue;
        }
        const costCents = s.costCents ?? undefined;
        if (s.status === "idle" && s.atBudget)
          await finish({ state: "failed", error: "Stopped at its spending limit", costCents });
        else if (s.status === "idle" || s.status === "terminated")
          await finish({ state: "ended", costCents });
        else if (Date.now() - Date.parse(run.startedAt) > STUCK_MS) {
          await runner
            .stop(run.sessionId!, run.organizationId)
            .catch((e) => console.error(`[events] stopping run ${run.id}`, e));
          await finish({
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

  /** One pass at a time, in this process and across them, however long one lasts. */
  let watching = false;
  async function watchTurn() {
    if (watching) return;
    watching = true;
    try {
      await events.withLease("events-watch", WATCH_LEASE_SECONDS, () => watch());
    } finally {
      watching = false;
    }
  }

  let timers: NodeJS.Timeout[] = [];
  return {
    inbound,
    poll,
    pollDue,
    /** After a trigger on the app changes: the organization's Graph subscription for it, if
     *  any, follows at once rather than at its next interval; while another change to it
     *  outlasts the wait, at the scheduler's next round. */
    async syncApp(org: string, app: string) {
      const sub = await events.subscriptionFor(org, app);
      if (sub?.source === "graph" && !(await sync(sub))) await events.pollSoon(sub.id);
    },
    releaseAll: graph.releaseAll,
    settle: graph.settle,
    watch,
    watchTurn,
    kick: queue.kick,
    idle: queue.idle,
    /** Drains the queue, watches runs and polls what is due, every few seconds. */
    start() {
      queue.start();
      retries.start();
      timers = [
        setInterval(() => {
          queue.kick();
          retries.kick();
          watchTurn().catch((error) => console.error("[events] watch", error));
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
      await graph.settle();
      await Promise.all([queue.stop(), retries.stop()]);
    },
    /** Hands over every run whose next attempt is due, now (tests). */
    async retryDue() {
      retries.kick();
      await retries.idle();
    },
  };
}

export type EventsRuntime = ReturnType<typeof createEventsRuntime>;
