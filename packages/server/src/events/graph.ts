import type { Connector } from "../../../connector-kit/src/connector.js";
import {
  ConnectError,
  type VendorRequest,
  type VendorResponse,
} from "../../../connector-kit/src/types.js";
import { tokensEqual } from "../secrets.js";
import type { AuditEntry } from "../store.js";
import { readCappedBytes } from "../teams/runtime.js";
import { fill, pickFields, rawAt, valueAt, type EventData } from "./deliver.js";
import type { EventsStore, SubscriptionRow, WatchRow } from "./store.js";

// Microsoft Graph change notifications as an event source (a connector's `events.graph`,
// docs/triggers.md "Outlook mail"): one Graph subscription per watched value (an agent's
// mailbox), notifications at /events/<subscription id> proven by clientState, the item read
// with the connection, and a delta catch-up on the poll scheduler for whatever notifications
// missed. Both paths queue an item under its id, so each starts runs once.

type Graph = NonNullable<NonNullable<Connector["events"]>["graph"]>;

const MAX_BODY = 1024 * 1024;
const MAX_ID = 1024;
// Older items are never news: deliveries are kept 30 days, so a 40-day-old message touched
// again would otherwise be news twice.
const NEWS_MS = 7 * 86400_000;
const DELTA_PAGES = 20;
const SUBSCRIPTIONS = "/v1.0/subscriptions";
// Renewed while a sync runs, so a process that dies holding it is waited for this long at most.
const LEASE_SECONDS = 60;
const LEASE_WAIT_MS = 500;
// How long work waits for another's turn on the subscription: an admin request is answered
// within it.
const TURN_WAIT_MS = 15_000;

/** Another change to the subscription's watches, in this process or another, outlasted the
 *  wait for its turn. */
export class SubscriptionBusy extends ConnectError {
  constructor() {
    super("Another change to this app's events is still being made. Try again in a minute.");
  }
}

export type GraphDeps = {
  events: EventsStore;
  publicUrl: string | null;
  appCall(org: string, app: string, request: VendorRequest): Promise<VendorResponse | null>;
  audit(entry: Omit<AuditEntry, "vendorStatus"> & { vendorStatus?: number | null }): Promise<void>;
  delivery(
    sub: SubscriptionRow,
    target: string,
    decision: AuditEntry["decision"],
    reason: string,
  ): Promise<void>;
  refused(sub: SubscriptionRow, reason: string): Promise<void>;
  accept(
    sub: SubscriptionRow,
    connector: Connector,
    type: string,
    deliveryId: string,
    data: EventData,
    quietDuplicates: boolean,
  ): Promise<void>;
  openSecret(sub: SubscriptionRow): string;
};

const lower = (v: unknown) => String(v).toLowerCase();
const listOf = (v: EventData[string] | undefined) =>
  v === null || v === undefined ? [] : Array.isArray(v) ? v : [v];
const ok = (res: VendorResponse | null) => res !== null && res.status >= 200 && res.status < 300;
const why = (res: VendorResponse | null) => {
  if (!res) return "the app is not connected";
  const message = rawAt(res.data, "error.message");
  return `Microsoft answered ${res.status}${typeof message === "string" ? `: ${message.slice(0, 300)}` : ""}`;
};
// A key with dots in it, as Graph's "@odata.nextLink".
const prop = (data: unknown, name: string) =>
  data && typeof data === "object" ? (data as Record<string, unknown>)[name] : undefined;
// Graph's next and delta links are absolute; only their path and query are kept, so a call
// with one stays on the app's own origin.
const pathOf = (link: string) => {
  const u = new URL(link);
  return u.pathname + u.search;
};

/** The event's data from one item: its kept fields, the `only` flags, and whether the watched
 *  account sent it itself. */
export function graphItem(
  g: Graph,
  fields: Record<string, string | string[]>,
  value: string,
  item: unknown,
) {
  const data = pickFields(fields, { [g.wrap]: item, [g.per]: value });
  for (const [name, from] of Object.entries(g.only)) {
    const got = from.flatMap((f) => listOf(data[f]));
    data[name] = got.length > 0 && got.every((v) => lower(v) === lower(value));
  }
  const own = g.own.some((f) => listOf(data[f]).some((v) => lower(v) === lower(value)));
  return { data, own };
}

export function createGraphSource(deps: GraphDeps) {
  const { events } = deps;
  const urlOf = (sub: SubscriptionRow) => `${deps.publicUrl}/events/${sub.id}`;
  // Under the subscription's lease, `signal` stops the work once another process took it over.
  const call = (sub: SubscriptionRow, request: VendorRequest, signal?: AbortSignal) => {
    signal?.throwIfAborted();
    return deps.appCall(sub.organizationId, sub.connectorId, request);
  };
  const failed = (sub: SubscriptionRow, target: string, reason: string, status?: number) =>
    deps.audit({
      organizationId: sub.organizationId,
      agentId: null,
      agentName: null,
      connectorId: sub.connectorId,
      tool: "events:poll",
      target,
      decision: "failed",
      reason,
      vendorStatus: status ?? null,
    });

  // Notifications are answered at once and worked on after; tests and stop() wait for them.
  const working = new Set<Promise<void>>();
  function later(work: Promise<void>) {
    const p = work
      .catch((error) => console.error("[events] graph notification", error))
      .finally(() => working.delete(p));
    working.add(p);
  }

  /** POST /events/:subscription for a Graph subscription: the validation handshake, or a
   *  batch of notifications, each proven by its clientState. */
  async function inbound(sub: SubscriptionRow, connector: Connector, req: Request) {
    const token = new URL(req.url).searchParams.get("validationToken");
    // Graph's check that the address answers, sent before the subscription exists here, so
    // it proves nothing and starts nothing: the token is echoed as plain text.
    if (token !== null) {
      if (token.length > MAX_ID) return new Response(null, { status: 400 });
      return new Response(token, {
        status: 200,
        headers: {
          "content-type": "text/plain; charset=utf-8",
          "x-content-type-options": "nosniff",
        },
      });
    }
    const raw = await readCappedBytes(req, MAX_BODY);
    if (raw === null) return new Response(null, { status: 413 });
    let items: unknown;
    try {
      items = rawAt(JSON.parse(raw.toString("utf8")), "value");
    } catch {
      items = null;
    }
    if (!Array.isArray(items)) {
      await deps.refused(sub, "Refused: not a batch of Graph notifications");
      return new Response(null, { status: 400 });
    }
    const secret = deps.openSecret(sub);
    const proven = items.filter((n) => {
      const state = rawAt(n, "clientState");
      return typeof state === "string" && tokensEqual(state, secret);
    });
    if (proven.length < items.length)
      await deps.refused(sub, "Refused: a notification's clientState did not match");
    if (!proven.length) return new Response(null, { status: 401 });
    for (const n of proven) later(notified(sub, connector, n));
    return new Response(null, { status: 202 });
  }

  async function notified(sub: SubscriptionRow, connector: Connector, n: unknown) {
    const g = connector.events!.graph!;
    const remoteId = rawAt(n, "subscriptionId");
    const watch =
      typeof remoteId === "string" ? await events.watchByRemote(sub.id, remoteId) : null;
    // A Graph subscription removed here lapses on its own.
    if (!watch) return;
    const lifecycle = rawAt(n, "lifecycleEvent");
    if (typeof lifecycle === "string") {
      await deps.delivery(sub, `${lifecycle} ${watch.value}`, "allowed", `Lifecycle: ${lifecycle}`);
      // A sync may be subscribing this watch right now, and may have replaced it by its turn.
      return inTurn(sub, async (signal) => {
        const now = await events.watchByRemote(sub.id, watch.remoteId!);
        if (!now) return;
        if (lifecycle === "reauthorizationRequired") await subscribe(sub, g, now, signal);
        else if (lifecycle === "subscriptionRemoved") {
          await events.watchSubscribed(now.id, null, null, "Removed by Microsoft");
          await subscribe(sub, g, { ...now, remoteId: null }, signal);
          await catchUp(sub, connector, now, signal);
        } else if (lifecycle === "missed") await catchUp(sub, connector, now, signal);
      });
    }
    await events.watchNotified(watch.id);
    const id = rawAt(n, "resourceData.id");
    if (typeof id !== "string" || !id || id.length > MAX_ID)
      return deps.delivery(sub, `unknown ${watch.value}`, "denied", "Ignored: no item id");
    const res = await call(sub, {
      ...g.read,
      path: fill(g.read.path, {
        value: encodeURIComponent(watch.value),
        id: encodeURIComponent(id),
      }) as string,
    } as VendorRequest);
    // Gone already (moved or deleted): nothing to start on. Any other failure is read again
    // by the next catch-up.
    if (res?.status === 404)
      return deps.delivery(sub, `${g.type} ${id}`, "denied", "Ignored: the item is gone");
    if (!ok(res)) return failed(sub, `read ${watch.value}`, why(res), res?.status);
    await offer(sub, connector, watch, res!.data, false);
  }

  async function offer(
    sub: SubscriptionRow,
    connector: Connector,
    watch: WatchRow,
    item: unknown,
    quiet: boolean,
  ) {
    const g = connector.events!.graph!;
    const body = { [g.wrap]: item, [g.per]: watch.value };
    const received = Date.parse(String(valueAt(body, g.received)));
    // Not news: from before the watch started (a message moved back in), or too old.
    if (!(received >= Date.parse(watch.createdAt) && received > Date.now() - NEWS_MS)) return;
    const deliveryId = valueAt(body, g.deliveryId);
    if (typeof deliveryId !== "string" || !deliveryId) return;
    const { data, own } = graphItem(g, connector.events!.fields, watch.value, item);
    if (own) {
      if (!quiet)
        await deps.delivery(
          sub,
          `${g.type} ${deliveryId}`,
          "denied",
          `Ignored: sent by ${watch.value} itself`,
        );
      return;
    }
    await deps.accept(sub, connector, g.type, deliveryId, data, quiet);
  }

  /** Makes the watch's Graph subscription, or renews it (which also reauthorizes it). */
  async function subscribe(
    sub: SubscriptionRow,
    g: Graph,
    watch: WatchRow,
    signal: AbortSignal,
  ): Promise<number | undefined> {
    const expires = new Date(Date.now() + g.maxMinutes * 60_000).toISOString();
    if (watch.remoteId) {
      const res = await call(
        sub,
        {
          method: "PATCH",
          path: `${SUBSCRIPTIONS}/${encodeURIComponent(watch.remoteId)}`,
          body: { expirationDateTime: expires },
        },
        signal,
      );
      if (ok(res)) {
        await events.watchSubscribed(watch.id, watch.remoteId, expires, null);
        return undefined;
      }
      if (res?.status !== 404) {
        await events.watchFailed(watch.id, `Renewing: ${why(res)}`);
        return res?.status;
      }
    }
    const url = urlOf(sub);
    const res = await call(
      sub,
      {
        method: "POST",
        path: SUBSCRIPTIONS,
        body: {
          changeType: g.changeType,
          notificationUrl: url,
          lifecycleNotificationUrl: url,
          resource: fill(g.resource, { value: watch.value }),
          expirationDateTime: expires,
          clientState: deps.openSecret(sub),
        },
      },
      signal,
    );
    const remoteId = ok(res) ? rawAt(res!.data, "id") : null;
    if (typeof remoteId !== "string") {
      await events.watchSubscribed(watch.id, null, null, `Subscribing: ${why(res)}`);
      return res?.status;
    }
    const at = rawAt(res!.data, "expirationDateTime");
    await events.watchSubscribed(watch.id, remoteId, typeof at === "string" ? at : expires, null);
    return undefined;
  }

  /** Reads the delta from the watch's cursor (from its start the first time) and offers each
   *  item; what notifications already queued is a repeat. */
  async function catchUp(
    sub: SubscriptionRow,
    connector: Connector,
    watch: WatchRow,
    signal: AbortSignal,
  ): Promise<number | undefined> {
    const g = connector.events!.graph!;
    let link =
      watch.cursor ??
      (fill(g.delta.path, {
        value: encodeURIComponent(watch.value),
        since: watch.createdAt,
      }) as string);
    for (let page = 0; page < DELTA_PAGES; page++) {
      const res = await call(sub, { method: g.delta.method, path: link } as VendorRequest, signal);
      // A delta link Graph no longer knows: start again from the watch's start.
      if (res?.status === 410) {
        await events.watchCaughtUp(watch.id, null);
        return undefined;
      }
      if (!ok(res)) {
        await events.watchFailed(watch.id, `Catching up: ${why(res)}`);
        await failed(sub, `delta ${watch.value}`, why(res), res?.status);
        return res?.status;
      }
      const items = rawAt(res!.data, "value");
      for (const item of Array.isArray(items) ? items : [])
        await offer(sub, connector, watch, item, true);
      const next = prop(res!.data, "@odata.nextLink");
      const delta = prop(res!.data, "@odata.deltaLink");
      const after = typeof next === "string" ? next : typeof delta === "string" ? delta : null;
      if (!after) {
        await events.watchFailed(watch.id, "Catching up: Microsoft gave no next or delta link");
        await failed(sub, `delta ${watch.value}`, "No next or delta link");
        return undefined;
      }
      link = pathOf(after);
      if (typeof next !== "string") break;
    }
    // A round longer than the pages read continues from its next link next time.
    await events.watchCaughtUp(watch.id, link);
    return undefined;
  }

  // The values of the `per` limit agents with a trigger on hold now.
  async function wanted(sub: SubscriptionRow, g: Graph) {
    const raw = await events.watchedValues(sub.organizationId, sub.connectorId, g.per);
    return new Set(
      raw
        .flatMap((v) => (Array.isArray(v) ? v : [v]))
        .filter((v): v is string => typeof v === "string" && v.length > 0),
    );
  }

  async function release(sub: SubscriptionRow, watch: WatchRow, signal: AbortSignal) {
    if (watch.remoteId) {
      const res = await call(
        sub,
        { method: "DELETE", path: `${SUBSCRIPTIONS}/${encodeURIComponent(watch.remoteId)}` },
        signal,
      );
      // It lapses at its expiry anyway, and its notifications find no watch here.
      if (!ok(res) && res?.status !== 404)
        await failed(sub, `unsubscribe ${watch.value}`, why(res), res?.status);
    }
    await events.removeWatch(watch.id);
  }

  /** Watches exactly the values agents with a trigger hold, renews subscriptions past half
   *  their life, and catches each one up. */
  /** Watches exactly the values agents with a trigger hold, renews subscriptions past half their
   *  life, and catches each one up. Answers the first failing watch's error and the HTTP status of
   *  this round's first refused call. */
  async function syncOnce(sub: SubscriptionRow, connector: Connector, signal: AbortSignal) {
    const g = connector.events!.graph!;
    const want = await wanted(sub, g);
    for (const w of await events.watches(sub.id))
      if (!want.has(w.value)) await release(sub, w, signal);
    for (const v of want) await events.addWatch(sub.id, v);
    const halfLife = (g.maxMinutes * 60_000) / 2;
    let status: number | undefined;
    for (const w of await events.watches(sub.id)) {
      if (!w.remoteId || !w.expiresAt || Date.parse(w.expiresAt) - Date.now() < halfLife) {
        const refused = await subscribe(sub, g, w, signal);
        status ??= refused;
      }
      const refused = await catchUp(sub, connector, w, signal);
      status ??= refused;
    }
    const errors = (await events.watches(sub.id)).filter((w) => w.lastError);
    return {
      error: errors.length ? `${errors[0].value}: ${errors[0].lastError}` : null,
      status,
    };
  }

  // One change to a subscription's watches at a time, in this process and across them: the poll
  // scheduler, a trigger change, a lifecycle notification and an unsubscribe can each start one,
  // on any process, and two would each make a Graph subscription for the same watch, the one
  // refused clearing the other's id. A sync waits on Graph for as long as it needs, so a lease
  // (each call its own), not a lock held on one of the pool's few connections. One that loses it
  // (stalled past its term, another took it) makes no further Graph call. Waiting for the turn
  // gives up after TURN_WAIT_MS with SubscriptionBusy.
  async function inTurn<T>(sub: SubscriptionRow, work: (signal: AbortSignal) => Promise<T>) {
    const until = Date.now() + TURN_WAIT_MS;
    for (;;) {
      const done = await events.withLease(`graph-sync:${sub.id}`, LEASE_SECONDS, work);
      if (done) return done.value;
      if (Date.now() >= until) throw new SubscriptionBusy();
      await new Promise((r) => setTimeout(r, LEASE_WAIT_MS));
    }
  }

  return {
    inbound,
    /** On the poll scheduler and after trigger changes. One started while another runs waits
     *  for it, then runs on what changed meanwhile; SubscriptionBusy past the wait. */
    sync: (sub: SubscriptionRow, connector: Connector) =>
      inTurn(sub, (signal) => syncOnce(sub, connector, signal)),
    /** Removes every watch's Graph subscription, then `remove`s the organization's, in its
     *  turn: a sync on any process waits for it, so none makes a Graph subscription meanwhile
     *  that nothing would delete. SubscriptionBusy (the admin's "try again") past the wait. */
    releaseAll: <T>(sub: SubscriptionRow, remove: () => Promise<T>) =>
      inTurn(sub, async (signal) => {
        for (const w of await events.watches(sub.id)) await release(sub, w, signal);
        return remove();
      }),
    async settle() {
      while (working.size) await Promise.all([...working]);
    },
  };
}
