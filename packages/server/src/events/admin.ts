import { randomBytes, randomUUID } from "node:crypto";
import { z } from "zod";
import { brand } from "../../../connector-kit/src/branding.js";
import type { Connector } from "../../../connector-kit/src/connector.js";
import {
  ConnectError,
  type VendorRequest,
  type VendorResponse,
} from "../../../connector-kit/src/types.js";
import type { CredentialKey } from "../secrets.js";
import { InvalidInput } from "../service.js";
import type { Store } from "../store.js";
import { CLAUDE_RUNNER, type TriggerRunner } from "../teams/runner.js";
import { fill, rawAt } from "./deliver.js";
import { eventTriggerAad, subscriptionAad } from "./runtime.js";
import type { EventTriggerRow, EventsStore, SubscriptionRow } from "./store.js";

// The admin side of app events: each organization's subscription per app (webhook or poll),
// each agent's trigger per app, and the runs they started. docs/triggers.md.

type Steps = NonNullable<NonNullable<Connector["events"]>["webhook"]>["setup"];

const id = z.string().regex(/^[A-Za-z0-9_-]{1,128}$/);
const by = z.string().trim().min(1).max(200);

export const subscriptionSchema = z.object({
  app: z.string().regex(/^[a-z0-9-]+$/),
  source: z.enum(["webhook", "poll"]),
  // In the app's own terms, e.g. Jira JQL: what it sends, or what a poll reads.
  filter: z.string().max(2000).default(""),
  intervalSeconds: z.number().int().min(60).max(86400).default(300),
  // Webhooks: register it through the app's API with the connection, where the app allows.
  register: z.boolean().default(false),
  by,
});

const condition = z.union([
  z.strictObject({ field: z.string(), in: z.array(z.string().max(500)).min(1).max(100) }),
  z.strictObject({ field: z.string(), notIn: z.array(z.string().max(500)).min(1).max(100) }),
]);

export const eventTriggerSchema = z
  .object({
    enabled: z.boolean(),
    eventTypes: z.array(z.string().max(100)).min(1).max(50),
    // All must hold, on the event's fields (the app's `events.fields`).
    conditions: z.array(condition).max(20).default([]),
    // What the agent is asked to do; {app} and {event} are filled in. The event's data follows.
    prompt: z.string().trim().min(1).max(4000),
    runner: z
      .string()
      .regex(/^[a-z][a-z0-9-]{1,40}$/)
      .default(CLAUDE_RUNNER),
    managedAgentId: id.optional(),
    environmentId: id.optional(),
    vaultIds: z.array(id).max(20).default([]),
    settings: z.record(z.string(), z.unknown()).default({}),
    signingSecret: z.string().min(32).max(256).optional(),
    budgetCents: z.number().int().min(1).max(2000),
    maxRunsPerHour: z.number().int().min(1).max(100).default(20),
    maxCentsPerDay: z.number().int().min(1).max(100_000).default(500),
    maxConcurrent: z.number().int().min(1).max(10).default(2),
    by,
  })
  .refine((t) => t.maxCentsPerDay >= t.budgetCents, {
    message: "The daily limit must be at least the limit per run",
    path: ["maxCentsPerDay"],
  });

export function createEventsAdmin(deps: {
  store: Store;
  events: EventsStore;
  key: CredentialKey;
  connectors: Map<string, Connector>;
  appCall(org: string, app: string, request: VendorRequest): Promise<VendorResponse | null>;
  runners: Map<string, TriggerRunner>;
  publicUrl: string | null;
}) {
  const { store, events, key } = deps;

  const log = (org: string, who: string, app: string, action: string, target: string) =>
    store.audit({
      organizationId: org,
      agentId: null,
      agentName: who,
      connectorId: app,
      tool: `admin:${action}`,
      target,
      decision: "allowed",
      reason: null,
      vendorStatus: null,
    });

  function eventsOf(app: string) {
    const connector = deps.connectors.get(app);
    if (!connector?.events) throw new InvalidInput(`${app} has no events agents can start from`);
    return { connector, events: connector.events };
  }

  const urlOf = (sub: SubscriptionRow) =>
    sub.source === "webhook" && deps.publicUrl ? `${deps.publicUrl}/events/${sub.id}` : null;

  function publicSubscription(sub: SubscriptionRow) {
    const { secret: _s, secretKeyId: _k, organizationId: _o, ...rest } = sub;
    return { ...rest, url: urlOf(sub) };
  }

  function publicTrigger(row: EventTriggerRow) {
    const { secret, secretKeyId: _k, ...rest } = row;
    return { ...rest, hasSigningSecret: secret !== null };
  }

  const newSecret = () => randomBytes(32).toString("hex");

  // Live runs of a trigger are stopped before it changes runner or goes.
  async function stopRuns(org: string, triggerId: string, runner?: string) {
    for (const run of await events.runningRuns()) {
      if (run.triggerId !== triggerId || (runner && run.runner !== runner)) continue;
      await deps.runners.get(run.runner)?.stop(run.sessionId!, org);
      await events.updateRun(run.id, { state: "failed", error: "Stopped by an administrator" });
    }
  }

  return {
    async subscriptions(org: string) {
      return (await events.subscriptions(org)).map(publicSubscription);
    },

    /** The secret is in this answer only. */
    async subscribe(org: string, input: z.output<typeof subscriptionSchema>) {
      const { connector, events: declared } = eventsOf(input.app);
      const hook = declared.webhook;
      if (input.source === "webhook" && !hook)
        throw new InvalidInput(
          `${connector.name} sends no webhooks; read it on an interval instead`,
        );
      if (input.source === "poll" && !declared.poll)
        throw new InvalidInput(`${connector.name} cannot be read on an interval; use its webhook`);
      if (!(await store.getConnection(org, input.app)))
        throw new InvalidInput(`Connect ${connector.name} first`);
      if (input.source === "webhook" && !deps.publicUrl)
        throw new ConnectError("This deployment has no public address for the app to send to");
      const subId = randomUUID();
      const secret = input.source === "webhook" ? newSecret() : null;
      const sealed = secret ? key.seal({ secret }, subscriptionAad(org, subId)) : null;
      const sub = await events.createSubscription({
        id: subId,
        organizationId: org,
        connectorId: input.app,
        source: input.source,
        secret: sealed?.ciphertext ?? null,
        secretKeyId: sealed?.keyId ?? null,
        filter: input.filter,
        intervalSeconds: input.intervalSeconds,
        // A poll starts from now: what changed before it is not news.
        cursor: input.source === "poll" ? new Date().toISOString() : null,
        createdBy: input.by,
      });
      if (!sub)
        throw new InvalidInput(`${connector.name} already sends events here; remove that first`);
      const url = urlOf(sub);
      let registration: { ok: boolean; message: string } | null = null;
      if (input.register && hook?.register && url && secret) {
        const res = await deps.appCall(
          org,
          input.app,
          fill(hook.register, {
            name: `${brand().productName} events`,
            url,
            secret,
            events: declared.types.map((t) => t.id),
            filter: input.filter,
          }) as VendorRequest,
        );
        const location = res && res.status < 300 ? rawAt(res.data, hook.register.location) : null;
        if (res && res.status >= 200 && res.status < 300) {
          await events.setRegistered(sub.id, typeof location === "string" ? location : "yes");
          registration = { ok: true, message: `Registered in ${connector.name}` };
        } else
          registration = {
            ok: false,
            message: `${connector.name} did not let its connection register the webhook (${res?.status ?? "not connected"}). Add it by hand with the steps.`,
          };
      }
      await log(org, input.by, input.app, "subscribe_events", `${input.source} ${sub.id}`);
      const fresh = (await events.subscription(sub.id))!;
      return {
        subscription: publicSubscription(fresh),
        secret,
        setup: (input.source === "webhook" ? hook!.setup : []).map((step) =>
          fill(step, { url: url ?? "", filter: input.filter || "(none)" }),
        ) as Steps,
        registration,
      };
    },

    async rotateSecret(org: string, subId: string, who: string) {
      const sub = await events.subscription(subId);
      if (!sub || sub.organizationId !== org || sub.source !== "webhook") return null;
      const secret = newSecret();
      const sealed = key.seal({ secret }, subscriptionAad(org, subId));
      if (!(await events.setSecret(org, subId, sealed.ciphertext, sealed.keyId))) return null;
      await log(org, who, sub.connectorId, "rotate_event_secret", subId);
      return { subscription: publicSubscription(sub), secret };
    },

    async unsubscribe(org: string, subId: string, who: string) {
      const sub = await events.deleteSubscription(org, subId);
      if (sub) await log(org, who, sub.connectorId, "unsubscribe_events", subId);
      return sub !== null;
    },

    async triggers(org: string, agentId: string) {
      return (await events.agentTriggers(org, agentId)).map(publicTrigger);
    },

    // Proven by its runner first, as Teams triggers are.
    async setTrigger(
      org: string,
      agentId: string,
      app: string,
      input: z.output<typeof eventTriggerSchema>,
    ) {
      const { connector, events: declared } = eventsOf(app);
      const { by: who, signingSecret, settings: rawSettings, ...rest } = input;
      const unknownType = rest.eventTypes.find((t) => !declared.types.some((x) => x.id === t));
      if (unknownType) throw new InvalidInput(`${connector.name} has no event ${unknownType}`);
      const unknownField = rest.conditions.find((c) => !(c.field in declared.fields));
      if (unknownField)
        throw new InvalidInput(
          `${connector.name} events have no field ${unknownField.field}; they have ${Object.keys(declared.fields).join(", ")}`,
        );
      const runner = deps.runners.get(rest.runner);
      if (!runner?.startEvent)
        throw new InvalidInput(`This deployment offers no runner ${rest.runner} for app events`);
      const settings = runner.settings.safeParse(rawSettings);
      if (!settings.success) throw new InvalidInput(z.prettifyError(settings.error));
      if (!deps.publicUrl) throw new ConnectError("This deployment has no public address");
      const existing = await events.trigger(org, agentId, app);
      const aad = eventTriggerAad(org, agentId, app);
      const sealed = signingSecret
        ? key.seal({ secret: signingSecret }, aad)
        : existing?.secret
          ? { ciphertext: existing.secret, keyId: existing.secretKeyId! }
          : null;
      const row = {
        ...rest,
        agentId,
        connectorId: app,
        managedAgentId: rest.managedAgentId ?? null,
        environmentId: rest.environmentId ?? null,
        settings: settings.data,
        secret: sealed?.ciphertext ?? null,
        secretKeyId: sealed?.keyId ?? null,
        updatedBy: who,
      };
      const secret = sealed ? key.open(sealed, aad).secret : null;
      await runner.verify(row, org, secret, "save");
      if (existing && existing.runner !== rest.runner)
        await stopRuns(org, existing.id, existing.runner);
      if (!(await events.setTrigger(org, row))) return null;
      await log(
        org,
        who,
        app,
        "set_event_trigger",
        `agent ${agentId} ${rest.enabled ? "on" : "off"} ${rest.eventTypes.join(", ")} runner ${rest.runner} budget ${rest.budgetCents}c`,
      );
      return publicTrigger((await events.trigger(org, agentId, app))!);
    },

    async removeTrigger(org: string, agentId: string, app: string, who: string) {
      const existing = await events.trigger(org, agentId, app);
      if (!existing) return false;
      await stopRuns(org, existing.id);
      if (!(await events.removeTrigger(org, agentId, app))) return false;
      await log(org, who, app, "remove_event_trigger", `agent ${agentId}`);
      return true;
    },

    async runs(org: string, agentId: string, limit: number) {
      return events.recentRuns(org, agentId, limit);
    },
  };
}

export type EventsAdmin = ReturnType<typeof createEventsAdmin>;

/** During a key rotation, moves every webhook and trigger secret still sealed with an old key
 *  onto the current one. */
export async function resealEventSecrets(events: EventsStore, key: CredentialKey) {
  if (!key.previousIds.length) return 0;
  const stale = await events.staleSecrets(key.id);
  let n = 0;
  for (const s of stale.subscriptions) {
    const aad = subscriptionAad(s.organizationId, s.id);
    const sealed = key.seal(key.open({ ciphertext: s.secret, keyId: s.secretKeyId }, aad), aad);
    if (await events.reseal("event_subscriptions", s.id, s.secret, sealed.ciphertext, sealed.keyId))
      n++;
  }
  for (const t of stale.triggers) {
    const aad = eventTriggerAad(t.organizationId, t.agentId, t.connectorId);
    const sealed = key.seal(key.open({ ciphertext: t.secret, keyId: t.secretKeyId }, aad), aad);
    if (await events.reseal("event_triggers", t.id, t.secret, sealed.ciphertext, sealed.keyId)) n++;
  }
  return n;
}
