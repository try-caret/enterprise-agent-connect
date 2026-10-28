import { tables, type Sql } from "../store.js";
import type { RunCounts } from "../teams/runner.js";

// Postgres for app events: each organization's subscriptions, the inbound queue, each agent's
// event triggers and the runs they started (migrations/0007_app_events.sql).

const iso = (column: string) =>
  `to_char(${column} at time zone 'UTC', 'YYYY-MM-DD"T"HH24:MI:SS.MS"Z"')`;

export type SubscriptionRow = {
  id: string;
  organizationId: string;
  connectorId: string;
  source: "webhook" | "poll";
  secret: Buffer | null;
  secretKeyId: string | null;
  filter: string;
  intervalSeconds: number;
  cursor: string | null;
  polledAt: string | null;
  lastError: string | null;
  registered: string | null;
  createdBy: string;
  createdAt: string;
};

export type Condition = { field: string; in: string[] } | { field: string; notIn: string[] };

export type EventTriggerRow = {
  id: string;
  agentId: string;
  connectorId: string;
  enabled: boolean;
  eventTypes: string[];
  conditions: Condition[];
  prompt: string;
  runner: string;
  settings: Record<string, unknown>;
  // Sealed (event-trigger:<agentId>:<app>); never leaves the service.
  secret: Buffer | null;
  secretKeyId: string | null;
  managedAgentId: string | null;
  environmentId: string | null;
  vaultIds: string[];
  budgetCents: number;
  maxRunsPerHour: number;
  maxCentsPerDay: number;
  maxConcurrent: number;
  updatedBy: string | null;
  updatedAt: string;
};

export type EventRunRow = {
  id: string;
  agentId: string;
  triggerId: string;
  connectorId: string;
  eventType: string;
  deliveryId: string;
  runner: string;
  sessionId: string | null;
  state: "starting" | "running" | "ended" | "failed";
  error: string | null;
  costCents: number | null;
  startedAt: string;
  finishedAt: string | null;
};

export type QueuedEvent = {
  id: string;
  subscriptionId: string;
  organizationId: string;
  connectorId: string;
  deliveryId: string;
  eventType: string;
  payload: Buffer | null;
  payloadKeyId: string | null;
};

// What deciding on an event needs to know about one agent and the app.
export type AgentAppState = {
  agentName: string;
  agentStatus: "active" | "disabled";
  grant: { capabilities: string[]; limits: Record<string, unknown> } | null;
  connection: { capabilities: string[]; connected: boolean } | null;
};

const SUBSCRIPTION = `id::text, organization_id::text as "organizationId",
  connector_id as "connectorId", source, secret, secret_key_id as "secretKeyId", filter,
  interval_seconds as "intervalSeconds", cursor, ${iso("polled_at")} as "polledAt",
  last_error as "lastError", registered, created_by as "createdBy",
  ${iso("created_at")} as "createdAt"`;

const TRIGGER = `t.id::text, t.agent_id::text as "agentId", t.connector_id as "connectorId",
  t.enabled, t.event_types as "eventTypes", t.conditions, t.prompt, t.runner, t.settings,
  t.secret, t.secret_key_id as "secretKeyId", t.managed_agent_id as "managedAgentId",
  t.environment_id as "environmentId", t.vault_ids as "vaultIds",
  t.budget_cents as "budgetCents", t.max_runs_per_hour as "maxRunsPerHour",
  t.max_cents_per_day as "maxCentsPerDay", t.max_concurrent as "maxConcurrent",
  t.updated_by as "updatedBy", ${iso("t.updated_at")} as "updatedAt"`;

const RUN = `r.id::text, r.agent_id::text as "agentId", r.trigger_id::text as "triggerId",
  t.connector_id as "connectorId", r.event_type as "eventType", r.delivery_id as "deliveryId",
  r.runner, r.session_id as "sessionId", r.state, r.error, r.cost_cents as "costCents",
  ${iso("r.started_at")} as "startedAt", ${iso("r.finished_at")} as "finishedAt"`;

export function createEventsStore(sql: Sql, schema: string) {
  const t = tables(sql, schema);
  return {
    // ---- subscriptions ----

    // Null when the organization already subscribed to the app.
    async createSubscription(input: {
      id: string;
      organizationId: string;
      connectorId: string;
      source: "webhook" | "poll";
      secret: Buffer | null;
      secretKeyId: string | null;
      filter: string;
      intervalSeconds: number;
      cursor: string | null;
      createdBy: string;
    }): Promise<SubscriptionRow | null> {
      const [row] = await sql.unsafe<SubscriptionRow[]>(
        `insert into ${t.event_subscriptions}
           (id, organization_id, connector_id, source, secret, secret_key_id, filter,
            interval_seconds, cursor, created_by)
         values ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10)
         on conflict (organization_id, connector_id) do nothing returning ${SUBSCRIPTION}`,
        [
          input.id,
          input.organizationId,
          input.connectorId,
          input.source,
          input.secret,
          input.secretKeyId,
          input.filter,
          input.intervalSeconds,
          input.cursor,
          input.createdBy,
        ],
      );
      return row ?? null;
    },

    async subscription(id: string): Promise<SubscriptionRow | null> {
      const [row] = await sql.unsafe<SubscriptionRow[]>(
        `select ${SUBSCRIPTION} from ${t.event_subscriptions} where id = $1`,
        [id],
      );
      return row ?? null;
    },

    async subscriptions(org: string): Promise<SubscriptionRow[]> {
      return sql.unsafe(
        `select ${SUBSCRIPTION} from ${t.event_subscriptions}
          where organization_id = $1 order by connector_id`,
        [org],
      );
    },

    async deleteSubscription(org: string, id: string): Promise<SubscriptionRow | null> {
      const [row] = await sql.unsafe<SubscriptionRow[]>(
        `delete from ${t.event_subscriptions} where organization_id = $1 and id = $2
         returning ${SUBSCRIPTION}`,
        [org, id],
      );
      return row ?? null;
    },

    async setSecret(org: string, id: string, secret: Buffer, keyId: string): Promise<boolean> {
      const rows = await sql`
        update ${t.event_subscriptions} set secret = ${secret}, secret_key_id = ${keyId}
         where organization_id = ${org} and id = ${id} and source = 'webhook' returning id`;
      return rows.length > 0;
    },

    async setRegistered(id: string, registered: string): Promise<void> {
      await sql`update ${t.event_subscriptions} set registered = ${registered} where id = ${id}`;
    },

    // Polls whose interval has passed since they last ran.
    async duePolls(): Promise<SubscriptionRow[]> {
      return sql.unsafe(
        `select ${SUBSCRIPTION} from ${t.event_subscriptions}
          where source = 'poll'
            and (polled_at is null
                 or polled_at < now() - make_interval(secs => interval_seconds))
          order by polled_at nulls first`,
      );
    },

    async polled(id: string, cursor: string | null, error: string | null): Promise<void> {
      await sql`
        update ${t.event_subscriptions}
           set polled_at = now(), cursor = coalesce(${cursor}, cursor), last_error = ${error}
         where id = ${id}`;
    },

    // During a key rotation: the secrets still sealed with another key than `keyId`.
    async staleSecrets(keyId: string) {
      const [subscriptions, triggers] = await Promise.all([
        sql<{ id: string; organizationId: string; secret: Buffer; secretKeyId: string }[]>`
          select s.id::text, s.organization_id::text as "organizationId", s.secret,
                 s.secret_key_id as "secretKeyId"
            from ${t.event_subscriptions} s
           where s.secret is not null and s.secret_key_id <> ${keyId}`,
        sql<
          {
            id: string;
            organizationId: string;
            agentId: string;
            connectorId: string;
            secret: Buffer;
            secretKeyId: string;
          }[]
        >`
          select t.id::text, a.organization_id::text as "organizationId",
                 t.agent_id::text as "agentId", t.connector_id as "connectorId", t.secret,
                 t.secret_key_id as "secretKeyId"
            from ${t.event_triggers} t join ${t.agents} a on a.id = t.agent_id
           where t.secret is not null and t.secret_key_id <> ${keyId}`,
      ]);
      return { subscriptions, triggers };
    },

    // Only over the secret read: one replaced meanwhile stays.
    async reseal(
      table: "event_subscriptions" | "event_triggers",
      id: string,
      was: Buffer,
      secret: Buffer,
      keyId: string,
    ): Promise<boolean> {
      const rows = await sql`
        update ${t[table]} set secret = ${secret}, secret_key_id = ${keyId}
         where id = ${id} and secret = ${was} returning id`;
      return rows.length > 0;
    },

    // ---- the queue ----

    // Null when this delivery was already received (a retry or a replay).
    async insertEvent(input: {
      subscriptionId: string;
      deliveryId: string;
      eventType: string;
      payload: Buffer;
      payloadKeyId: string;
    }): Promise<string | null> {
      const [row] = await sql<{ id: string }[]>`
        insert into ${t.app_events}
          (subscription_id, delivery_id, event_type, payload, payload_key_id)
        values (${input.subscriptionId}, ${input.deliveryId}, ${input.eventType},
                ${input.payload}, ${input.payloadKeyId})
        on conflict (subscription_id, delivery_id) do nothing returning id::text`;
      return row?.id ?? null;
    },

    // The oldest pending event, or one whose worker died. Never the same one to two workers.
    async claimEvent(): Promise<QueuedEvent | null> {
      const [row] = await sql.unsafe<QueuedEvent[]>(
        `with claimed as (
           update ${t.app_events} set state = 'processing', claimed_at = now()
            where id = (
              select id from ${t.app_events}
               where state = 'pending'
                  or (state = 'processing' and claimed_at < now() - interval '10 minutes')
               order by id for update skip locked limit 1)
           returning *)
         select c.id::text, c.subscription_id::text as "subscriptionId",
                s.organization_id::text as "organizationId", s.connector_id as "connectorId",
                c.delivery_id as "deliveryId", c.event_type as "eventType", c.payload,
                c.payload_key_id as "payloadKeyId"
           from claimed c join ${t.event_subscriptions} s on s.id = c.subscription_id`,
      );
      return row ?? null;
    },

    // The payload is kept only until the event is handled.
    async finishEvent(id: string, state: "done" | "failed", reason: string | null) {
      await sql`
        update ${t.app_events}
           set state = ${state}, reason = ${reason}, payload = null, payload_key_id = null
         where id = ${id}`;
    },

    async pruneEvents(): Promise<void> {
      await sql`delete from ${t.app_events} where received_at < now() - interval '30 days'`;
    },

    // ---- triggers ----

    async agentState(agentId: string, connectorId: string): Promise<AgentAppState | null> {
      const [row] = await sql<AgentAppState[]>`
        select a.name as "agentName", a.status as "agentStatus",
               case when g.agent_id is null then null else
                 jsonb_build_object('capabilities', to_jsonb(g.capabilities), 'limits', g.limits)
               end as grant,
               case when c.id is null then null else
                 jsonb_build_object('capabilities', to_jsonb(c.capabilities),
                   'connected', c.status = 'connected'
                     and (c.provider_id is null or p.status = 'connected'))
               end as connection
          from ${t.agents} a
          left join ${t.grants} g on g.agent_id = a.id and g.connector_id = ${connectorId}
          left join ${t.connections} c
            on c.organization_id = a.organization_id and c.connector_id = ${connectorId}
          left join ${t.providers} p on p.id = c.provider_id
         where a.id = ${agentId}`;
      return row ?? null;
    },

    // Every enabled trigger on the app, of the organization's agents.
    async triggersFor(org: string, connectorId: string): Promise<EventTriggerRow[]> {
      return sql.unsafe(
        `select ${TRIGGER} from ${t.event_triggers} t join ${t.agents} a on a.id = t.agent_id
          where a.organization_id = $1 and t.connector_id = $2 and t.enabled
          order by t.agent_id`,
        [org, connectorId],
      );
    },

    async trigger(org: string, agentId: string, connectorId: string) {
      const [row] = await sql.unsafe<EventTriggerRow[]>(
        `select ${TRIGGER} from ${t.event_triggers} t join ${t.agents} a on a.id = t.agent_id
          where a.organization_id = $1 and t.agent_id = $2 and t.connector_id = $3`,
        [org, agentId, connectorId],
      );
      return row ?? null;
    },

    async triggerById(id: string): Promise<EventTriggerRow | null> {
      const [row] = await sql.unsafe<EventTriggerRow[]>(
        `select ${TRIGGER} from ${t.event_triggers} t where t.id = $1`,
        [id],
      );
      return row ?? null;
    },

    async agentTriggers(org: string, agentId: string): Promise<EventTriggerRow[]> {
      return sql.unsafe(
        `select ${TRIGGER} from ${t.event_triggers} t join ${t.agents} a on a.id = t.agent_id
          where a.organization_id = $1 and t.agent_id = $2 order by t.connector_id`,
        [org, agentId],
      );
    },

    // Only for an agent of this organization; false otherwise.
    async setTrigger(
      org: string,
      row: Omit<EventTriggerRow, "id" | "updatedAt">,
    ): Promise<boolean> {
      const rows = await sql`
        insert into ${t.event_triggers}
          (agent_id, connector_id, enabled, event_types, conditions, prompt, runner, settings,
           secret, secret_key_id, managed_agent_id, environment_id, vault_ids, budget_cents,
           max_runs_per_hour, max_cents_per_day, max_concurrent, updated_by)
        select a.id, ${row.connectorId}, ${row.enabled}, ${sql.array(row.eventTypes)},
               ${sql.json(row.conditions)}, ${row.prompt}, ${row.runner},
               ${sql.json(row.settings as Parameters<typeof sql.json>[0])}, ${row.secret},
               ${row.secretKeyId}, ${row.managedAgentId}, ${row.environmentId},
               ${sql.array(row.vaultIds)}, ${row.budgetCents}, ${row.maxRunsPerHour},
               ${row.maxCentsPerDay}, ${row.maxConcurrent}, ${row.updatedBy}
          from ${t.agents} a
         where a.id = ${row.agentId} and a.organization_id = ${org}
        on conflict (agent_id, connector_id) do update set
          enabled = excluded.enabled, event_types = excluded.event_types,
          conditions = excluded.conditions, prompt = excluded.prompt, runner = excluded.runner,
          settings = excluded.settings, secret = excluded.secret,
          secret_key_id = excluded.secret_key_id, managed_agent_id = excluded.managed_agent_id,
          environment_id = excluded.environment_id, vault_ids = excluded.vault_ids,
          budget_cents = excluded.budget_cents, max_runs_per_hour = excluded.max_runs_per_hour,
          max_cents_per_day = excluded.max_cents_per_day,
          max_concurrent = excluded.max_concurrent, updated_by = excluded.updated_by,
          updated_at = now()
        returning agent_id`;
      return rows.length > 0;
    },

    async removeTrigger(org: string, agentId: string, connectorId: string): Promise<boolean> {
      const rows = await sql`
        delete from ${t.event_triggers} t using ${t.agents} a
         where a.id = t.agent_id and a.organization_id = ${org} and t.agent_id = ${agentId}
           and t.connector_id = ${connectorId}
        returning t.id`;
      return rows.length > 0;
    },

    // ---- runs ----

    // Runs live now, started in the last hour (failed starts too), and what finished runs
    // cost in the last 24 hours, for one trigger.
    async runCounts(triggerId: string): Promise<RunCounts> {
      const [row] = await sql<RunCounts[]>`
        select count(*) filter (where state in ('starting', 'running'))::int as active,
               count(*) filter (where started_at > now() - interval '1 hour')::int as "lastHour",
               coalesce(sum(cost_cents) filter (where started_at > now() - interval '24 hours'),
                        0)::int as "spentToday"
          from ${t.event_runs}
         where trigger_id = ${triggerId} and started_at > now() - interval '24 hours'
                or (trigger_id = ${triggerId} and state in ('starting', 'running'))`;
      return row;
    },

    async startRun(input: {
      agentId: string;
      triggerId: string;
      eventId: string;
      eventType: string;
      deliveryId: string;
      runner: string;
    }): Promise<string> {
      const [row] = await sql<{ id: string }[]>`
        insert into ${t.event_runs}
          (agent_id, trigger_id, event_id, event_type, delivery_id, runner, state)
        values (${input.agentId}, ${input.triggerId}, ${input.eventId}, ${input.eventType},
                ${input.deliveryId}, ${input.runner}, 'starting')
        returning id::text`;
      return row.id;
    },

    async updateRun(
      id: string,
      patch: {
        state?: EventRunRow["state"];
        sessionId?: string;
        error?: string;
        costCents?: number;
      },
    ): Promise<void> {
      const finished = patch.state === "ended" || patch.state === "failed";
      await sql`
        update ${t.event_runs}
           set state = coalesce(${patch.state ?? null}, state),
               session_id = coalesce(${patch.sessionId ?? null}, session_id),
               error = coalesce(${patch.error ?? null}, error),
               cost_cents = coalesce(${patch.costCents ?? null}, cost_cents),
               finished_at = case when ${finished} then now() else finished_at end
         where id = ${id}`;
    },

    // Runs left starting (the process stopped mid-start) hold a concurrency slot forever.
    async failStaleStarts(): Promise<number> {
      const rows = await sql`
        update ${t.event_runs}
           set state = 'failed', error = 'Did not start', finished_at = now()
         where state = 'starting' and started_at < now() - interval '10 minutes'
        returning id`;
      return rows.length;
    },

    async runningRuns(): Promise<(EventRunRow & { organizationId: string })[]> {
      return sql.unsafe(
        `select ${RUN}, a.organization_id::text as "organizationId"
           from ${t.event_runs} r join ${t.event_triggers} t on t.id = r.trigger_id
           join ${t.agents} a on a.id = r.agent_id
          where r.state = 'running' and r.session_id is not null order by r.started_at`,
      );
    },

    async recentRuns(org: string, agentId: string, limit: number): Promise<EventRunRow[]> {
      return sql.unsafe(
        `select ${RUN} from ${t.event_runs} r join ${t.event_triggers} t on t.id = r.trigger_id
          where r.agent_id = $2 and exists (select 1 from ${t.agents} a
                                             where a.id = $2 and a.organization_id = $1)
          order by r.started_at desc limit $3`,
        [org, agentId, limit],
      );
    },
  };
}

export type EventsStore = ReturnType<typeof createEventsStore>;
