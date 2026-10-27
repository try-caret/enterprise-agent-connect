import { tables, type Sql } from "../store.js";
import type { ApprovalFact, ApprovalRow } from "./approval.js";

// Postgres for Teams bots, where they are installed, the inbound queue, each agent's
// real-time settings and the sessions those started (migrations/), plus the organization's
// Claude API key.

const iso = (column: string) =>
  `to_char(${column} at time zone 'UTC', 'YYYY-MM-DD"T"HH24:MI:SS.MS"Z"')`;

export type BotRow = {
  id: string;
  organizationId: string;
  agentId: string;
  botAppId: string | null;
  teamsAppId: string;
  tenantId: string;
  displayName: string;
  manifestVersion: number;
  status: "awaiting_setup" | "connected" | "needs_reconnect";
  // { certificate, nextCertificate? }
  config: Record<string, string>;
  credential: Buffer;
  credentialKeyId: string;
  createdBy: string | null;
  createdAt: string;
  lastRefusedAt: string | null;
  lastRefusedReason: string | null;
};

export type ConversationRow = {
  botId: string;
  conversationId: string;
  kind: "personal" | "groupChat" | "team";
  serviceUrl: string;
  teamId: string | null;
  teamGroupId: string | null;
  name: string | null;
  userAadId: string | null;
  removedAt: string | null;
  limitNoticeAt: string | null;
};

export type EventRow = {
  id: string;
  botId: string;
  activityId: string;
  type: string;
  conversationId: string;
  threadId: string | null;
  payload: Buffer | null;
  payloadKeyId: string | null;
};

export type TriggerRow = {
  agentId: string;
  enabled: boolean;
  // What answers (see runner.ts) and its settings. The Claude runner's settings are the
  // managedAgentId, environmentId and vaultIds columns, which predate the others.
  runner: string;
  settings: Record<string, unknown>;
  // A signing secret, sealed (teams-trigger:<agentId>); never leaves the service.
  secret: Buffer | null;
  secretKeyId: string | null;
  managedAgentId: string | null;
  environmentId: string | null;
  vaultIds: string[];
  budgetCents: number;
  maxRunsPerHour: number;
  maxCentsPerDay: number;
  maxConcurrent: number;
  replyMode: "mentions" | "all";
  allowedTeams: string[];
  allowedPeople: string[];
  anyoneInOrganization: boolean;
  updatedBy: string | null;
  updatedAt: string;
};

export type RunRow = {
  id: string;
  agentId: string;
  botId: string;
  conversationId: string;
  threadId: string;
  triggerActivityId: string;
  // The runner that started the session (runner.ts); only it may continue or stop it.
  runner: string;
  sessionId: string | null;
  state: "starting" | "running" | "idle" | "ended" | "failed" | "no_reply";
  error: string | null;
  note: string | null;
  requestedBy: string | null;
  replyCount: number;
  repliedAt: string | null;
  costCents: number | null;
  turns: number;
  startedAt: string;
  finishedAt: string | null;
  replyTokenAt: string;
};

// What deciding on a message needs to know about the bot's agent.
export type AgentTeamsState = {
  agentName: string;
  agentStatus: "active" | "disabled";
  // The agent's Teams grant, and the Teams app's ceiling and status. Null when missing.
  grant: { capabilities: string[]; limits: Record<string, string[] | string> } | null;
  connection: { capabilities: string[]; status: string } | null;
};

const BOT = `id::text, organization_id::text as "organizationId", agent_id::text as "agentId",
  bot_app_id::text as "botAppId", teams_app_id::text as "teamsAppId", tenant_id::text as "tenantId",
  display_name as "displayName", manifest_version as "manifestVersion", status, config,
  credential, credential_key_id as "credentialKeyId", created_by as "createdBy",
  ${iso("created_at")} as "createdAt", ${iso("last_refused_at")} as "lastRefusedAt",
  last_refused_reason as "lastRefusedReason"`;

const CONVERSATION = `bot_id::text as "botId", conversation_id as "conversationId", kind,
  service_url as "serviceUrl", team_id as "teamId", team_group_id::text as "teamGroupId", name,
  user_aad_id::text as "userAadId", ${iso("removed_at")} as "removedAt",
  ${iso("limit_notice_at")} as "limitNoticeAt"`;

const TRIGGER = `agent_id::text as "agentId", enabled, runner, settings, secret,
  secret_key_id as "secretKeyId", managed_agent_id as "managedAgentId",
  environment_id as "environmentId", vault_ids as "vaultIds", budget_cents as "budgetCents",
  max_runs_per_hour as "maxRunsPerHour", max_cents_per_day as "maxCentsPerDay",
  anyone_in_organization as "anyoneInOrganization", max_concurrent as "maxConcurrent",
  reply_mode as "replyMode", allowed_teams::text[] as "allowedTeams",
  allowed_people::text[] as "allowedPeople", updated_by as "updatedBy",
  ${iso("updated_at")} as "updatedAt"`;

const RUN = `id::text, agent_id::text as "agentId", bot_id::text as "botId",
  conversation_id as "conversationId", thread_id as "threadId",
  trigger_activity_id as "triggerActivityId", runner, session_id as "sessionId", state, error, note,
  requested_by_aad_id::text as "requestedBy", reply_count as "replyCount",
  ${iso("replied_at")} as "repliedAt", cost_cents as "costCents", turns, ${iso("started_at")} as "startedAt",
  ${iso("finished_at")} as "finishedAt", ${iso("reply_token_at")} as "replyTokenAt"`;

const APPROVAL = `id::text, agent_id::text as "agentId",
  bot_id::text as "botId", conversation_id as "conversationId", service_url as "serviceUrl",
  thread_id as "threadId", activity_id as "activityId", approver_aad_id::text as "approverAadId",
  approver_name as "approverName", title, body, facts, status,
  decided_by_aad_id::text as "decidedByAadId", decided_by_name as "decidedByName",
  ${iso("decided_at")} as "decidedAt", comment, ${iso("created_at")} as "createdAt",
  ${iso("expires_at")} as "expiresAt", (status = 'pending' and expires_at <= now()) as expired`;

// Replies the agent may send to one message. A run still starting may reply: a runner can be
// answered before its start call returns (a webhook receiver replying in its handler).
export const MAX_REPLIES = 3;
const REPLYABLE = `state in ('starting', 'running', 'idle') and reply_token_at > now() - interval '12 hours'
  and reply_count < ${MAX_REPLIES}`;

export function createTeamsStore(sql: Sql, schema: string) {
  const t = tables(sql, schema);
  return {
    // ---- bots ----

    async createBot(input: {
      id: string;
      organizationId: string;
      agentId: string;
      tenantId: string;
      displayName: string;
      config: Record<string, string>;
      credential: Buffer;
      credentialKeyId: string;
      createdBy: string;
    }): Promise<boolean> {
      const rows = await sql`
        insert into ${t.teams_bots}
          (id, organization_id, agent_id, tenant_id, display_name, status, config, credential,
           credential_key_id, created_by)
        select ${input.id}, a.organization_id, a.id, ${input.tenantId}, ${input.displayName},
               'awaiting_setup', ${sql.json(input.config)}, ${input.credential},
               ${input.credentialKeyId}, ${input.createdBy}
          from ${t.agents} a
         where a.id = ${input.agentId} and a.organization_id = ${input.organizationId}
        on conflict (agent_id) do nothing
        returning id`;
      return rows.length > 0;
    },

    async listBots(org: string): Promise<BotRow[]> {
      return sql.unsafe(
        `select ${BOT} from ${t.teams_bots} where organization_id = $1
          order by created_at`,
        [org],
      );
    },

    async botForAgent(org: string, agentId: string): Promise<BotRow | null> {
      const [row] = await sql.unsafe<BotRow[]>(
        `select ${BOT} from ${t.teams_bots}
          where organization_id = $1 and agent_id = $2`,
        [org, agentId],
      );
      return row ?? null;
    },

    async botByAppId(botAppId: string): Promise<BotRow | null> {
      const [row] = await sql.unsafe<BotRow[]>(
        `select ${BOT} from ${t.teams_bots} where bot_app_id = $1`,
        [botAppId],
      );
      return row ?? null;
    },

    async botById(id: string): Promise<BotRow | null> {
      const [row] = await sql.unsafe<BotRow[]>(`select ${BOT} from ${t.teams_bots} where id = $1`, [
        id,
      ]);
      return row ?? null;
    },

    // Proven by a sign-in: the bot's app id is now known. False when another bot already has
    // it, or the bot changed meanwhile.
    async connectBot(id: string, botAppId: string, replaces: Buffer): Promise<boolean> {
      try {
        const rows = await sql`
          update ${t.teams_bots}
             set bot_app_id = ${botAppId}, status = 'connected', checked_at = now()
           where id = ${id} and credential = ${replaces} returning id`;
        return rows.length > 0;
      } catch (error) {
        if ((error as { code?: string }).code === "23505") return false;
        throw error;
      }
    },

    async updateBotCredential(
      id: string,
      config: Record<string, string>,
      credential: Buffer,
      credentialKeyId: string,
      replaces: Buffer,
    ): Promise<boolean> {
      const rows = await sql`
        update ${t.teams_bots}
           set config = ${sql.json(config)}, credential = ${credential},
               credential_key_id = ${credentialKeyId}
         where id = ${id} and credential = ${replaces} returning id`;
      return rows.length > 0;
    },

    // Bots and Claude keys still sealed with one of `keyIds` (the previous keys of a rotation).
    async staleSealed(keyIds: string[]) {
      const bots = await sql<
        {
          id: string;
          organizationId: string;
          config: Record<string, string>;
          credential: Buffer;
          credentialKeyId: string;
        }[]
      >`
        select id::text, organization_id::text as "organizationId", config, credential,
               credential_key_id as "credentialKeyId"
          from ${t.teams_bots} where credential_key_id = any(${keyIds})`;
      const claudeKeys = await sql<
        { id: string; organizationId: string; credential: Buffer; credentialKeyId: string }[]
      >`
        select id::text, organization_id::text as "organizationId", credential,
               credential_key_id as "credentialKeyId"
          from ${t.providers}
         where provider = 'anthropic' and credential_key_id = any(${keyIds})`;
      const triggers = await sql<
        { agentId: string; organizationId: string; secret: Buffer; secretKeyId: string }[]
      >`
        select t.agent_id::text as "agentId", a.organization_id::text as "organizationId",
               t.secret, t.secret_key_id as "secretKeyId"
          from ${t.teams_triggers} t join ${t.agents} a on a.id = t.agent_id
         where t.secret_key_id = any(${keyIds})`;
      return { bots, claudeKeys, triggers };
    },

    // Only over the secret read: one saved meanwhile wins.
    async resealTriggerSecret(
      agentId: string,
      secret: Buffer,
      secretKeyId: string,
      replaces: Buffer,
    ): Promise<boolean> {
      const rows = await sql`
        update ${t.teams_triggers} set secret = ${secret}, secret_key_id = ${secretKeyId}
         where agent_id = ${agentId} and secret = ${replaces} returning agent_id`;
      return rows.length > 0;
    },

    // Only over the credential read: a key re-entered meanwhile wins.
    async resealClaudeKey(
      id: string,
      credential: Buffer,
      credentialKeyId: string,
      replaces: Buffer,
    ): Promise<boolean> {
      const rows = await sql`
        update ${t.providers}
           set credential = ${credential}, credential_key_id = ${credentialKeyId}
         where id = ${id} and provider = 'anthropic' and credential = ${replaces} returning id`;
      return rows.length > 0;
    },

    async markBot(id: string, ok: boolean): Promise<void> {
      await sql`
        update ${t.teams_bots}
           set status = ${ok ? "connected" : "needs_reconnect"}, checked_at = now()
         where id = ${id} and status <> 'awaiting_setup'`;
    },

    // Why the last request for this bot was refused, for the admin.
    async refused(botAppId: string, reason: string): Promise<void> {
      await sql`
        update ${t.teams_bots}
           set last_refused_at = now(), last_refused_reason = ${reason}
         where bot_app_id = ${botAppId}`;
    },

    async deleteBot(org: string, agentId: string): Promise<boolean> {
      const rows = await sql`
        delete from ${t.teams_bots}
         where organization_id = ${org} and agent_id = ${agentId} returning id`;
      return rows.length > 0;
    },

    async agentState(agentId: string): Promise<AgentTeamsState | null> {
      const [row] = await sql<AgentTeamsState[]>`
        select a.name as "agentName", a.status as "agentStatus",
               case when g.agent_id is null then null else
                 jsonb_build_object('capabilities', to_jsonb(g.capabilities), 'limits', g.limits)
               end as grant,
               case when c.id is null then null else
                 jsonb_build_object('capabilities', to_jsonb(c.capabilities), 'status', c.status)
               end as connection
          from ${t.agents} a
          left join ${t.grants} g
            on g.agent_id = a.id and g.connector_id = 'teams'
          left join ${t.connections} c
            on c.organization_id = a.organization_id and c.connector_id = 'teams'
         where a.id = ${agentId}`;
      return row ?? null;
    },

    // ---- conversations ----

    async upsertConversation(
      row: Omit<ConversationRow, "removedAt" | "limitNoticeAt">,
    ): Promise<void> {
      await sql`
        insert into ${t.teams_conversations} as c
          (bot_id, conversation_id, kind, service_url, team_id, team_group_id, name, user_aad_id)
        values (${row.botId}, ${row.conversationId}, ${row.kind}, ${row.serviceUrl}, ${row.teamId},
                ${row.teamGroupId}, ${row.name}, ${row.userAadId})
        on conflict (bot_id, conversation_id) do update set
          kind = excluded.kind, service_url = excluded.service_url,
          team_id = coalesce(excluded.team_id, c.team_id),
          team_group_id = coalesce(excluded.team_group_id, c.team_group_id),
          name = coalesce(excluded.name, c.name),
          user_aad_id = coalesce(excluded.user_aad_id, c.user_aad_id),
          removed_at = null`;
    },

    async removeConversation(botId: string, conversationId: string): Promise<void> {
      await sql`
        update ${t.teams_conversations} set removed_at = now()
         where bot_id = ${botId} and conversation_id = ${conversationId} and removed_at is null`;
    },

    async conversations(botId: string): Promise<ConversationRow[]> {
      return sql.unsafe(
        `select ${CONVERSATION} from ${t.teams_conversations}
          where bot_id = $1 and removed_at is null order by installed_at`,
        [botId],
      );
    },

    async conversationsOf(botIds: string[]): Promise<ConversationRow[]> {
      if (!botIds.length) return [];
      return sql.unsafe(
        `select ${CONVERSATION} from ${t.teams_conversations}
          where bot_id = any($1::uuid[]) and removed_at is null order by installed_at`,
        [botIds],
      );
    },

    // True when no notice was sent there in the last hour (and records this one).
    async claimLimitNotice(botId: string, conversationId: string): Promise<boolean> {
      const rows = await sql`
        update ${t.teams_conversations} set limit_notice_at = now()
         where bot_id = ${botId} and conversation_id = ${conversationId}
           and (limit_notice_at is null or limit_notice_at < now() - interval '1 hour')
        returning bot_id`;
      return rows.length > 0;
    },

    // ---- inbound queue ----

    // False when Teams already delivered this activity (a retry): nothing more to do.
    async insertEvent(input: {
      botId: string;
      activityId: string;
      type: string;
      conversationId: string;
      threadId: string | null;
      fromAadId: string | null;
      payload: Buffer;
      payloadKeyId: string;
    }): Promise<boolean> {
      const rows = await sql`
        insert into ${t.teams_events}
          (bot_id, activity_id, type, conversation_id, thread_id, from_aad_id, state, payload,
           payload_key_id)
        values (${input.botId}, ${input.activityId}, ${input.type}, ${input.conversationId},
                ${input.threadId}, ${input.fromAadId}, 'pending', ${input.payload},
                ${input.payloadKeyId})
        on conflict (bot_id, activity_id) do nothing returning id`;
      return rows.length > 0;
    },

    // The oldest pending event, or one whose worker died (claimed over 10 minutes ago). Never
    // the same event to two workers.
    async claimEvent(): Promise<EventRow | null> {
      const [row] = await sql<EventRow[]>`
        update ${t.teams_events} set state = 'processing', claimed_at = now()
         where id = (
           select id from ${t.teams_events}
            where state = 'pending'
               or (state = 'processing' and claimed_at < now() - interval '10 minutes')
            order by id for update skip locked limit 1)
        returning id::text, bot_id::text as "botId", activity_id as "activityId", type,
                  conversation_id as "conversationId", thread_id as "threadId", payload,
                  payload_key_id as "payloadKeyId"`;
      return row ?? null;
    },

    // The payload (message text) is kept only while the event may still be sent on.
    async finishEvent(
      id: string,
      state: "ignored" | "run" | "done" | "failed" | "held",
      reason: string | null,
    ): Promise<void> {
      await sql`
        update ${t.teams_events}
           set state = ${state}, reason = ${reason},
               payload = case when ${state} = 'held' then payload else null end,
               payload_key_id = case when ${state} = 'held' then payload_key_id else null end
         where id = ${id}`;
    },

    // Messages that arrived while their thread's session was busy go back into the queue once
    // no session is active there. Returns how many.
    async requeueHeld(): Promise<number> {
      const rows = await sql`
        update ${t.teams_events} e set state = 'pending'
         where e.state = 'held'
           and not exists (select 1 from ${t.teams_runs} r
                            where r.bot_id = e.bot_id and r.thread_id = e.thread_id
                              and r.state in ('starting', 'running'))
        returning e.id`;
      return rows.length;
    },

    async pruneEvents(): Promise<void> {
      await sql`
        delete from ${t.teams_events} where received_at < now() - interval '30 days'`;
    },

    // ---- triggers ----

    async trigger(agentId: string): Promise<TriggerRow | null> {
      const [row] = await sql.unsafe<TriggerRow[]>(
        `select ${TRIGGER} from ${t.teams_triggers} where agent_id = $1`,
        [agentId],
      );
      return row ?? null;
    },

    async setTrigger(org: string, row: Omit<TriggerRow, "updatedAt">): Promise<boolean> {
      const rows = await sql`
        insert into ${t.teams_triggers}
          (agent_id, enabled, runner, settings, secret, secret_key_id, managed_agent_id,
           environment_id, vault_ids, budget_cents, max_runs_per_hour, max_cents_per_day,
           max_concurrent, reply_mode, allowed_teams, allowed_people, anyone_in_organization,
           updated_by)
        select a.id, ${row.enabled}, ${row.runner}, ${sql.json(row.settings as Parameters<typeof sql.json>[0])},
               ${row.secret}, ${row.secretKeyId}, ${row.managedAgentId}, ${row.environmentId},
               ${sql.array(row.vaultIds)}, ${row.budgetCents}, ${row.maxRunsPerHour},
               ${row.maxCentsPerDay}, ${row.maxConcurrent}, ${row.replyMode},
               ${sql.array(row.allowedTeams)}::uuid[], ${sql.array(row.allowedPeople)}::uuid[],
               ${row.anyoneInOrganization}, ${row.updatedBy}
          from ${t.agents} a
         where a.id = ${row.agentId} and a.organization_id = ${org}
        on conflict (agent_id) do update set
          enabled = excluded.enabled, runner = excluded.runner, settings = excluded.settings,
          secret = excluded.secret, secret_key_id = excluded.secret_key_id,
          managed_agent_id = excluded.managed_agent_id,
          environment_id = excluded.environment_id, vault_ids = excluded.vault_ids,
          budget_cents = excluded.budget_cents, max_runs_per_hour = excluded.max_runs_per_hour,
          max_cents_per_day = excluded.max_cents_per_day,
          anyone_in_organization = excluded.anyone_in_organization,
          max_concurrent = excluded.max_concurrent, reply_mode = excluded.reply_mode,
          allowed_teams = excluded.allowed_teams, allowed_people = excluded.allowed_people,
          updated_by = excluded.updated_by, updated_at = now()
        returning agent_id`;
      return rows.length > 0;
    },

    // ---- runs ----

    // Sessions running now; messages answered or tried in the last hour (each turn counts, and
    // so does a start that failed); what this agent's turns cost in the last 24 hours.
    async runCounts(
      agentId: string,
      botId: string,
    ): Promise<{ active: number; lastHour: number; spentToday: number }> {
      const [row] = await sql<{ active: number; lastHour: number; spentToday: number }[]>`
        select (select count(*) from ${t.teams_runs}
                 where agent_id = ${agentId} and state in ('starting', 'running'))::int as active,
               (select count(*) from ${t.teams_events}
                 where bot_id = ${botId} and state in ('run', 'failed')
                   and received_at > now() - interval '1 hour')::int as "lastHour",
               (select coalesce(sum(cents), 0) from ${t.teams_spend}
                 where agent_id = ${agentId}
                   and spent_at > now() - interval '24 hours')::int as "spentToday"`;
      return row;
    },

    // The thread's latest session, whatever its state.
    async latestRun(botId: string, threadId: string): Promise<RunRow | null> {
      const [row] = await sql.unsafe<RunRow[]>(
        `select ${RUN} from ${t.teams_runs}
          where bot_id = $1 and thread_id = $2 order by started_at desc limit 1`,
        [botId, threadId],
      );
      return row ?? null;
    },

    // A new session for the thread. Null when one is already active there (unique index).
    async startRun(input: {
      agentId: string;
      botId: string;
      conversationId: string;
      threadId: string;
      triggerActivityId: string;
      replyTokenHash: string;
      requestedBy: string;
      runner: string;
    }): Promise<RunRow | null> {
      try {
        const [row] = await sql.unsafe<RunRow[]>(
          `insert into ${t.teams_runs}
             (agent_id, bot_id, conversation_id, thread_id, trigger_activity_id, reply_token_hash,
              requested_by_aad_id, runner, state)
           values ($1, $2, $3, $4, $5, $6, $7, $8, 'starting') returning ${RUN}`,
          [
            input.agentId,
            input.botId,
            input.conversationId,
            input.threadId,
            input.triggerActivityId,
            input.replyTokenHash,
            input.requestedBy,
            input.runner,
          ],
        );
        return row;
      } catch (error) {
        if ((error as { code?: string }).code === "23505") return null;
        throw error;
      }
    },

    // The next turn in an idle session: a new reply token, for the new message. False when the
    // session is no longer idle (another message got there first). It stays 'starting' until
    // the message is sent, so the watcher never reads the previous turn's idle as this one's.
    async continueRun(
      id: string,
      triggerActivityId: string,
      replyTokenHash: string,
    ): Promise<boolean> {
      try {
        const rows = await sql`
          update ${t.teams_runs}
             set state = 'starting', trigger_activity_id = ${triggerActivityId},
                 reply_token_hash = ${replyTokenHash}, reply_token_at = now(),
                 reply_count = 0, turns = turns + 1, finished_at = null
           where id = ${id} and state = 'idle' returning id`;
        return rows.length > 0;
      } catch (error) {
        if ((error as { code?: string }).code === "23505") return false;
        throw error;
      }
    },

    async updateRun(
      id: string,
      patch: {
        state?: RunRow["state"];
        sessionId?: string;
        error?: string;
        note?: string;
        costCents?: number;
      },
    ): Promise<void> {
      const finished = patch.state && patch.state !== "running" && patch.state !== "starting";
      // cost_cents is the session's total so far; its increase is this turn's spend, logged.
      // prev reads the row as the statement began (one replica, so one writer per run).
      await sql`
        with prev as (
          select cost_cents as old from ${t.teams_runs} where id = ${id}
        ), updated as (
          update ${t.teams_runs}
             set state = coalesce(${patch.state ?? null}, state),
                 session_id = coalesce(${patch.sessionId ?? null}, session_id),
                 error = coalesce(${patch.error ?? null}, error),
                 note = coalesce(${patch.note ?? null}, note),
                 cost_cents = coalesce(${patch.costCents ?? null}, cost_cents),
                 finished_at = case when ${Boolean(finished)} then now() else finished_at end
           where id = ${id}
          returning agent_id, cost_cents
        )
        insert into ${t.teams_spend} (agent_id, run_id, cents)
        select u.agent_id, ${id}, u.cost_cents - coalesce(p.old, 0)
          from updated u, prev p
         where u.cost_cents > coalesce(p.old, 0)`;
    },

    // Runs stuck starting (the process stopped, or a write failed, before the start finished)
    // would hold their thread forever under the one-active index. Returns how many failed.
    async failStaleStarts(): Promise<number> {
      const rows = await sql`
        update ${t.teams_runs}
           set state = 'failed', error = 'Did not start', finished_at = now()
         where state = 'starting' and reply_token_at < now() - interval '10 minutes'
        returning id`;
      return rows.length;
    },

    async runningRuns(): Promise<RunRow[]> {
      return sql.unsafe(
        `select ${RUN} from ${t.teams_runs}
          where state = 'running' and session_id is not null order by started_at`,
      );
    },

    // The run a reply token belongs to, only for the agent calling, while its session is live,
    // and for a few replies per message. Reading it uses no reply: claimReply does.
    async runForReply(agentId: string, tokenHash: string): Promise<RunRow | null> {
      const [row] = await sql.unsafe<RunRow[]>(
        `select ${RUN} from ${t.teams_runs}
          where agent_id = $1 and reply_token_hash = $2 and ${REPLYABLE}`,
        [agentId, tokenHash],
      );
      return row ?? null;
    },

    // Uses one of the run's replies, if it still has one and is still live.
    async claimReply(agentId: string, runId: string): Promise<boolean> {
      const rows = await sql.unsafe(
        `update ${t.teams_runs}
            set reply_count = reply_count + 1, replied_at = now()
          where agent_id = $1 and id = $2 and ${REPLYABLE} returning id`,
        [agentId, runId],
      );
      return rows.length > 0;
    },

    // Gives back a reply whose post failed.
    async releaseReply(runId: string): Promise<void> {
      await sql`
        update ${t.teams_runs} set reply_count = reply_count - 1
         where id = ${runId} and reply_count > 0`;
    },

    async recentRuns(org: string, agentId: string, limit: number): Promise<RunRow[]> {
      return sql.unsafe(
        `select ${RUN} from ${t.teams_runs}
          where agent_id = $2 and exists (select 1 from ${t.agents} a
                                           where a.id = $2 and a.organization_id = $1)
          order by started_at desc limit $3`,
        [org, agentId, limit],
      );
    },

    // ---- approval cards (approval.ts) ----

    async createApproval(input: {
      agentId: string;
      botId: string;
      tokenHash: string;
      conversationId: string;
      serviceUrl: string;
      approverAadId: string;
      approverName: string;
      title: string;
      body: string;
      facts: ApprovalFact[];
      days: number;
    }): Promise<ApprovalRow> {
      const [row] = await sql<ApprovalRow[]>`
        insert into ${t.approvals}
          (agent_id, bot_id, token_hash, conversation_id, service_url,
           approver_aad_id, approver_name, title, body, facts, expires_at)
        select b.agent_id, b.id, ${input.tokenHash}, ${input.conversationId},
               ${input.serviceUrl}, ${input.approverAadId}, ${input.approverName}, ${input.title},
               ${input.body}, ${sql.json(input.facts)}, now() + make_interval(days => ${input.days})
          from ${t.teams_bots} b where b.id = ${input.botId} and b.agent_id = ${input.agentId}
        returning ${sql.unsafe(APPROVAL)}`;
      if (!row) throw new Error("The agent's bot is gone");
      return row;
    },

    // Where the card was posted, so a decision can edit it. Answers the row as it now is: a
    // click can land before this runs, and then only this caller can edit the card.
    async approvalPosted(
      id: string,
      threadId: string,
      activityId: string,
    ): Promise<ApprovalRow | null> {
      const [row] = await sql.unsafe<ApprovalRow[]>(
        `update ${t.approvals} set thread_id = $2, activity_id = $3 where id = $1
          returning ${APPROVAL}`,
        [id, threadId, activityId],
      );
      return row ?? null;
    },

    // A card that could not be posted asks nobody.
    async discardApproval(id: string): Promise<void> {
      await sql`delete from ${t.approvals} where id = ${id} and status = 'pending'`;
    },

    // Only the agent that asked reads an approval.
    async approvalForAgent(agentId: string, tokenHash: string): Promise<ApprovalRow | null> {
      const [row] = await sql.unsafe<ApprovalRow[]>(
        `select ${APPROVAL} from ${t.approvals} where agent_id = $1 and token_hash = $2`,
        [agentId, tokenHash],
      );
      return row ?? null;
    },

    // A card's approval, only through the bot that posted it and from the place (the team or
    // the chat) it was posted in, so a forwarded card or a crafted 1:1 invoke finds nothing.
    async approvalForBot(
      botId: string,
      id: string,
      conversationId: string,
    ): Promise<ApprovalRow | null> {
      const [row] = await sql.unsafe<ApprovalRow[]>(
        `select ${APPROVAL} from ${t.approvals}
          where bot_id = $1 and id = $2 and conversation_id = $3`,
        [botId, id, conversationId],
      );
      return row ?? null;
    },

    // Records the answer once, and only the approver's: null when it was already answered, has
    // expired, or the clicker or the place is not the card's. Every condition is checked here
    // again, in the one statement that writes, whatever the caller checked before.
    async decideApproval(
      botId: string,
      id: string,
      conversationId: string,
      decision: { status: "approved" | "rejected"; byAadId: string; byName: string },
      comment: string | null,
    ): Promise<ApprovalRow | null> {
      const [row] = await sql.unsafe<ApprovalRow[]>(
        `update ${t.approvals}
            set status = $3, decided_by_aad_id = $4, decided_by_name = $5, decided_at = now(),
                comment = $6
          where bot_id = $1 and id = $2 and conversation_id = $7 and status = 'pending'
            and expires_at > now() and approver_aad_id = $4
          returning ${APPROVAL}`,
        [botId, id, decision.status, decision.byAadId, decision.byName, comment, conversationId],
      );
      return row ?? null;
    },

    // ---- the organization's Claude API key (providers, provider anthropic) ----

    async claudeKey(org: string) {
      const [row] = await sql<
        {
          credential: Buffer;
          credentialKeyId: string;
          config: Record<string, string>;
          connectedBy: string | null;
          connectedAt: string;
        }[]
      >`
        select credential, credential_key_id as "credentialKeyId", config,
               connected_by as "connectedBy",
               to_char(connected_at at time zone 'UTC', 'YYYY-MM-DD"T"HH24:MI:SS.MS"Z"') as "connectedAt"
          from ${t.providers}
         where organization_id = ${org} and provider = 'anthropic'`;
      return row ?? null;
    },

    async setClaudeKey(input: {
      organizationId: string;
      config: Record<string, string>;
      credential: Buffer;
      credentialKeyId: string;
      credentialScope: string;
      connectedBy: string;
    }): Promise<void> {
      await sql`
        insert into ${t.providers}
          (organization_id, provider, status, config, credential, credential_key_id,
           credential_scope, account_label, connected_by)
        values (${input.organizationId}, 'anthropic', 'connected', ${sql.json(input.config)},
                ${input.credential}, ${input.credentialKeyId}, ${input.credentialScope},
                ${input.config.workspace ?? null}, ${input.connectedBy})
        on conflict (organization_id, provider) do update set
          status = 'connected', config = excluded.config, credential = excluded.credential,
          credential_key_id = excluded.credential_key_id,
          credential_scope = excluded.credential_scope, connected_by = excluded.connected_by,
          connected_at = now(), checked_at = now()`;
    },

    async removeClaudeKey(org: string): Promise<boolean> {
      const rows = await sql`
        delete from ${t.providers}
         where organization_id = ${org} and provider = 'anthropic' returning id`;
      return rows.length > 0;
    },
  };
}

export type TeamsStore = ReturnType<typeof createTeamsStore>;
