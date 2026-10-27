-- The earlier layout, frozen: apps an organization connected once for its agents, the
-- agents (API keys) allowed to use them, what each agent may do in each app, a record of every
-- call, and Microsoft Teams bots for agents. Only the agent-connect service reads and writes
-- these: RLS is forced and every grant to Supabase's roles revoked. Create only: nothing here
-- reads or rewrites existing rows.

-- An agent authenticates with a key only it holds; we keep its SHA-256. last_used_at is
-- stamped with every audit row, so listing keys does not aggregate the log.
create table public.agent_connect_agents (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete restrict,
  name text not null check (length(name) between 1 and 200),
  key_hash text not null unique check (key_hash ~ '^[0-9a-f]{64}$'),
  status text not null check (status in ('active', 'disabled')),
  created_by text,
  created_at timestamptz not null default now(),
  last_used_at timestamptz,
  unique (organization_id, name)
);

-- One sign-in per organization shared by several apps (Microsoft), and the organization's
-- Claude API key (provider 'anthropic'). The credential is AES-256-GCM sealed by the service
-- under credential_scope; credential_key_id names the key that sealed it.
create table public.agent_connect_providers (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete restrict,
  provider text not null check (provider ~ '^[a-z0-9-]+$'),
  status text not null check (status in ('pending', 'connected', 'needs_reconnect')),
  config jsonb not null default '{}'::jsonb,
  credential bytea not null,
  credential_key_id text not null,
  credential_scope text not null,
  account_label text,
  connected_by text,
  connected_at timestamptz not null default now(),
  checked_at timestamptz not null default now(),
  unique (organization_id, provider)
);

-- The deployment's shared app reaches any tenant that approved it, so one tenant binds to one
-- organization. A customer's own app reaches only its own tenant.
create unique index agent_connect_providers_one_org_per_tenant
  on public.agent_connect_providers (provider, (config->>'tenantId'))
  where config->>'auth' = 'arbium_app';

-- One row per connected app: capabilities is the most any agent may do there; config holds
-- non-secret values. An app signs in with its own sealed credential or through a provider
-- row, never both. access_ceiling is unused (nullable, never written).
create table public.agent_connect_connections (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete restrict,
  connector_id text not null check (connector_id ~ '^[a-z0-9-]+$'),
  access_ceiling text,
  config jsonb not null default '{}'::jsonb,
  credential bytea,
  credential_key_id text,
  status text not null check (status in ('connected', 'needs_reconnect')),
  account_label text,
  connected_by text,
  connected_at timestamptz not null default now(),
  checked_at timestamptz not null default now(),
  capabilities text[] not null,
  provider_id uuid references public.agent_connect_providers(id) on delete cascade,
  unique (organization_id, connector_id),
  constraint agent_connect_connections_one_credential
    check ((provider_id is null) = (credential is not null and credential_key_id is not null))
);

-- Each agent's ticked capabilities per app (the effective set is these within the
-- connection's), plus value limits such as {"projects": ["KAN"]}. access_level is unused.
create table public.agent_connect_grants (
  agent_id uuid not null references public.agent_connect_agents(id) on delete cascade,
  connector_id text not null check (connector_id ~ '^[a-z0-9-]+$'),
  access_level text,
  limits jsonb not null default '{}'::jsonb check (jsonb_typeof(limits) = 'object'),
  updated_at timestamptz not null default now(),
  capabilities text[] not null,
  primary key (agent_id, connector_id)
);

-- Vendors only see the shared connection, so this is the record of which agent did what. Rows
-- keep the agent's name after the agent is removed; admin changes have no agent.
create table public.agent_connect_audit (
  id bigint generated always as identity primary key,
  at timestamptz not null default now(),
  organization_id uuid not null references public.organizations(id) on delete restrict,
  agent_id uuid references public.agent_connect_agents(id) on delete set null,
  connector_id text not null,
  tool text not null,
  target text,
  decision text not null check (decision in ('allowed', 'denied', 'failed')),
  reason text,
  vendor_status integer,
  agent_name text
);
create index agent_connect_audit_org_at on public.agent_connect_audit (organization_id, at desc);

-- Setup steps the service cannot check live, marked done by an admin; steps of different apps with
-- the same id share one mark.
create table public.agent_connect_confirmed (
  organization_id uuid not null references public.organizations(id) on delete restrict,
  step text not null check (step ~ '^[a-z0-9-]+$'),
  confirmed_by text not null,
  confirmed_at timestamptz not null default now(),
  primary key (organization_id, step)
);

-- Teams (design: docs/teams.md). One bot per agent; its private key is made by
-- the service and sealed like other credentials.
create table public.agent_connect_teams_bots (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete restrict,
  agent_id uuid not null unique references public.agent_connect_agents(id) on delete cascade,
  -- The Entra application id; null until the setup script made the app and it was proven.
  bot_app_id uuid unique,
  -- The Teams app manifest id, minted here once.
  teams_app_id uuid not null unique default gen_random_uuid(),
  tenant_id uuid not null,
  display_name text not null check (length(display_name) between 1 and 30),
  -- The manifest's version is 1.0.<manifest_version>.
  manifest_version int not null default 1 check (manifest_version >= 1),
  status text not null check (status in ('awaiting_setup', 'connected', 'needs_reconnect')),
  -- {certificate, nextCertificate?}; the private keys are in the sealed credential.
  config jsonb not null default '{}'::jsonb,
  credential bytea not null,
  credential_key_id text not null,
  created_by text,
  created_at timestamptz not null default now(),
  checked_at timestamptz,
  -- The last request refused for this bot, in plain words, so the admin sees why nothing
  -- arrives.
  last_refused_at timestamptz,
  last_refused_reason text
);

-- Where the bot is installed. A team install is one row (the bot sees all its channels);
-- team_id is Teams' thread id for the team, team_group_id its Microsoft 365 group id.
create table public.agent_connect_teams_conversations (
  bot_id uuid not null references public.agent_connect_teams_bots(id) on delete cascade,
  conversation_id text not null,
  kind text not null check (kind in ('personal', 'groupChat', 'team')),
  service_url text not null check (service_url like 'https://%'),
  team_id text,
  team_group_id uuid,
  name text,
  user_aad_id uuid,
  installed_at timestamptz not null default now(),
  removed_at timestamptz,
  -- The last "limit reached" notice sent there, at most one an hour.
  limit_notice_at timestamptz,
  primary key (bot_id, conversation_id)
);

-- Dedupe and work queue. The activity is sealed (it holds the message text) and cleared once
-- processed. Pruned after 30 days by the service.
create table public.agent_connect_teams_events (
  id bigint generated always as identity primary key,
  bot_id uuid not null references public.agent_connect_teams_bots(id) on delete cascade,
  activity_id text not null,
  type text not null,
  conversation_id text not null,
  thread_id text,
  from_aad_id text,
  received_at timestamptz not null default now(),
  state text not null check (state in ('pending', 'processing', 'held', 'ignored', 'run', 'done', 'failed')),
  reason text,
  payload bytea,
  payload_key_id text,
  claimed_at timestamptz,
  unique (bot_id, activity_id)
);
create index agent_connect_teams_events_pending on public.agent_connect_teams_events (id)
  where state = 'pending';
create index agent_connect_teams_events_held on public.agent_connect_teams_events (bot_id, thread_id)
  where state = 'held';

-- Each agent's real-time response settings.
create table public.agent_connect_teams_triggers (
  agent_id uuid primary key references public.agent_connect_agents(id) on delete cascade,
  enabled boolean not null default false,
  managed_agent_id text not null check (managed_agent_id ~ '^[A-Za-z0-9_-]{1,128}$'),
  environment_id text not null check (environment_id ~ '^[A-Za-z0-9_-]{1,128}$'),
  vault_ids text[] not null default '{}',
  budget_cents int not null check (budget_cents between 1 and 2000),
  max_runs_per_hour int not null default 20 check (max_runs_per_hour between 1 and 100),
  -- What this agent's answers may cost together in any 24 hours.
  max_cents_per_day int not null default 500 check (max_cents_per_day between 1 and 100000),
  max_concurrent int not null default 2 check (max_concurrent between 1 and 10),
  reply_mode text not null default 'mentions' check (reply_mode in ('mentions', 'all')),
  allowed_teams uuid[] not null default '{}',
  allowed_people uuid[] not null default '{}',
  -- An empty allowed_people means nobody unless this is set: anyone in the organization (never
  -- guests) may then ask the agent, which acts with everything the agent was granted.
  anyone_in_organization boolean not null default false,
  check (anyone_in_organization or cardinality(allowed_people) > 0 or not enabled),
  updated_by text,
  updated_at timestamptz not null default now()
);

-- One Claude Managed Agents session per thread, reused for later messages there. The reply
-- token is for the latest message only.
create table public.agent_connect_teams_runs (
  id uuid primary key default gen_random_uuid(),
  agent_id uuid not null references public.agent_connect_agents(id) on delete cascade,
  bot_id uuid not null references public.agent_connect_teams_bots(id) on delete cascade,
  conversation_id text not null,
  thread_id text not null,
  trigger_activity_id text not null,
  session_id text,
  reply_token_hash text not null unique check (reply_token_hash ~ '^[0-9a-f]{64}$'),
  reply_token_at timestamptz not null default now(),
  -- no_reply: the session finished without answering in Teams.
  state text not null check (state in ('starting', 'running', 'idle', 'ended', 'failed', 'no_reply')),
  error text,
  -- Something the admin should know that did not stop the answer.
  note text,
  -- Who asked: a session is only reused for the same person.
  requested_by_aad_id uuid,
  -- Replies sent with the current reply token (a few per message at most), and the last one.
  reply_count int not null default 0,
  replied_at timestamptz,
  cost_cents int,
  turns int not null default 1,
  started_at timestamptz not null default now(),
  finished_at timestamptz
);
create index agent_connect_teams_runs_agent on public.agent_connect_teams_runs (agent_id, started_at desc);
-- One active session per thread: a second message while one runs is held, not a second run.
create unique index agent_connect_teams_runs_one_active
  on public.agent_connect_teams_runs (bot_id, thread_id) where state in ('starting', 'running');

-- What each Teams answer cost, per turn (a run's cost_cents is its session's running total);
-- the 24-hour spend limit sums these.
create table public.agent_connect_teams_spend (
  id bigint generated always as identity primary key,
  agent_id uuid not null references public.agent_connect_agents(id) on delete cascade,
  run_id uuid not null references public.agent_connect_teams_runs(id) on delete cascade,
  cents int not null check (cents > 0),
  spent_at timestamptz not null default now()
);
create index agent_connect_teams_spend_agent
  on public.agent_connect_teams_spend (agent_id, spent_at desc);

alter table public.agent_connect_agents enable row level security;
alter table public.agent_connect_agents force row level security;
alter table public.agent_connect_providers enable row level security;
alter table public.agent_connect_providers force row level security;
alter table public.agent_connect_connections enable row level security;
alter table public.agent_connect_connections force row level security;
alter table public.agent_connect_grants enable row level security;
alter table public.agent_connect_grants force row level security;
alter table public.agent_connect_audit enable row level security;
alter table public.agent_connect_audit force row level security;
alter table public.agent_connect_confirmed enable row level security;
alter table public.agent_connect_confirmed force row level security;
alter table public.agent_connect_teams_bots enable row level security;
alter table public.agent_connect_teams_bots force row level security;
alter table public.agent_connect_teams_conversations enable row level security;
alter table public.agent_connect_teams_conversations force row level security;
alter table public.agent_connect_teams_events enable row level security;
alter table public.agent_connect_teams_events force row level security;
alter table public.agent_connect_teams_triggers enable row level security;
alter table public.agent_connect_teams_triggers force row level security;
alter table public.agent_connect_teams_runs enable row level security;
alter table public.agent_connect_teams_runs force row level security;
alter table public.agent_connect_teams_spend enable row level security;
alter table public.agent_connect_teams_spend force row level security;
revoke all on public.agent_connect_agents from public,anon,authenticated,service_role;
revoke all on public.agent_connect_providers from public,anon,authenticated,service_role;
revoke all on public.agent_connect_connections from public,anon,authenticated,service_role;
revoke all on public.agent_connect_grants from public,anon,authenticated,service_role;
revoke all on public.agent_connect_audit from public,anon,authenticated,service_role;
revoke all on public.agent_connect_confirmed from public,anon,authenticated,service_role;
revoke all on public.agent_connect_teams_bots from public,anon,authenticated,service_role;
revoke all on public.agent_connect_teams_conversations from public,anon,authenticated,service_role;
revoke all on public.agent_connect_teams_events from public,anon,authenticated,service_role;
revoke all on public.agent_connect_teams_triggers from public,anon,authenticated,service_role;
revoke all on public.agent_connect_teams_runs from public,anon,authenticated,service_role;
revoke all on public.agent_connect_teams_spend from public,anon,authenticated,service_role;
