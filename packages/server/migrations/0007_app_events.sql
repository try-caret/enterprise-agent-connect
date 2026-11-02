-- Events from connected apps (a Jira issue created, ...) that start agents. An organization
-- subscribes once per app, by webhook (the app posts to /agent-connect/events/<id>, signed with
-- the subscription's secret) or by polling the app on an interval from a cursor.
create table event_subscriptions (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references organizations (id) on delete cascade,
  connector_id text not null,
  source text not null check (source in ('webhook', 'poll')),
  -- The webhook's secret, sealed (event-subscription:<id>); null for a poll.
  secret bytea,
  secret_key_id text,
  -- Narrows what the app sends or what a poll reads, in the app's own terms (Jira: JQL).
  filter text not null default '',
  interval_seconds int not null default 300 check (interval_seconds between 60 and 86400),
  -- A poll's position: the newest change it has read.
  cursor text,
  polled_at timestamptz,
  last_error text,
  -- Where the app keeps the webhook when it was registered through its API.
  registered text,
  created_by text not null,
  created_at timestamptz not null default now(),
  unique (organization_id, connector_id),
  check ((source = 'webhook') = (secret is not null))
);

-- Deliveries waiting to be matched against triggers, and those done (kept 30 days, so a signed
-- delivery cannot be replayed while its timestamp would still pass). Only the fields the
-- connector keeps are stored, sealed, and only until processed.
create table app_events (
  id bigint generated always as identity primary key,
  subscription_id uuid not null references event_subscriptions (id) on delete cascade,
  delivery_id text not null,
  event_type text not null,
  state text not null default 'pending'
    check (state in ('pending', 'processing', 'done', 'failed')),
  reason text,
  payload bytea,
  payload_key_id text,
  received_at timestamptz not null default now(),
  claimed_at timestamptz,
  unique (subscription_id, delivery_id)
);
create index agent_connect_app_events_pending on app_events (id)
  where state in ('pending', 'processing');
create index agent_connect_app_events_received on app_events (received_at);

-- What an agent does when an app's event arrives: which events, conditions on their fields,
-- its instructions, what runs it (as teams_triggers: runner, settings, sealed signing secret,
-- the Claude agent) and its caps.
create table event_triggers (
  id uuid primary key default gen_random_uuid(),
  agent_id uuid not null references agents (id) on delete cascade,
  connector_id text not null,
  enabled boolean not null default false,
  event_types text[] not null check (cardinality(event_types) > 0),
  conditions jsonb not null default '[]',
  prompt text not null check (length(prompt) between 1 and 4000),
  runner text not null check (runner ~ '^[a-z][a-z0-9-]{1,40}$'),
  settings jsonb not null default '{}',
  secret bytea,
  secret_key_id text,
  managed_agent_id text check (managed_agent_id ~ '^[A-Za-z0-9_-]{1,128}$'),
  environment_id text check (environment_id ~ '^[A-Za-z0-9_-]{1,128}$'),
  vault_ids text[] not null default '{}',
  budget_cents int not null check (budget_cents between 1 and 2000),
  max_runs_per_hour int not null default 20 check (max_runs_per_hour between 1 and 100),
  max_cents_per_day int not null default 500 check (max_cents_per_day between 1 and 100000),
  max_concurrent int not null default 2 check (max_concurrent between 1 and 10),
  updated_by text,
  updated_at timestamptz not null default now(),
  unique (agent_id, connector_id),
  check (
    runner <> 'claude-managed-agents'
    or (managed_agent_id is not null and environment_id is not null)
  )
);
create index agent_connect_event_triggers_app on event_triggers (connector_id);

create table event_runs (
  id uuid primary key default gen_random_uuid(),
  agent_id uuid not null references agents (id) on delete cascade,
  trigger_id uuid not null references event_triggers (id) on delete cascade,
  event_id bigint references app_events (id) on delete set null,
  event_type text not null,
  delivery_id text not null,
  runner text not null,
  session_id text,
  state text not null check (state in ('starting', 'running', 'ended', 'failed')),
  error text,
  cost_cents int,
  started_at timestamptz not null default now(),
  finished_at timestamptz
);
create index agent_connect_event_runs_trigger on event_runs (trigger_id, started_at desc);
create index agent_connect_event_runs_agent on event_runs (agent_id, started_at desc);
create index agent_connect_event_runs_live on event_runs (state) where state in ('starting', 'running');

alter table event_subscriptions enable row level security;
alter table event_subscriptions force row level security;
create policy service on event_subscriptions for all to current_user using (true) with check (true);
alter table app_events enable row level security;
alter table app_events force row level security;
create policy service on app_events for all to current_user using (true) with check (true);
alter table event_triggers enable row level security;
alter table event_triggers force row level security;
create policy service on event_triggers for all to current_user using (true) with check (true);
alter table event_runs enable row level security;
alter table event_runs force row level security;
create policy service on event_runs for all to current_user using (true) with check (true);
