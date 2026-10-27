-- Approval cards: an agent asks one person, through its Teams bot, to approve or reject. The
-- answer is a click (adaptiveCard/action) in a request authenticated by a Bot Framework token,
-- accepted only from that person's Entra object id, in the team or chat the card was posted in.
create table approvals (
  id uuid primary key default gen_random_uuid(),
  agent_id uuid not null references agents (id) on delete cascade,
  bot_id uuid not null references teams_bots (id) on delete cascade,
  -- What the agent holds is a random id; only its SHA-256 is kept.
  token_hash text not null unique check (token_hash ~ '^[0-9a-f]{64}$'),
  -- Where the card is: the installed place, its Bot Connector address, the conversation to edit
  -- the card in (a channel thread or the chat) and the card's activity id (set once posted).
  conversation_id text not null,
  service_url text not null,
  thread_id text,
  activity_id text,
  approver_aad_id uuid not null,
  approver_name text not null,
  -- What the card says, to show it again with the decision.
  title text not null,
  body text not null,
  facts jsonb not null default '[]',
  status text not null default 'pending' check (status in ('pending', 'approved', 'rejected')),
  decided_by_aad_id uuid,
  decided_by_name text,
  decided_at timestamptz,
  comment text,
  created_at timestamptz not null default now(),
  expires_at timestamptz not null,
  check ((status = 'pending') = (decided_at is null))
);
create index agent_connect_approvals_bot on approvals (bot_id, created_at desc);
-- An agent's own lookups and the cascade when it is deleted.
create index agent_connect_approvals_agent on approvals (agent_id);
alter table approvals enable row level security;
alter table approvals force row level security;
create policy service on approvals for all to current_user using (true) with check (true);
