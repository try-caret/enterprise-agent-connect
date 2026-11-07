-- Files an agent sends with teams_reply in a direct chat, held (sealed) until the person accepts
-- or declines Teams' file consent card; on accept the file goes to their OneDrive. Gone after a day.
create table teams_file_offers (
  id uuid primary key default gen_random_uuid(),
  bot_id uuid not null references teams_bots (id) on delete cascade,
  conversation_id text not null,
  service_url text not null,
  -- Only this person's answer counts: the one who messaged the bot.
  recipient_aad_id uuid not null,
  name text not null,
  size integer not null check (size > 0),
  content bytea not null,
  content_key_id text not null,
  created_at timestamptz not null default now(),
  expires_at timestamptz not null
);
create index agent_connect_teams_file_offers_bot on teams_file_offers (bot_id);
create index agent_connect_teams_file_offers_expires on teams_file_offers (expires_at);
alter table teams_file_offers enable row level security;
alter table teams_file_offers force row level security;
create policy service on teams_file_offers for all to current_user using (true) with check (true);
