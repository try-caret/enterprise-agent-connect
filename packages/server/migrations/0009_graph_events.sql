-- Microsoft Graph change notifications as a third source of app events: the organization's
-- subscription holds the clientState secret Graph sends back with every notification, and one
-- watch per watched value (an agent's mailbox) holds its Graph subscription and its catch-up
-- position. Runs of a threaded event (a mail conversation, an issue) record the thread's id.
alter table event_subscriptions
  drop constraint event_subscriptions_source_check,
  add constraint event_subscriptions_source_check
    check (source in ('webhook', 'poll', 'graph')),
  drop constraint event_subscriptions_check,
  add constraint event_subscriptions_check check ((source <> 'poll') = (secret is not null));

create table event_watches (
  id uuid primary key default gen_random_uuid(),
  subscription_id uuid not null references event_subscriptions (id) on delete cascade,
  -- The watched value of the app's `graph.per` limit, e.g. a mailbox address.
  value text not null,
  -- Graph's subscription id, null until made (or after Graph removed it).
  remote_id text,
  expires_at timestamptz,
  -- The catch-up's next delta link (path and query only).
  cursor text,
  last_notification_at timestamptz,
  caught_up_at timestamptz,
  last_error text,
  created_at timestamptz not null default now(),
  unique (subscription_id, value),
  unique (subscription_id, remote_id)
);

alter table event_runs add column thread text;

alter table event_watches enable row level security;
alter table event_watches force row level security;
create policy service on event_watches for all to current_user using (true) with check (true);
