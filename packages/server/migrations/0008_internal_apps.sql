-- An organization's own apps, added through the admin API: a connector.json (with an http,
-- soap, odata or sql block) kept as written, run on the generic module of its kind. Never
-- holds a secret: those are entered at connect time and sealed in connections.
create table internal_apps (
  organization_id uuid not null references organizations (id) on delete cascade,
  id text not null check (id ~ '^[a-z0-9-]+$'),
  definition jsonb not null,
  created_by text not null,
  created_at timestamptz not null default now(),
  updated_by text not null,
  updated_at timestamptz not null default now(),
  primary key (organization_id, id)
);
alter table internal_apps enable row level security;
alter table internal_apps force row level security;
create policy service on internal_apps for all to current_user using (true) with check (true);
