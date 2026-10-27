-- The service's own organizations, and table names without the prefix they needed in public.
-- Index and constraint names keep their old spelling (a rename leaves them), so code matching
-- one works on every install.

create table organizations (
  id uuid primary key default gen_random_uuid(),
  -- Null for organizations taken over from rows saved before this table existed.
  name text check (name is null or length(name) between 1 and 200),
  created_at timestamptz not null default now()
);
alter table organizations enable row level security;
alter table organizations force row level security;

alter table agent_connect_agents rename to agents;
alter table agent_connect_providers rename to providers;
alter table agent_connect_connections rename to connections;
alter table agent_connect_grants rename to grants;
alter table agent_connect_audit rename to audit;
alter table agent_connect_confirmed rename to confirmed;
alter table agent_connect_teams_bots rename to teams_bots;
alter table agent_connect_teams_conversations rename to teams_conversations;
alter table agent_connect_teams_events rename to teams_events;
alter table agent_connect_teams_triggers rename to teams_triggers;
alter table agent_connect_teams_runs rename to teams_runs;
alter table agent_connect_teams_spend rename to teams_spend;

-- RLS stays forced. A role that owns these tables but cannot bypass RLS (one owning only this
-- schema) needs a policy to see any row, before the backfill below reads them; the role running
-- migrations is the service's own.
create policy service on organizations for all to current_user using (true) with check (true);
create policy service on agents for all to current_user using (true) with check (true);
create policy service on providers for all to current_user using (true) with check (true);
create policy service on connections for all to current_user using (true) with check (true);
create policy service on grants for all to current_user using (true) with check (true);
create policy service on audit for all to current_user using (true) with check (true);
create policy service on confirmed for all to current_user using (true) with check (true);
create policy service on teams_bots for all to current_user using (true) with check (true);
create policy service on teams_conversations for all to current_user using (true) with check (true);
create policy service on teams_events for all to current_user using (true) with check (true);
create policy service on teams_triggers for all to current_user using (true) with check (true);
create policy service on teams_runs for all to current_user using (true) with check (true);
create policy service on teams_spend for all to current_user using (true) with check (true);

-- Every organization already holding rows keeps its id.
insert into organizations (id)
select organization_id from agents
union select organization_id from providers
union select organization_id from connections
union select organization_id from audit
union select organization_id from confirmed
union select organization_id from teams_bots
on conflict do nothing;

-- An install adopted from public still points these at public.organizations.
alter table agents drop constraint if exists agent_connect_agents_organization_id_fkey,
  add constraint agent_connect_agents_organization_id_fkey
  foreign key (organization_id) references organizations (id) on delete restrict;
alter table providers drop constraint if exists agent_connect_providers_organization_id_fkey,
  add constraint agent_connect_providers_organization_id_fkey
  foreign key (organization_id) references organizations (id) on delete restrict;
alter table connections drop constraint if exists agent_connect_connections_organization_id_fkey,
  add constraint agent_connect_connections_organization_id_fkey
  foreign key (organization_id) references organizations (id) on delete restrict;
alter table audit drop constraint if exists agent_connect_audit_organization_id_fkey,
  add constraint agent_connect_audit_organization_id_fkey
  foreign key (organization_id) references organizations (id) on delete restrict;
alter table confirmed drop constraint if exists agent_connect_confirmed_organization_id_fkey,
  add constraint agent_connect_confirmed_organization_id_fkey
  foreign key (organization_id) references organizations (id) on delete restrict;
alter table teams_bots drop constraint if exists agent_connect_teams_bots_organization_id_fkey,
  add constraint agent_connect_teams_bots_organization_id_fkey
  foreign key (organization_id) references organizations (id) on delete restrict;
