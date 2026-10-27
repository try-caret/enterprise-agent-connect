-- Rolls an install back to the earlier layout: the tables back in public under their old
-- names, pointing at public.organizations, and the service's schema gone. By hand, only to run
-- an older image again, in one transaction:
--   psql "$DATABASE_URL" -v ON_ERROR_STOP=1 -1 -f migrations/rollback/to-legacy.sql
-- Written for the default schema agent_connect. Fails, changing nothing, if an organization
-- here is missing from public.organizations.

-- The views migrate() made for the previous release, in the way of the tables' old names.
drop view if exists public.agent_connect_agents, public.agent_connect_providers,
  public.agent_connect_connections, public.agent_connect_grants, public.agent_connect_audit,
  public.agent_connect_confirmed, public.agent_connect_teams_bots,
  public.agent_connect_teams_conversations, public.agent_connect_teams_events,
  public.agent_connect_teams_triggers, public.agent_connect_teams_runs,
  public.agent_connect_teams_spend;
-- 0006: approval cards, which 0.7.5 does not have.
drop table agent_connect.approvals;
-- 0005: the runner of each session.
alter table agent_connect.teams_runs drop column runner;
-- 0004: triggers of other runners than Claude Managed Agents, which 0.7.5 cannot run, go.
delete from agent_connect.teams_triggers where runner <> 'claude-managed-agents';
alter table agent_connect.teams_triggers
  drop constraint teams_triggers_claude_ids,
  drop column runner,
  drop column settings,
  drop column secret,
  drop column secret_key_id,
  alter column managed_agent_id set not null,
  alter column environment_id set not null;
alter table agent_connect.connections add column access_ceiling text;
alter table agent_connect.grants add column access_level text;
drop policy service on agent_connect.organizations;
drop policy service on agent_connect.agents;
drop policy service on agent_connect.providers;
drop policy service on agent_connect.connections;
drop policy service on agent_connect.grants;
drop policy service on agent_connect.audit;
drop policy service on agent_connect.confirmed;
drop policy service on agent_connect.teams_bots;
drop policy service on agent_connect.teams_conversations;
drop policy service on agent_connect.teams_events;
drop policy service on agent_connect.teams_triggers;
drop policy service on agent_connect.teams_runs;
drop policy service on agent_connect.teams_spend;
alter table agent_connect.agents drop constraint agent_connect_agents_organization_id_fkey,
  add constraint agent_connect_agents_organization_id_fkey
  foreign key (organization_id) references public.organizations (id) on delete restrict;
alter table agent_connect.providers drop constraint agent_connect_providers_organization_id_fkey,
  add constraint agent_connect_providers_organization_id_fkey
  foreign key (organization_id) references public.organizations (id) on delete restrict;
alter table agent_connect.connections drop constraint agent_connect_connections_organization_id_fkey,
  add constraint agent_connect_connections_organization_id_fkey
  foreign key (organization_id) references public.organizations (id) on delete restrict;
alter table agent_connect.audit drop constraint agent_connect_audit_organization_id_fkey,
  add constraint agent_connect_audit_organization_id_fkey
  foreign key (organization_id) references public.organizations (id) on delete restrict;
alter table agent_connect.confirmed drop constraint agent_connect_confirmed_organization_id_fkey,
  add constraint agent_connect_confirmed_organization_id_fkey
  foreign key (organization_id) references public.organizations (id) on delete restrict;
alter table agent_connect.teams_bots drop constraint agent_connect_teams_bots_organization_id_fkey,
  add constraint agent_connect_teams_bots_organization_id_fkey
  foreign key (organization_id) references public.organizations (id) on delete restrict;
alter table agent_connect.agents rename to agent_connect_agents;
alter table agent_connect.providers rename to agent_connect_providers;
alter table agent_connect.connections rename to agent_connect_connections;
alter table agent_connect.grants rename to agent_connect_grants;
alter table agent_connect.audit rename to agent_connect_audit;
alter table agent_connect.confirmed rename to agent_connect_confirmed;
alter table agent_connect.teams_bots rename to agent_connect_teams_bots;
alter table agent_connect.teams_conversations rename to agent_connect_teams_conversations;
alter table agent_connect.teams_events rename to agent_connect_teams_events;
alter table agent_connect.teams_triggers rename to agent_connect_teams_triggers;
alter table agent_connect.teams_runs rename to agent_connect_teams_runs;
alter table agent_connect.teams_spend rename to agent_connect_teams_spend;
alter table agent_connect.agent_connect_agents set schema public;
alter table agent_connect.agent_connect_providers set schema public;
alter table agent_connect.agent_connect_connections set schema public;
alter table agent_connect.agent_connect_grants set schema public;
alter table agent_connect.agent_connect_audit set schema public;
alter table agent_connect.agent_connect_confirmed set schema public;
alter table agent_connect.agent_connect_teams_bots set schema public;
alter table agent_connect.agent_connect_teams_conversations set schema public;
alter table agent_connect.agent_connect_teams_events set schema public;
alter table agent_connect.agent_connect_teams_triggers set schema public;
alter table agent_connect.agent_connect_teams_runs set schema public;
alter table agent_connect.agent_connect_teams_spend set schema public;
drop table agent_connect.organizations, agent_connect.schema_migrations;
drop schema agent_connect;
