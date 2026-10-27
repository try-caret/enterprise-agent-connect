-- What answers a Teams trigger: Claude Managed Agents (every row before this, so the default)
-- or another runner, with its own settings and, for a signed webhook, a signing secret sealed
-- like other credentials. Only the Claude runner needs a Claude agent and environment.
alter table teams_triggers
  add column runner text not null default 'claude-managed-agents'
    check (runner ~ '^[a-z][a-z0-9-]{1,40}$'),
  add column settings jsonb not null default '{}',
  add column secret bytea,
  add column secret_key_id text,
  alter column managed_agent_id drop not null,
  alter column environment_id drop not null,
  add constraint teams_triggers_claude_ids check (
    runner <> 'claude-managed-agents'
    or (managed_agent_id is not null and environment_id is not null)
  );
