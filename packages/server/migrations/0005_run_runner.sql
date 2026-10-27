-- The runner that started each session, so a trigger moved to another runner never hands a
-- session to a runner that did not start it. Every run before this was Claude's.
alter table teams_runs add column runner text not null default 'claude-managed-agents';
