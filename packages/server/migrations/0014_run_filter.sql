-- A run a trigger's log-only conditions would have filtered out (it started anyway). Such runs do
-- not count as the agent having seen the thread (the `seen` condition), so a log-only trial
-- reports what enforcing would drop.
alter table event_runs add column would_filter boolean not null default false;
create index agent_connect_event_runs_thread on event_runs (trigger_id, thread, started_at desc)
  where thread is not null;
