-- A run whose runner did not take it yet (no answer, 5xx, 408, 409, 429) stays starting and is
-- handed over again at next_attempt_at, with what to send sealed in `delivery` (the prompt and
-- reply token, the event's fields) until it starts or fails. attempts: hand-overs failed so far.
alter table teams_runs
  add column attempts int not null default 0,
  add column next_attempt_at timestamptz,
  add column delivery bytea,
  add column delivery_key_id text;
create index agent_connect_teams_runs_retry on teams_runs (next_attempt_at)
  where next_attempt_at is not null;

alter table event_runs
  add column attempts int not null default 0,
  add column next_attempt_at timestamptz,
  add column delivery bytea,
  add column delivery_key_id text;
create index agent_connect_event_runs_retry on event_runs (next_attempt_at)
  where next_attempt_at is not null;
