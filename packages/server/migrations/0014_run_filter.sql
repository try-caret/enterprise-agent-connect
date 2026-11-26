-- Whether an event run started only because the conditions it failed were log-only (a trial of a
-- filter). A trial's `seen` leaves such runs out, so it reports what enforcing it would drop.
alter table event_runs add column would_filter boolean not null default false;
