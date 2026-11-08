-- The files that came with a run's latest message (a direct chat's download links, a channel's
-- SharePoint addresses), sealed, for teams_get_attachment. Replaced with each message.
alter table teams_runs
  add column files bytea,
  add column files_key_id text;
