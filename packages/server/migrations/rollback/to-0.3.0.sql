-- Undoes migrations 0009 to 0013, so an image of 0.3.0 (whose last migration is 0008) starts
-- again on a database a later version migrated: an image refuses a database with a migration it
-- does not know. By hand, only after the later version is stopped (one that starts again applies
-- them again), in one transaction, then restart the service:
--   psql "$DATABASE_URL" -v ON_ERROR_STOP=1 -1 -f migrations/rollback/to-0.3.0.sql
-- Written for the default schema agent_connect. Fails, changing nothing, unless exactly 0009 to
-- 0013 are applied after 0008. Every row 0.3.0 reads is kept. Gone for good: Graph event
-- subscriptions (0.3.0 has no such source; their app events and watches go with them, and Graph
-- keeps sending to the old address until its subscriptions lapse, at most about three days),
-- threads of event runs, files of Teams runs, files offered in direct chats, and runs waiting to
-- be handed to their runner again (0.3.0 fails them as "Did not start" after 10 minutes).

do $$ begin
  if (select array_agg(version order by version) from agent_connect.schema_migrations
       where version > 8) is distinct from array[9, 10, 11, 12, 13] then
    raise exception 'agent_connect.schema_migrations does not end in exactly 0009 to 0013';
  end if;
end $$;

-- 0013: who is doing the work every process would otherwise do at once.
drop table agent_connect.leases;
-- 0012: runs waiting to be handed to their runner again (the retry indexes go with the columns).
alter table agent_connect.teams_runs drop column attempts, drop column next_attempt_at,
  drop column delivery, drop column delivery_key_id;
alter table agent_connect.event_runs drop column attempts, drop column next_attempt_at,
  drop column delivery, drop column delivery_key_id;
-- 0011: files offered in direct chats, waiting for an answer.
drop table agent_connect.teams_file_offers;
-- 0010: the files of each run's latest message.
alter table agent_connect.teams_runs drop column files, drop column files_key_id;
-- 0009: Graph notification watches, run threads, and the graph source with its subscriptions.
drop table agent_connect.event_watches;
alter table agent_connect.event_runs drop column thread;
delete from agent_connect.event_subscriptions where source = 'graph';
alter table agent_connect.event_subscriptions
  drop constraint event_subscriptions_source_check,
  add constraint event_subscriptions_source_check check (source in ('webhook', 'poll')),
  drop constraint event_subscriptions_check,
  add constraint event_subscriptions_check check ((source = 'webhook') = (secret is not null));

delete from agent_connect.schema_migrations where version between 9 and 13;
