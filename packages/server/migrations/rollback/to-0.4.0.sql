-- Undoes migration 0014, so an image of 0.4.0 (whose last migration is 0013) starts again on a
-- database a later version migrated. By hand, only after the later version is stopped, in one
-- transaction, then restart the service:
--   psql "$DATABASE_URL" -v ON_ERROR_STOP=1 -1 -f migrations/rollback/to-0.4.0.sql
-- Clear any trigger conditions of the kinds 0.4.0 does not know (contains, anyOf, seen, logOnly)
-- first: 0.4.0 fails on them. Written for the default schema agent_connect.

do $$ begin
  if (select array_agg(version order by version) from agent_connect.schema_migrations
       where version > 13) is distinct from array[14] then
    raise exception 'agent_connect.schema_migrations does not end in exactly 0014 after 0013';
  end if;
end $$;

drop index agent_connect.agent_connect_event_runs_thread;
alter table agent_connect.event_runs drop column would_filter;
delete from agent_connect.schema_migrations where version = 14;
