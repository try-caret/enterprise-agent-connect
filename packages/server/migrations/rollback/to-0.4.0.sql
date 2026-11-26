-- Undoes migration 0014, so an image of 0.4.0 (whose last migration is 0013) starts again on a
-- database a later version migrated: an image refuses a database with a migration it does not
-- know. By hand, only after the later version is stopped (one that starts again applies it again),
-- in one transaction, then restart the service:
--   psql "$DATABASE_URL" -v ON_ERROR_STOP=1 -1 -f migrations/rollback/to-0.4.0.sql
-- Written for the default schema agent_connect. Fails, changing nothing, unless exactly 0014 is
-- applied after 0013, or while a trigger holds a condition 0.4.0 cannot read (contains, anyOf,
-- seen, logOnly) or reads otherwise ({self} in a list, which 0.4.0 takes as text): set those
-- triggers' conditions back first. Every row 0.4.0 reads is kept. Gone
-- for good: which event runs a log-only trial would have filtered.

do $$ begin
  if (select array_agg(version order by version) from agent_connect.schema_migrations
       where version > 13) is distinct from array[14] then
    raise exception 'agent_connect.schema_migrations does not end in exactly 0014';
  end if;
  if exists (select 1 from agent_connect.event_triggers t, jsonb_array_elements(t.conditions) c
              where not (c ? 'field' and (c ? 'in' or c ? 'notIn'))
                 or c->'in' ? '{self}' or c->'notIn' ? '{self}') then
    raise exception 'a trigger holds a condition 0.4.0 cannot read; set its conditions back first';
  end if;
end $$;

-- 0014: which event runs a log-only trial would have filtered.
alter table agent_connect.event_runs drop column would_filter;

delete from agent_connect.schema_migrations where version = 14;
