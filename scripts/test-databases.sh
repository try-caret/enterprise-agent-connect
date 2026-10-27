#!/usr/bin/env bash
# Makes ac_test, the scratch database the DB-backed tests use (TEST_DATABASE_URL), on a
# disposable Postgres server: dropped and made again, empty. store.test.ts migrates it;
# upgrade.test.ts makes its own databases next to it, so the user needs CREATEDB.
# Usage: scripts/test-databases.sh postgres://postgres@localhost:5432
set -euo pipefail
server=${1:?usage: $0 postgres://user@host:port}
run() { psql "$server/$1" -v ON_ERROR_STOP=1 -q "${@:2}"; }

# The earlier layout's migration upgrade.test.ts starts from revokes from Supabase's roles.
run postgres -c "do \$\$ begin
  if not exists (select from pg_roles where rolname = 'anon') then create role anon; end if;
  if not exists (select from pg_roles where rolname = 'authenticated') then create role authenticated; end if;
  if not exists (select from pg_roles where rolname = 'service_role') then create role service_role; end if;
end \$\$"
run postgres -c "drop database if exists ac_test" -c "create database ac_test"
echo "ac_test ready on $server"
