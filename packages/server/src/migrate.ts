import { createHash } from "node:crypto";
import type postgres from "postgres";
import { readdirSync, readFileSync } from "node:fs";
import { assertSchemaName, type Sql } from "./store.js";

// The service's own schema migrations: plain numbered SQL in migrations/, written with
// unqualified names and run with search_path set to the configured schema, recorded in that
// schema's schema_migrations.

const DIR = new URL("../migrations/", import.meta.url);

// What the earlier layout had in public, made by the host's migration before the service had a schema.
// 0001 makes the same tables, so an install that has them adopts them as 0001.
const PREFIX = "agent_connect_";
const LEGACY = [
  "agents",
  "providers",
  "connections",
  "grants",
  "audit",
  "confirmed",
  "teams_bots",
  "teams_conversations",
  "teams_events",
  "teams_triggers",
  "teams_runs",
  "teams_spend",
].map((name) => `${PREFIX}${name}`);

// id: the file name without .sql, e.g. 0001_tables.
export type Migration = {
  id: string;
  version: number;
  name: string;
  sql: string;
  checksum: string;
};

export function migrationFiles(): Migration[] {
  return readdirSync(DIR)
    .filter((file) => /^\d{4}_[a-z0-9_]+\.sql$/.test(file))
    .sort()
    .map((file) => {
      const sql = readFileSync(new URL(file, DIR), "utf8");
      return {
        id: file.slice(0, -4),
        version: Number(file.slice(0, 4)),
        name: file.slice(5, -4),
        sql,
        checksum: createHash("sha256").update(sql).digest("hex"),
      };
    });
}

/** Applies pending migrations in one transaction; returns what it did, in order. */
export async function migrate(
  sql: Sql,
  schema: string,
  log: (line: string) => void = console.log,
): Promise<string[]> {
  assertSchemaName(schema);
  const files = migrationFiles();
  return sql.begin(async (tx) => {
    // One run at a time, e.g. two pods starting together; released with the transaction.
    await tx`select pg_advisory_xact_lock(hashtext(${`${schema}.migrate`}))`;
    // "Already exists, skipping" on every start is noise.
    await tx`set local client_min_messages = warning`;
    await tx.unsafe(`create schema if not exists ${schema}`);
    await tx.unsafe(`revoke all on schema ${schema} from public`);
    await tx.unsafe(`create table if not exists ${schema}.schema_migrations (
      version int primary key,
      name text not null,
      checksum text not null,
      adopted boolean not null default false,
      applied_at timestamptz not null default now())`);
    const applied = await tx.unsafe<{ version: number; checksum: string }[]>(
      `select version, checksum from ${schema}.schema_migrations order by version`,
    );
    for (const row of applied) {
      const file = files.find((f) => f.version === row.version);
      if (!file)
        throw new Error(
          `Migration ${row.version} is applied in ${schema} but not known to this version`,
        );
      if (file.checksum !== row.checksum)
        throw new Error(`Migration ${file.id} changed after it was applied`);
    }

    const ran: string[] = [];
    const legacy = await legacyTables(tx);
    const adopting = legacy.length > 0 && !applied.length;
    if (adopting) {
      // Moved, not copied: the rows, ids and sequences stay as they are.
      for (const table of LEGACY)
        await tx.unsafe(`alter table public.${table} set schema ${schema}`);
      const first = files[0];
      await tx.unsafe(
        `insert into ${schema}.schema_migrations (version, name, checksum, adopted)
         values ($1, $2, $3, true)`,
        [first.version, first.name, first.checksum],
      );
      applied.push({ version: first.version, checksum: first.checksum });
      ran.push(`adopted public.agent_connect_* as ${first.id}`);
    } else if (legacy.length) {
      // The host's migration made them after this schema existed: a fresh install where the
      // service started before the host's migrations ran. Empty, they are dropped; rows are not.
      const filled = [];
      for (const table of legacy) {
        // FORCE RLS with no policies hides every row from an owner without bypassrls, and
        // this check is all that stands before the drop. Rolled back with the transaction.
        await tx.unsafe(`alter table public.${table} no force row level security`);
        const [{ any }] = await tx.unsafe(`select exists (select from public.${table}) as any`);
        if (any) filled.push(table);
      }
      if (filled.length)
        throw new Error(
          `public.${filled.join(", public.")} hold rows, but ${schema} already has the service's tables`,
        );
      await tx.unsafe(`drop table ${legacy.map((t) => `public.${t}`).join(", ")}`);
      ran.push(`dropped empty public.${legacy.join(", public.")}`);
    }

    await tx.unsafe(`set local search_path to ${schema}`);
    const done = new Set(applied.map((row) => row.version));
    for (const file of files.filter((f) => !done.has(f.version))) {
      await tx.unsafe(file.sql).simple();
      await tx.unsafe(
        `insert into ${schema}.schema_migrations (version, name, checksum) values ($1, $2, $3)`,
        [file.version, file.name, file.checksum],
      );
      ran.push(`applied ${file.id}`);
    }
    if (adopting) {
      // So that a rolled-back 0.7.x image, which reads public.agent_connect_*, keeps serving
      // until to-legacy.sql runs. Simple views, so its inserts, upserts and row locks go
      // through. TODO(next release): drop them in a migration before any column changes.
      for (const table of LEGACY)
        await tx.unsafe(
          `create view public.${table} as select * from ${schema}.${table.slice(PREFIX.length)}`,
        );
      // Default privileges on public can grant them to API roles; they bypass RLS, so only
      // their owner keeps them.
      const grants = await tx.unsafe<{ view: string; role: string }[]>(
        `select distinct c.relname as view, a.grantee::regrole::text as role
           from pg_class c join pg_namespace n on n.oid = c.relnamespace,
                aclexplode(c.relacl) a
          where n.nspname = 'public' and c.relname = any($1::text[])
            and a.grantee not in (0, c.relowner)`,
        [LEGACY],
      );
      for (const { view, role } of grants)
        await tx.unsafe(`revoke all on public.${view} from ${role}`);
      for (const table of LEGACY) await tx.unsafe(`revoke all on public.${table} from public`);
      ran.push(`made views public.${PREFIX}* for the previous release`);
    }
    for (const line of ran) log(`[agent-connect] migrations: ${line} (schema ${schema})`);
    return ran;
  });
}

async function legacyTables(tx: postgres.TransactionSql): Promise<string[]> {
  // Tables only: the views made on adoption have the same names. In LEGACY's order, not the
  // catalog's, which moves as other tables are made.
  const rows = await tx.unsafe<{ name: string }[]>(
    `select c.relname as name from pg_class c join pg_namespace n on n.oid = c.relnamespace
      where n.nspname = 'public' and c.relkind in ('r', 'p') and c.relname = any($1::text[])
      order by array_position($1::text[], c.relname::text)`,
    [LEGACY],
  );
  return rows.map((row) => row.name);
}
