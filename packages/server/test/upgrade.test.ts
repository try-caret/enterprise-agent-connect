import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { after, before, test } from "node:test";
import postgres from "postgres";
import { migrate, migrationFiles } from "../src/migrate.ts";
import { credentialAad, CredentialKey, hashAgentKey } from "../src/secrets.ts";
import { connectDatabase, createStore, type Sql } from "../src/store.ts";

// An install of the earlier layout (tables in public, made by the host's migration 20261023200000;
// frozen copy in fixtures/) upgrading to the service's own schema, against a fresh install. Makes its own
// databases next to TEST_DATABASE_URL's; skipped without one.
const url = process.env.TEST_DATABASE_URL;
const SCHEMA = "agent_connect";
const LEGACY_SQL = readFileSync(
  new URL("fixtures/legacy-20261023200000.sql", import.meta.url),
  "utf8",
);
const ROLLBACK_SQL = readFileSync(
  new URL("../migrations/rollback/to-legacy.sql", import.meta.url),
  "utf8",
);
const TABLES = [
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
];
const quiet = () => {};

const admin = url ? postgres(url, { max: 1, onnotice: () => {} }) : null;
const pools: Sql[] = [];
function connect(name: string): Sql {
  const at = new URL(url!);
  at.pathname = `/${name}`;
  const sql = connectDatabase(at.toString());
  pools.push(sql);
  return sql;
}
async function database(name: string): Promise<Sql> {
  await admin!.unsafe(`drop database if exists ${name} with (force)`);
  await admin!.unsafe(`create database ${name}`);
  return connect(name);
}

let legacy: Sql;
let fresh: Sql;
before(async () => {
  if (!admin) return;
  legacy = await database("ac_upgrade_legacy");
  fresh = await database("ac_upgrade_fresh");
});
after(async () => {
  for (const sql of pools) await sql.end();
  await admin?.unsafe("drop database if exists ac_upgrade_legacy with (force)");
  await admin?.unsafe("drop database if exists ac_upgrade_fresh with (force)");
  await admin?.end();
});

// Tables, columns (by name: the rollback re-adds dropped ones last), constraints, indexes,
// RLS and policies, and sequences of one schema. ACLs are left out.
async function catalog(sql: Sql, schema: string) {
  const tables = await sql`
    select c.relname as table, c.relrowsecurity as rls, c.relforcerowsecurity as forced,
      (select json_agg(json_build_object('name', a.attname,
                'type', format_type(a.atttypid, a.atttypmod), 'notnull', a.attnotnull,
                'default', pg_get_expr(d.adbin, d.adrelid), 'identity', a.attidentity)
              order by a.attname)
         from pg_attribute a left join pg_attrdef d on d.adrelid = a.attrelid and d.adnum = a.attnum
        where a.attrelid = c.oid and a.attnum > 0 and not a.attisdropped) as columns,
      (select json_agg(json_build_object('name', conname, 'def', pg_get_constraintdef(oid))
              order by conname)
         from pg_constraint where conrelid = c.oid) as constraints,
      (select json_agg(indexdef order by indexname)
         from pg_indexes where schemaname = n.nspname and tablename = c.relname) as indexes,
      (select json_agg(json_build_object('name', polname, 'cmd', polcmd,
                'roles', polroles::regrole[]::text, 'using', pg_get_expr(polqual, polrelid),
                'check', pg_get_expr(polwithcheck, polrelid)) order by polname)
         from pg_policy where polrelid = c.oid) as policies
      from pg_class c join pg_namespace n on n.oid = c.relnamespace
     where n.nspname = ${schema} and c.relkind = 'r'
       and (${schema} <> 'public' or c.relname like 'agent\\_connect\\_%')
     order by c.relname`;
  const sequences = await sql`
    select c.relname from pg_class c join pg_namespace n on n.oid = c.relnamespace
     where n.nspname = ${schema} and c.relkind = 'S' order by c.relname`;
  return { tables: [...tables], sequences: sequences.map((s) => s.relname) };
}

// Each table's row count and a hash of every row (without the two never-used columns 0003
// drops and the runner columns 0004 and 0005 add), so ids, values and ciphertexts are compared byte
// for byte.
async function contents(sql: Sql, name: (table: string) => string) {
  const out: Record<string, { n: number; hash: string }> = {};
  for (const table of TABLES) {
    const [row] = await sql.unsafe(
      `select count(*)::int as n,
              md5(coalesce(string_agg(r::text, ',' order by r::text), '')) as hash
         from (select to_jsonb(x) - 'access_ceiling' - 'access_level' - 'runner' - 'settings' - 'secret'
                       - 'secret_key_id' as r from ${name(table)} x) s`,
    );
    out[table] = { n: row.n, hash: row.hash };
  }
  return out;
}
const legacyName = (table: string) => `public.agent_connect_${table}`;
const movedName = (table: string) => `${SCHEMA}.${table}`;

const key = new CredentialKey(Buffer.alloc(32, 7).toString("base64"));
const ORG_A = "6d2b9f4e-1c3a-4e5b-8f7d-0a1b2c3d4e5f";
const ORG_B = "7e3c0a5f-2d4b-4f6c-9a8e-1b2c3d4e5f60";
const UNUSED_ORG = "8f4d1b60-3e5c-4a7d-8b9f-2c3d4e5f6071";
const AGENT = "9a5e2c71-4f6d-4b8e-9cab-3d4e5f607182";
const PROVIDER = "ab6f3d82-5a7e-4c9f-8dbc-4e5f60718293";
const BOT = "bc7a4e93-6b8f-4dab-9ecd-5f60718293a4";
const RUN = "cd8b5fa4-7c9a-4ebc-8fde-60718293a4b5";
const AGENT_KEY = `arb_agent_${"A".repeat(43)}`;

async function seedLegacy(sql: Sql) {
  await sql`create table public.organizations (id uuid primary key)`;
  await sql`insert into public.organizations (id) values (${ORG_A}), (${ORG_B}), (${UNUSED_ORG})`;
  await sql.unsafe(LEGACY_SQL).simple();
  const provider = key.seal(
    { client_secret: "provider" },
    credentialAad(ORG_A, "provider:microsoft"),
  );
  const jira = key.seal({ client_secret: "jira" }, credentialAad(ORG_A, "jira"));
  const bot = key.seal({ private_key: "bot" }, credentialAad(ORG_A, `teams-bot:${BOT}`));
  await sql`insert into public.agent_connect_agents (id, organization_id, name, key_hash, status)
            values (${AGENT}, ${ORG_A}, 'Triage', ${hashAgentKey(AGENT_KEY)}, 'active')`;
  await sql`insert into public.agent_connect_agents (organization_id, name, key_hash, status)
            values (${ORG_B}, 'Other', ${"b".repeat(64)}, 'disabled')`;
  await sql`insert into public.agent_connect_providers
              (id, organization_id, provider, status, config, credential, credential_key_id,
               credential_scope)
            values (${PROVIDER}, ${ORG_A}, 'microsoft', 'connected',
                    ${sql.json({ auth: "arbium_app", tenantId: ORG_B })}, ${provider.ciphertext},
                    ${provider.keyId}, 'provider:microsoft')`;
  await sql`insert into public.agent_connect_connections
              (organization_id, connector_id, access_ceiling, credential, credential_key_id,
               status, capabilities)
            values (${ORG_A}, 'jira', 'old', ${jira.ciphertext}, ${jira.keyId}, 'connected',
                    ${sql.array(["read", "write"])})`;
  await sql`insert into public.agent_connect_connections
              (organization_id, connector_id, provider_id, status, capabilities)
            values (${ORG_A}, 'outlook-mail', ${PROVIDER}, 'connected', ${sql.array(["read"])})`;
  await sql`insert into public.agent_connect_grants (agent_id, connector_id, access_level, limits, capabilities)
            values (${AGENT}, 'jira', 'old', ${sql.json({ projects: ["KAN"] })}, ${sql.array(["read"])})`;
  await sql`insert into public.agent_connect_audit
              (organization_id, agent_id, agent_name, connector_id, tool, decision)
            select ${ORG_A}, ${AGENT}, 'Triage', 'jira', 'jira_search_' || i, 'allowed'
              from generate_series(1, 100) i`;
  await sql`insert into public.agent_connect_confirmed (organization_id, step, confirmed_by)
            values (${ORG_B}, 'grant-consent', 'admin')`;
  await sql`insert into public.agent_connect_teams_bots
              (id, organization_id, agent_id, tenant_id, display_name, status, credential,
               credential_key_id)
            values (${BOT}, ${ORG_A}, ${AGENT}, ${ORG_B}, 'Triage', 'connected', ${bot.ciphertext},
                    ${bot.keyId})`;
  await sql`insert into public.agent_connect_teams_conversations
              (bot_id, conversation_id, kind, service_url)
            values (${BOT}, 'a:1', 'personal', 'https://smba.example/')`;
  await sql`insert into public.agent_connect_teams_events
              (bot_id, activity_id, type, conversation_id, state)
            values (${BOT}, 'act-1', 'message', 'a:1', 'done')`;
  await sql`insert into public.agent_connect_teams_triggers
              (agent_id, managed_agent_id, environment_id, budget_cents)
            values (${AGENT}, 'agent_1', 'env_1', 50)`;
  await sql`insert into public.agent_connect_teams_runs
              (id, agent_id, bot_id, conversation_id, thread_id, trigger_activity_id,
               reply_token_hash, state)
            values (${RUN}, ${AGENT}, ${BOT}, 'a:1', 'a:1', 'act-1', ${"c".repeat(64)}, 'ended')`;
  await sql`insert into public.agent_connect_teams_spend (agent_id, run_id, cents)
            values (${AGENT}, ${RUN}, 12)`;
}

const versions = migrationFiles().map((f) => f.id);

test(
  "an install of the earlier layout moves into the schema with every row as it was",
  { skip: !admin },
  async () => {
    await seedLegacy(legacy);
    const legacyCatalog = await catalog(legacy, "public");
    const before = await contents(legacy, legacyName);
    for (const table of TABLES) assert.ok(before[table].n > 0, `${table} is seeded`);

    // As Supabase's default privileges do, which the views must not inherit.
    await legacy`alter default privileges in schema public grant select on tables to public, anon`;
    const ran = await migrate(legacy, SCHEMA, quiet);
    assert.match(ran[0], /^adopted /);
    assert.deepEqual(ran.slice(1), [
      ...versions.slice(1).map((v) => `applied ${v}`),
      "made views public.agent_connect_* for the previous release",
    ]);
    assert.deepEqual(await contents(legacy, movedName), before);
    // Triggers saved before runners existed are Claude Managed Agents ones.
    assert.deepEqual(
      [...(await legacy`select runner, settings, secret from agent_connect.teams_triggers`)],
      [{ runner: "claude-managed-agents", settings: {}, secret: null }],
    );
    // Only views are left in public under the old names.
    assert.deepEqual(
      [
        ...(await legacy`select distinct c.relkind from pg_class c
                          where c.relnamespace = 'public'::regnamespace
                            and c.relname like 'agent\\_connect\\_%'`),
      ],
      [{ relkind: "v" }],
    );

    // Organizations: every id already holding rows, and nothing points at public any more.
    const orgs = await legacy`select id::text from agent_connect.organizations order by id`;
    assert.deepEqual(
      orgs.map((o) => o.id),
      [ORG_A, ORG_B],
    );
    const fks = await legacy`
    select confrelid::regclass::text as target, count(*)::int as n from pg_constraint
     where contype = 'f' and confrelid in ('public.organizations'::regclass,
                                           'agent_connect.organizations'::regclass)
     group by 1`;
    assert.deepEqual([...fks], [{ target: "agent_connect.organizations", n: 7 }]);

    // The same key opens what was sealed before, the agent's key still finds it, and the audit
    // log keeps counting.
    const store = createStore(legacy, SCHEMA);
    const p = (await store.getProvider(ORG_A, "microsoft"))!;
    assert.deepEqual(
      key.open(
        { ciphertext: p.credential, keyId: p.credentialKeyId },
        credentialAad(ORG_A, p.credentialScope),
      ),
      { client_secret: "provider" },
    );
    const c = (await store.getConnection(ORG_A, "jira"))!;
    assert.deepEqual(
      key.open(
        { ciphertext: c.credential!, keyId: c.credentialKeyId! },
        credentialAad(ORG_A, "jira"),
      ),
      { client_secret: "jira" },
    );
    assert.equal((await store.agentByKeyHash(hashAgentKey(AGENT_KEY)))?.id, AGENT);
    await store.audit({
      organizationId: ORG_A,
      agentId: AGENT,
      agentName: "Triage",
      connectorId: "jira",
      tool: "t",
      target: null,
      decision: "allowed",
      reason: null,
      vendorStatus: null,
    });
    const [{ max }] = await legacy`select max(id)::int as max from agent_connect.audit`;
    assert.equal(max, 101);

    // The previous release, rolled back to, still reads and writes through its old names,
    // and nobody else can.
    assert.deepEqual(
      [
        ...(await legacy`
      select count(*)::int as n from pg_class c, aclexplode(c.relacl) a
       where c.relnamespace = 'public'::regnamespace and c.relkind = 'v'
         and a.grantee <> c.relowner`),
      ],
      [{ n: 0 }],
    );
    assert.deepEqual(await contents(legacy, legacyName), await contents(legacy, movedName));
    await legacy`insert into public.agent_connect_audit
                   (organization_id, agent_id, agent_name, connector_id, tool, decision)
                 values (${ORG_A}, ${AGENT}, 'Triage', 'jira', 't', 'allowed')`;
    await legacy`insert into public.agent_connect_confirmed (organization_id, step, confirmed_by)
                 values (${ORG_A}, 'x', 'admin') on conflict (organization_id, step) do nothing`;
    await legacy`update public.agent_connect_agents set name = name where id = ${AGENT}`;
    await legacy`select id from public.agent_connect_teams_events for update skip locked`;
    assert.deepEqual(
      [...(await legacy`select max(id)::int as max from agent_connect.audit`)],
      [{ max: 102 }],
    );
    await legacy`delete from agent_connect.audit where id = 102`;
    await legacy`delete from agent_connect.confirmed where step = 'x'`;

    assert.deepEqual(await migrate(legacy, SCHEMA, quiet), [], "a second run does nothing");

    // A fresh install ends up with the same schema.
    assert.deepEqual(
      await migrate(fresh, SCHEMA, quiet),
      versions.map((v) => `applied ${v}`),
    );
    assert.deepEqual(await catalog(legacy, SCHEMA), await catalog(fresh, SCHEMA));

    // Rolling back restores the earlier layout and every row.
    const moved = await contents(legacy, movedName);
    await legacy.begin((tx) => tx.unsafe(ROLLBACK_SQL).simple());
    assert.deepEqual(await catalog(legacy, "public"), legacyCatalog);
    assert.deepEqual(await contents(legacy, legacyName), moved);
    assert.deepEqual(
      [...(await legacy`select to_regnamespace('agent_connect') as gone`)],
      [{ gone: null }],
    );

    // And upgrading again adopts it again.
    assert.match((await migrate(legacy, SCHEMA, quiet))[0], /^adopted /);
  },
);

test(
  "an owner that cannot bypass RLS upgrades with every organization kept",
  { skip: !admin },
  async () => {
    const sql = await database("ac_upgrade_owner");
    await seedLegacy(sql);
    const before = await contents(sql, legacyName);
    const role = "ac_upgrade_owner";
    await admin!.unsafe(`drop role if exists ${role}`);
    await admin!.unsafe(`create role ${role} login nosuperuser nobypassrls`);
    await sql.unsafe(`grant create on database ac_upgrade_owner to ${role}`);
    // It made those tables in public, so it may make the views there too.
    await sql.unsafe(`grant create on schema public to ${role}`);
    for (const table of TABLES)
      await sql.unsafe(`alter table public.agent_connect_${table} owner to ${role}`);
    const at = new URL(url!);
    at.pathname = "/ac_upgrade_owner";
    at.username = role;
    const owner = connectDatabase(at.toString());
    try {
      assert.match((await migrate(owner, SCHEMA, quiet))[0], /^adopted /);
      const orgs = await owner`select id::text from agent_connect.organizations order by id`;
      assert.deepEqual(
        orgs.map((o) => o.id),
        [ORG_A, ORG_B],
      );
      assert.deepEqual(await contents(sql, movedName), before);
    } finally {
      await owner.end();
      await admin!.unsafe("drop database if exists ac_upgrade_owner with (force)");
      await admin!.unsafe(`drop role if exists ${role}`);
    }
  },
);

test(
  "legacy tables made after the schema are dropped when empty, refused when not",
  { skip: !admin },
  async () => {
    const sql = await database("ac_upgrade_leftover");
    await migrate(sql, SCHEMA, quiet);
    await sql`create table public.organizations (id uuid primary key)`;
    await sql.unsafe(LEGACY_SQL).simple();
    assert.match(
      (await migrate(sql, SCHEMA, quiet))[0],
      /^dropped empty public\.agent_connect_agents/,
    );
    assert.deepEqual(
      [...(await sql`select to_regclass('public.agent_connect_agents') as t`)],
      [{ t: null }],
    );

    await sql.unsafe(LEGACY_SQL).simple();
    await sql`insert into public.organizations (id) values (${ORG_A})`;
    await sql`insert into public.agent_connect_confirmed (organization_id, step, confirmed_by)
            values (${ORG_A}, 'x', 'admin')`;
    await assert.rejects(migrate(sql, SCHEMA, quiet), /public\.agent_connect_confirmed hold rows/);

    // An owner that cannot bypass RLS sees no rows in the FORCE RLS tables, yet still refuses.
    const role = "ac_upgrade_leftover_owner";
    await admin!.unsafe(`drop role if exists ${role}`);
    await admin!.unsafe(`create role ${role} login nosuperuser nobypassrls`);
    await sql.unsafe(`grant create on database ac_upgrade_leftover to ${role}`);
    await sql.unsafe(`alter schema ${SCHEMA} owner to ${role}`);
    await sql.unsafe(`alter table ${SCHEMA}.schema_migrations owner to ${role}`);
    for (const table of TABLES)
      await sql.unsafe(`alter table public.agent_connect_${table} owner to ${role}`);
    const at = new URL(url!);
    at.pathname = "/ac_upgrade_leftover";
    at.username = role;
    const owner = connectDatabase(at.toString());
    try {
      await assert.rejects(
        migrate(owner, SCHEMA, quiet),
        /public\.agent_connect_confirmed hold rows/,
      );
      assert.equal(
        (await sql`select count(*)::int as n from public.agent_connect_confirmed`)[0].n,
        1,
      );
    } finally {
      await owner.end();
      await admin!.unsafe("drop database if exists ac_upgrade_leftover with (force)");
      await admin!.unsafe(`drop role if exists ${role}`);
    }
  },
);

test(
  "a changed migration is refused, and two runs at once apply each migration once",
  { skip: !admin },
  async () => {
    const sql = await database("ac_upgrade_race");
    const other = connect("ac_upgrade_race");
    const runs = await Promise.all([migrate(sql, SCHEMA, quiet), migrate(other, SCHEMA, quiet)]);
    assert.deepEqual(runs.map((r) => r.length).sort(), [0, versions.length]);

    await sql`update agent_connect.schema_migrations set checksum = 'x' where version = 1`;
    await assert.rejects(migrate(sql, SCHEMA, quiet), /changed after it was applied/);
    await admin!.unsafe("drop database if exists ac_upgrade_race with (force)");
  },
);
