import { randomUUID } from "node:crypto";
import postgres from "postgres";
import { brand } from "../../connector-kit/src/branding.js";
import { ConnectError, type Limits } from "../../connector-kit/src/types.js";

export type Sql = ReturnType<typeof postgres>;

export function connectDatabase(url: string): Sql {
  // Same convention as the admin app and edge functions: TLS unless explicitly disabled.
  const ssl =
    process.env.DATABASE_SSL === "disable" || url.includes("sslmode=disable") ? false : "require";
  return postgres(url, { prepare: false, idle_timeout: 20, max: 5, ssl });
}

// Unquoted in SQL, so only a plain lower-case identifier. Never public or a system schema:
// migrating revokes the schema from PUBLIC.
export function assertSchemaName(schema: string): void {
  if (!/^[a-z_][a-z0-9_]{0,62}$/.test(schema))
    throw new Error(
      `DATABASE_SCHEMA must be a lower-case identifier, not ${JSON.stringify(schema)}`,
    );
  if (schema === "public" || schema.startsWith("pg_"))
    throw new Error(`DATABASE_SCHEMA must be the service's own schema, not ${schema}`);
}

const TABLES = [
  "organizations",
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
  "approvals",
  "teams_file_offers",
  "event_subscriptions",
  "app_events",
  "event_triggers",
  "event_runs",
  "event_watches",
  "internal_apps",
  "leases",
] as const;

export type Tables = Record<(typeof TABLES)[number], postgres.Helper<string, []>>;

// Every table named in the configured schema, never through search_path (a pooler may drop
// it). The store writes queries both as tagged templates and as strings for sql.unsafe, so
// each name is an escaped identifier in the first and prints as schema.table in the second.
export function tables(sql: Sql, schema: string): Tables {
  assertSchemaName(schema);
  return Object.fromEntries(
    TABLES.map((table) => {
      const name = `${schema}.${table}`;
      return [table, Object.assign(sql(name), { toString: () => name })];
    }),
  ) as Tables;
}

/** `work` in a transaction holding the advisory lock on `key`, so a count and the insert it
 *  allows are one step across every process. For short work only: it holds a connection. */
export const withAdvisoryLock = <T>(sql: Sql, key: string, work: (tx: Sql) => Promise<T>) =>
  sql.begin(async (tx) => {
    await tx`select pg_advisory_xact_lock(hashtext(${key}))`;
    return work(tx as unknown as Sql);
  }) as Promise<T>;

/** Work every process would otherwise do at once, for its holder alone. No connection is held
 *  meanwhile, so it is the guard for work that waits on a vendor. */
export function leases(sql: Sql, t: Tables) {
  return {
    /** A token for `name` for `seconds`; null while another holds it. Taken with the token it
     *  gave, it renews. */
    async takeLease(name: string, seconds: number, token: string = randomUUID()) {
      const rows = await sql`
        insert into ${t.leases} as l (name, token, expires_at)
        values (${name}, ${token}, now() + make_interval(secs => ${seconds}))
        on conflict (name) do update set token = excluded.token, expires_at = excluded.expires_at
         where l.expires_at <= now() or l.token = excluded.token
        returning name`;
      return rows.length ? token : null;
    },

    async dropLease(name: string, token: string): Promise<void> {
      await sql`delete from ${t.leases} where name = ${name} and token = ${token}`;
    },

    /** `work` while holding `name`, renewed every third of `seconds` and dropped after; null,
     *  and no work, while another holds it. `signal` aborts once a renewal finds it taken. */
    async withLease<T>(
      name: string,
      seconds: number,
      work: (signal: AbortSignal) => Promise<T>,
    ): Promise<{ value: T } | null> {
      const token = await this.takeLease(name, seconds);
      if (!token) return null;
      const lost = new AbortController();
      // One renewal at a time, the last one landed before the drop: landing after it, it
      // would take the lease again.
      let renewal = Promise.resolve();
      const renewing = setInterval(() => {
        renewal = renewal
          .then(() => this.takeLease(name, seconds, token))
          .then(
            (held) => {
              if (!held) lost.abort(new Error(`Lost the lease ${name} to another process`));
            },
            (error) => console.error(`[leases] renewing ${name}`, error),
          );
      }, (seconds * 1000) / 3);
      renewing.unref();
      try {
        return { value: await work(lost.signal) };
      } finally {
        clearInterval(renewing);
        await renewal;
        await this.dropLease(name, token);
      }
    },
  };
}

export type ConnectionRow = {
  id: string;
  organizationId: string;
  connectorId: string;
  // The most any agent may do here (capability ids).
  capabilities: string[];
  config: Record<string, string>;
  // Null for an app signing in through a provider connection, which holds the credential.
  credential: Buffer | null;
  credentialKeyId: string | null;
  providerId: string | null;
  status: "connected" | "needs_reconnect";
  accountLabel: string | null;
  connectedBy: string | null;
  connectedAt: string;
  checkedAt: string;
};

export type PublicConnection = Omit<ConnectionRow, "credential" | "credentialKeyId">;

// One sign-in shared by a provider's apps, e.g. the customer's Microsoft Entra app.
// "pending": a certificate was made but nothing is connected yet.
export type ProviderRow = {
  id: string;
  organizationId: string;
  provider: string;
  status: "pending" | "connected" | "needs_reconnect";
  config: Record<string, string>;
  credential: Buffer;
  credentialKeyId: string;
  // What the credential is bound to (its AAD); rows moved from an app connection keep that
  // app's until re-sealed.
  credentialScope: string;
  accountLabel: string | null;
  connectedBy: string | null;
  connectedAt: string;
  checkedAt: string;
};

export type Organization = { id: string; name: string | null; createdAt: string };

export type AgentRow = {
  id: string;
  organizationId: string;
  name: string;
  status: "active" | "disabled";
  createdBy: string | null;
  createdAt: string;
};

export type { Limits };

export type GrantRow = {
  agentId: string;
  connectorId: string;
  capabilities: string[];
  limits: Limits;
  updatedAt: string;
};

// An organization's own app: its connector.json as the admin gave it.
export type InternalAppRow = {
  id: string;
  organizationId: string;
  definition: Record<string, unknown>;
  createdBy: string;
  createdAt: string;
  updatedBy: string;
  updatedAt: string;
};

// What replacing an internal app's definition does to its connection: nothing (not
// connected), keep it cut to what the new definition still has, or remove it with its grants.
export type InternalAppChange =
  | { kind: "none" }
  | { kind: "disconnect"; reason: string }
  | {
      kind: "keep";
      ceiling: string[];
      grants: { agentId: string; capabilities: string[]; limits: Limits }[];
      removed: string[];
    };

// Another admin action changed the provider connection between reading and saving it.
export class ProviderChanged extends ConnectError {
  constructor() {
    super("This connection changed meanwhile. Reload the page and try again.");
  }
}

export type AuditEntry = {
  organizationId: string;
  agentId: string | null;
  agentName: string | null;
  connectorId: string;
  tool: string;
  target: string | null;
  decision: "allowed" | "denied" | "failed";
  reason: string | null;
  vendorStatus: number | null;
};

// ISO 8601 in UTC, so every client parses timestamps the same way.
const iso = (column: string) =>
  `to_char(${column} at time zone 'UTC', 'YYYY-MM-DD"T"HH24:MI:SS.MS"Z"')`;

const CONNECTION_COLUMNS = `id::text, organization_id::text as "organizationId",
  connector_id as "connectorId", capabilities, config,
  credential, credential_key_id as "credentialKeyId", provider_id::text as "providerId", status,
  account_label as "accountLabel", connected_by as "connectedBy",
  ${iso("connected_at")} as "connectedAt", ${iso("checked_at")} as "checkedAt"`;

const PROVIDER_COLUMNS = `id::text, organization_id::text as "organizationId", provider, status,
  config, credential, credential_key_id as "credentialKeyId",
  credential_scope as "credentialScope", account_label as "accountLabel",
  connected_by as "connectedBy", ${iso("connected_at")} as "connectedAt",
  ${iso("checked_at")} as "checkedAt"`;

const ORGANIZATION_COLUMNS = `id::text, name, ${iso("created_at")} as "createdAt"`;

const AGENT_COLUMNS = `id::text, organization_id::text as "organizationId", name, status,
  created_by as "createdBy", ${iso("created_at")} as "createdAt"`;

const GRANT_COLUMNS = `g.agent_id::text as "agentId", g.connector_id as "connectorId",
  g.capabilities, g.limits, ${iso("g.updated_at")} as "updatedAt"`;

const INTERNAL_APP_COLUMNS = `id, organization_id::text as "organizationId", definition,
  created_by as "createdBy", ${iso("created_at")} as "createdAt", updated_by as "updatedBy",
  ${iso("updated_at")} as "updatedAt"`;

const ONE_ORG_PER_TENANT = "agent_connect_providers_one_org_per_tenant";

async function clearGrants(
  t: Tables,
  tx: postgres.TransactionSql,
  org: string,
  connectorId: string,
) {
  await tx`
    delete from ${t.grants} g using ${t.agents} a
     where g.agent_id = a.id and a.organization_id = ${org} and g.connector_id = ${connectorId}`;
}

// Lowering the connection's ceiling takes the dropped capabilities off every agent too, so
// raising it again grants nothing until each agent is ticked again.
async function trimGrants(
  t: Tables,
  tx: postgres.TransactionSql,
  org: string,
  connectorId: string,
  ceiling: string[],
) {
  await tx`
    update ${t.grants} g
       set capabilities = array(
             select c from unnest(g.capabilities) with ordinality u(c, o)
              where c = any(${tx.array(ceiling)}) order by o)
      from ${t.agents} a
     where a.id = g.agent_id and a.organization_id = ${org} and g.connector_id = ${connectorId}
       and not g.capabilities <@ ${tx.array(ceiling)}`;
}

async function removeProviderApps(
  t: Tables,
  tx: postgres.TransactionSql,
  org: string,
  provider: string,
) {
  await tx`
    delete from ${t.grants} g
     using ${t.agents} a, ${t.connections} c,
           ${t.providers} p
     where g.agent_id = a.id and a.organization_id = ${org}
       and c.organization_id = ${org} and c.connector_id = g.connector_id
       and c.provider_id = p.id and p.provider = ${provider}`;
  await tx`
    delete from ${t.connections} c using ${t.providers} p
     where c.provider_id = p.id and p.organization_id = ${org} and p.provider = ${provider}`;
}

async function insertAudit(t: Tables, q: Sql | postgres.TransactionSql, entry: AuditEntry) {
  await q`
    with logged as (
      insert into ${t.audit}
        (organization_id, agent_id, agent_name, connector_id, tool, target, decision, reason,
         vendor_status)
      values (${entry.organizationId}, ${entry.agentId}, ${entry.agentName},
              ${entry.connectorId}, ${entry.tool},
              ${entry.target?.slice(0, 2000) ?? null}, ${entry.decision}, ${entry.reason},
              ${entry.vendorStatus})
      returning agent_id, at)
    update ${t.agents} a set last_used_at = logged.at
      from logged where a.id = logged.agent_id`;
}

export function createStore(sql: Sql, schema: string) {
  const t = tables(sql, schema);
  return {
    async listOrganizations(): Promise<Organization[]> {
      return sql<Organization[]>`
        select ${sql.unsafe(ORGANIZATION_COLUMNS)} from ${t.organizations} order by created_at, id`;
    },

    async getOrganization(org: string): Promise<Organization | null> {
      const [row] = await sql<Organization[]>`
        select ${sql.unsafe(ORGANIZATION_COLUMNS)} from ${t.organizations} where id = ${org}`;
      return row ?? null;
    },

    async createOrganization(name: string): Promise<Organization> {
      const [row] = await sql<Organization[]>`
        insert into ${t.organizations} (name) values (${name})
        returning ${sql.unsafe(ORGANIZATION_COLUMNS)}`;
      return row;
    },

    /** Makes the organization with this id if missing; a name given replaces the kept one. */
    async ensureOrganization(org: string, name?: string): Promise<Organization> {
      // Without a name, an existing row is read, not rewritten (the legacy prefix calls this on
      // every request).
      if (name === undefined) {
        await sql`insert into ${t.organizations} (id) values (${org}) on conflict (id) do nothing`;
        const [row] = await sql<Organization[]>`
          select ${sql.unsafe(ORGANIZATION_COLUMNS)} from ${t.organizations} where id = ${org}`;
        return row;
      }
      const [row] = await sql<Organization[]>`
        insert into ${t.organizations} (id, name) values (${org}, ${name})
        on conflict (id) do update set name = excluded.name
        returning ${sql.unsafe(ORGANIZATION_COLUMNS)}`;
      return row;
    },

    async listConnections(org: string): Promise<ConnectionRow[]> {
      return sql.unsafe(
        `select ${CONNECTION_COLUMNS} from ${t.connections}
          where organization_id = $1 order by connector_id`,
        [org],
      );
    },

    async getConnection(org: string, connectorId: string): Promise<ConnectionRow | null> {
      const [row] = await sql.unsafe<ConnectionRow[]>(
        `select ${CONNECTION_COLUMNS} from ${t.connections}
          where organization_id = $1 and connector_id = $2`,
        [org, connectorId],
      );
      return row ?? null;
    },

    async saveConnection(input: {
      organizationId: string;
      connectorId: string;
      capabilities: string[];
      config: Record<string, string>;
      // Either the app's own sealed credential or the provider connection it signs in through.
      credential: Buffer | null;
      credentialKeyId: string | null;
      providerId: string | null;
      accountLabel: string;
      connectedBy: string;
      // The connection now reaches a different account (site, org, tenant), where the
      // agents' value limits would mean something else.
      clearGrants: boolean;
    }): Promise<void> {
      await sql.begin(async (tx) => {
        if (input.clearGrants) await clearGrants(t, tx, input.organizationId, input.connectorId);
        else await trimGrants(t, tx, input.organizationId, input.connectorId, input.capabilities);
        await tx`
        insert into ${t.connections}
          (organization_id, connector_id, capabilities, config, credential, credential_key_id,
           provider_id, status, account_label, connected_by)
        values (${input.organizationId}, ${input.connectorId}, ${sql.array(input.capabilities)},
                ${sql.json(input.config)}, ${input.credential}, ${input.credentialKeyId},
                ${input.providerId}, 'connected', ${input.accountLabel}, ${input.connectedBy})
        on conflict (organization_id, connector_id) do update set
          capabilities = excluded.capabilities, config = excluded.config,
          credential = excluded.credential, credential_key_id = excluded.credential_key_id,
          provider_id = excluded.provider_id, status = 'connected',
          account_label = excluded.account_label,
          connected_by = excluded.connected_by, connected_at = now(), checked_at = now()`;
      });
    },

    async listProviders(org: string): Promise<ProviderRow[]> {
      return sql.unsafe(
        `select ${PROVIDER_COLUMNS} from ${t.providers}
          where organization_id = $1 order by provider`,
        [org],
      );
    },

    async getProvider(org: string, provider: string): Promise<ProviderRow | null> {
      const [row] = await sql.unsafe<ProviderRow[]>(
        `select ${PROVIDER_COLUMNS} from ${t.providers}
          where organization_id = $1 and provider = $2`,
        [org, provider],
      );
      return row ?? null;
    },

    // A credential kept before the provider is connected, e.g. a certificate's key. Never
    // replaces a connected row's.
    async savePendingProvider(input: {
      organizationId: string;
      provider: string;
      config: Record<string, string>;
      credential: Buffer;
      credentialKeyId: string;
      credentialScope: string;
    }): Promise<boolean> {
      const rows = await sql`
        insert into ${t.providers} as p
          (organization_id, provider, status, config, credential, credential_key_id,
           credential_scope)
        values (${input.organizationId}, ${input.provider}, 'pending', ${sql.json(input.config)},
                ${input.credential}, ${input.credentialKeyId}, ${input.credentialScope})
        on conflict (organization_id, provider) do update set
          config = excluded.config, credential = excluded.credential,
          credential_key_id = excluded.credential_key_id,
          credential_scope = excluded.credential_scope
        where p.status = 'pending'
        returning id`;
      return rows.length > 0;
    },

    async saveProvider(input: {
      organizationId: string;
      provider: string;
      config: Record<string, string>;
      credential: Buffer;
      credentialKeyId: string;
      credentialScope: string;
      accountLabel: string;
      connectedBy: string;
      // The provider now reaches another account: its apps' settings and agents' grants
      // there meant something in the old one, so the apps are disconnected.
      removeApps: boolean;
      // The credential the caller read (null: no row). Saving fails if it changed meanwhile,
      // e.g. a certificate made for rotating, so that is not silently dropped.
      replaces: Buffer | null;
    }): Promise<void> {
      try {
        await sql.begin(async (tx) => {
          if (input.removeApps)
            await removeProviderApps(t, tx, input.organizationId, input.provider);
          const saved = await tx`
          insert into ${t.providers} as p
            (organization_id, provider, status, config, credential, credential_key_id,
             credential_scope, account_label, connected_by)
          values (${input.organizationId}, ${input.provider}, 'connected', ${sql.json(input.config)},
                  ${input.credential}, ${input.credentialKeyId}, ${input.credentialScope},
                  ${input.accountLabel}, ${input.connectedBy})
          on conflict (organization_id, provider) do update set
            status = 'connected', config = excluded.config, credential = excluded.credential,
            credential_key_id = excluded.credential_key_id,
            credential_scope = excluded.credential_scope, account_label = excluded.account_label,
            connected_by = excluded.connected_by, connected_at = now(), checked_at = now()
          where p.credential = ${input.replaces}
          returning id`;
          if (!saved.length) throw new ProviderChanged();
        });
      } catch (error) {
        if ((error as { constraint_name?: string }).constraint_name === ONE_ORG_PER_TENANT)
          throw new ConnectError(
            `This organization is already connected to another ${brand().productName} account. Disconnect it there first.`,
          );
        throw error;
      }
    },

    // A new credential for the same account, e.g. a certificate made or switched to. Only
    // over the credential the caller read (`replaces`): false if another write came first.
    async updateProviderCredential(
      id: string,
      config: Record<string, string>,
      credential: Buffer,
      credentialKeyId: string,
      credentialScope: string,
      replaces: Buffer,
      // The caller just signed in with this credential: the provider is connected.
      verified = false,
    ): Promise<boolean> {
      const rows = await sql`
        update ${t.providers}
           set config = ${sql.json(config)}, credential = ${credential},
               credential_key_id = ${credentialKeyId}, credential_scope = ${credentialScope}
               ${verified ? sql`, status = 'connected', checked_at = now()` : sql``}
         where id = ${id} and credential = ${replaces} returning id`;
      return rows.length > 0;
    },

    // Only for the credential that was checked: a reconnect meanwhile keeps its own status.
    async markProviderChecked(id: string, ok: boolean, credential: Buffer): Promise<void> {
      await sql`
        update ${t.providers}
           set checked_at = now(), status = ${ok ? "connected" : "needs_reconnect"}
         where id = ${id} and status <> 'pending' and credential = ${credential}`;
    },

    // Disconnects the provider and every app signing in through it.
    async removeProvider(org: string, provider: string): Promise<boolean> {
      return sql.begin(async (tx) => {
        await removeProviderApps(t, tx, org, provider);
        const rows = await tx`
          delete from ${t.providers}
           where organization_id = ${org} and provider = ${provider} returning id`;
        return rows.length > 0;
      });
    },

    // Only while the row still holds `replaces` (what was read): false when a reconnect saved
    // another credential meanwhile, which then stays.
    async reseal(
      id: string,
      replaces: Buffer,
      credential: Buffer,
      credentialKeyId: string,
    ): Promise<boolean> {
      const rows = await sql`
        update ${t.connections}
           set credential = ${credential}, credential_key_id = ${credentialKeyId}
         where id = ${id} and credential = ${replaces} returning id`;
      return rows.length > 0;
    },

    // Settings changed after connecting, e.g. the sites added one at a time. Only while the
    // row still holds `replaces` (the config read): false when it changed meanwhile.
    async updateConfig(
      org: string,
      connectorId: string,
      config: Record<string, string>,
      replaces: Record<string, string>,
    ): Promise<boolean> {
      const rows = await sql`
        update ${t.connections} set config = ${sql.json(config)}
         where organization_id = ${org} and connector_id = ${connectorId}
           and config = ${sql.json(replaces)} returning id`;
      return rows.length > 0;
    },

    async markChecked(id: string, ok: boolean): Promise<void> {
      await sql`
        update ${t.connections}
           set checked_at = now(), status = ${ok ? "connected" : "needs_reconnect"}
         where id = ${id}`;
    },

    async setCeiling(org: string, connectorId: string, ceiling: string[]): Promise<boolean> {
      return sql.begin(async (tx) => {
        const rows = await tx`
          update ${t.connections} set capabilities = ${tx.array(ceiling)}
           where organization_id = ${org} and connector_id = ${connectorId} returning id`;
        if (rows.length) await trimGrants(t, tx, org, connectorId, ceiling);
        return rows.length > 0;
      });
    },

    async removeConnection(org: string, connectorId: string): Promise<boolean> {
      return sql.begin(async (tx) => {
        await clearGrants(t, tx, org, connectorId);
        const rows = await tx`
          delete from ${t.connections}
           where organization_id = ${org} and connector_id = ${connectorId} returning id`;
        return rows.length > 0;
      });
    },

    // lastUsedAt: the agent's latest request, stamped with each audit row; null if none.
    async listAgents(
      org: string,
    ): Promise<(AgentRow & { lastUsedAt: string | null; grants: GrantRow[] })[]> {
      const agents: (AgentRow & { lastUsedAt: string | null })[] = await sql.unsafe(
        `select ${AGENT_COLUMNS}, ${iso("last_used_at")} as "lastUsedAt"
           from ${t.agents}
          where organization_id = $1 order by created_at`,
        [org],
      );
      const grants: GrantRow[] = await sql.unsafe(
        `select ${GRANT_COLUMNS} from ${t.grants} g
          join ${t.agents} a on a.id = g.agent_id
         where a.organization_id = $1
         order by g.connector_id`,
        [org],
      );
      return agents.map((a) => ({ ...a, grants: grants.filter((g) => g.agentId === a.id) }));
    },

    async createAgent(
      input: {
        organizationId: string;
        name: string;
        keyHash: string;
        createdBy: string;
      },
      audit: (agent: AgentRow) => AuditEntry,
    ): Promise<AgentRow | null> {
      // With its audit row or not at all: a key nobody was shown must not exist.
      return sql.begin(async (tx) => {
        const [row] = await tx.unsafe<AgentRow[]>(
          `insert into ${t.agents} (organization_id, name, key_hash, status, created_by)
           values ($1, $2, $3, 'active', $4)
           on conflict (organization_id, name) do nothing
           returning ${AGENT_COLUMNS}`,
          [input.organizationId, input.name, input.keyHash, input.createdBy],
        );
        if (row) await insertAudit(t, tx, audit(row));
        return row ?? null;
      });
    },

    async updateAgent(
      org: string,
      agentId: string,
      patch: { status?: AgentRow["status"]; keyHash?: string },
      audit: AuditEntry,
    ): Promise<boolean> {
      // With its audit row or not at all, so a failed audit never leaves a rotated key unshown.
      return sql.begin(async (tx) => {
        const rows = await tx`
          update ${t.agents}
             set status = coalesce(${patch.status ?? null}, status),
                 key_hash = coalesce(${patch.keyHash ?? null}, key_hash)
           where id = ${agentId} and organization_id = ${org} returning id`;
        if (rows.length) await insertAudit(t, tx, audit);
        return rows.length > 0;
      });
    },

    async deleteAgent(org: string, agentId: string): Promise<boolean> {
      const rows = await sql`
        delete from ${t.agents}
         where id = ${agentId} and organization_id = ${org} returning id`;
      return rows.length > 0;
    },

    async agentByKeyHash(keyHash: string): Promise<AgentRow | null> {
      const [row] = await sql.unsafe<AgentRow[]>(
        `select ${AGENT_COLUMNS} from ${t.agents} where key_hash = $1`,
        [keyHash],
      );
      return row ?? null;
    },

    async grantsFor(agentId: string): Promise<GrantRow[]> {
      return sql.unsafe(`select ${GRANT_COLUMNS} from ${t.grants} g where g.agent_id = $1`, [
        agentId,
      ]);
    },

    async setGrant(
      org: string,
      agentId: string,
      connectorId: string,
      capabilities: string[],
      limits: Limits,
    ): Promise<boolean> {
      // Only within the app's ceiling as it is now: the connection row is locked, so a
      // ceiling lowered meanwhile either lands first (nothing saved) or strips this grant.
      const rows = await sql`
        insert into ${t.grants} (agent_id, connector_id, capabilities, limits)
        select a.id, ${connectorId}, ${sql.array(capabilities)}, ${sql.json(limits)}
          from ${t.agents} a
          join ${t.connections} c
            on c.organization_id = a.organization_id and c.connector_id = ${connectorId}
         where a.id = ${agentId} and a.organization_id = ${org}
           and c.capabilities @> ${sql.array(capabilities)}
           for share of c
        on conflict (agent_id, connector_id) do update
          set capabilities = excluded.capabilities, limits = excluded.limits, updated_at = now()
        returning agent_id`;
      return rows.length > 0;
    },

    async removeGrant(org: string, agentId: string, connectorId: string): Promise<boolean> {
      const rows = await sql`
        delete from ${t.grants} g using ${t.agents} a
         where g.agent_id = a.id and a.id = ${agentId} and a.organization_id = ${org}
           and g.connector_id = ${connectorId}
        returning g.agent_id`;
      return rows.length > 0;
    },

    async audit(entry: AuditEntry): Promise<void> {
      await insertAudit(t, sql, entry);
    },

    // Setup steps an admin marked done (the service cannot check them live).
    async listConfirmed(org: string): Promise<{ step: string; by: string; at: string }[]> {
      return sql.unsafe(
        `select step, confirmed_by as by, ${iso("confirmed_at")} as at
           from ${t.confirmed} where organization_id = $1 order by step`,
        [org],
      ) as unknown as Promise<{ step: string; by: string; at: string }[]>;
    },

    async setConfirmed(org: string, step: string, by: string, on: boolean): Promise<void> {
      if (on)
        await sql`
          insert into ${t.confirmed} (organization_id, step, confirmed_by)
          values (${org}, ${step}, ${by})
          on conflict (organization_id, step) do nothing`;
      else
        await sql`
          delete from ${t.confirmed}
           where organization_id = ${org} and step = ${step}`;
    },

    async listInternalApps(org: string): Promise<InternalAppRow[]> {
      return sql.unsafe(
        `select ${INTERNAL_APP_COLUMNS} from ${t.internal_apps}
          where organization_id = $1 order by id`,
        [org],
      );
    },

    async getInternalApp(org: string, id: string): Promise<InternalAppRow | null> {
      const [row] = await sql.unsafe<InternalAppRow[]>(
        `select ${INTERNAL_APP_COLUMNS} from ${t.internal_apps}
          where organization_id = $1 and id = $2`,
        [org, id],
      );
      return row ?? null;
    },

    /** Creates or replaces the definition, and applies what `plan` decides for its connection
     *  and grants, read and changed under locks in the same transaction, with its audit row. */
    async saveInternalApp(
      input: { organizationId: string; id: string; definition: unknown; by: string },
      plan: (
        previous: Record<string, unknown> | null,
        connection: ConnectionRow | null,
        grants: GrantRow[],
      ) => InternalAppChange,
      audit: (created: boolean, change: InternalAppChange) => AuditEntry,
    ): Promise<{ row: InternalAppRow; created: boolean; change: InternalAppChange }> {
      const org = input.organizationId;
      return sql.begin(async (tx) => {
        const [previous] = await tx`
          select definition from ${t.internal_apps}
           where organization_id = ${org} and id = ${input.id} for update`;
        // Locked, so a grant saved meanwhile (setGrant shares this lock) lands before or after.
        const [connection] = await tx.unsafe<ConnectionRow[]>(
          `select ${CONNECTION_COLUMNS} from ${t.connections}
            where organization_id = $1 and connector_id = $2 for update`,
          [org, input.id],
        );
        const grants: GrantRow[] = connection
          ? await tx.unsafe(
              `select ${GRANT_COLUMNS} from ${t.grants} g join ${t.agents} a on a.id = g.agent_id
                where a.organization_id = $1 and g.connector_id = $2`,
              [org, input.id],
            )
          : [];
        const change = plan(previous?.definition ?? null, connection ?? null, grants);
        if (change.kind === "disconnect") {
          await clearGrants(t, tx, org, input.id);
          await tx`
            delete from ${t.connections}
             where organization_id = ${org} and connector_id = ${input.id}`;
        } else if (change.kind === "keep") {
          await tx`
            update ${t.connections} set capabilities = ${tx.array(change.ceiling)}
             where organization_id = ${org} and connector_id = ${input.id}`;
          for (const g of change.grants)
            await tx`
              update ${t.grants}
                 set capabilities = ${tx.array(g.capabilities)}, limits = ${tx.json(g.limits)},
                     updated_at = now()
               where agent_id = ${g.agentId} and connector_id = ${input.id}`;
          if (change.removed.length)
            await tx`
              delete from ${t.grants}
               where agent_id = any(${tx.array(change.removed)}::uuid[])
                 and connector_id = ${input.id}`;
        }
        const [row] = await tx<InternalAppRow[]>`
          insert into ${t.internal_apps}
            (organization_id, id, definition, created_by, updated_by)
          values (${org}, ${input.id}, ${tx.json(input.definition as never)}, ${input.by},
                  ${input.by})
          on conflict (organization_id, id) do update set
            definition = excluded.definition, updated_by = excluded.updated_by, updated_at = now()
          returning ${tx.unsafe(INTERNAL_APP_COLUMNS)}`;
        const created = !previous;
        await insertAudit(t, tx, audit(created, change));
        return { row, created, change };
      });
    },

    /** Removes the definition, its connection and every grant of it, with its audit row. */
    async deleteInternalApp(org: string, id: string, audit: AuditEntry): Promise<boolean> {
      return sql.begin(async (tx) => {
        const rows = await tx`
          delete from ${t.internal_apps} where organization_id = ${org} and id = ${id}
          returning id`;
        if (!rows.length) return false;
        await clearGrants(t, tx, org, id);
        await tx`
          delete from ${t.connections} where organization_id = ${org} and connector_id = ${id}`;
        await insertAudit(t, tx, audit);
        return true;
      });
    },

    async recentAudit(
      org: string,
      limit: number,
      filter: { agentId: string | null; connectorId: string | null },
    ): Promise<(Omit<AuditEntry, "organizationId"> & { at: string })[]> {
      return sql.unsafe(
        `select ${iso("l.at")} as at, l.agent_id::text as "agentId",
                coalesce(a.name, l.agent_name) as "agentName",
                l.connector_id as "connectorId", l.tool, l.target, l.decision, l.reason,
                l.vendor_status as "vendorStatus"
           from ${t.audit} l
           left join ${t.agents} a on a.id = l.agent_id
          where l.organization_id = $1
            and ($2::uuid is null or l.agent_id = $2::uuid)
            and ($3::text is null or l.connector_id = $3)
          order by l.at desc, l.id desc limit $4`,
        [org, filter.agentId, filter.connectorId, limit],
      );
    },
  };
}

export type Store = ReturnType<typeof createStore>;
