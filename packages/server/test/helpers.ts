import { randomBytes, randomUUID } from "node:crypto";
import { createApp } from "../src/app.ts";
import { loadConnectors, loadProviders } from "../../connector-kit/src/connector.ts";
import { CredentialKey, credentialAad, hashAgentKey, newAgentKey } from "../src/secrets.ts";
import { createService, type App } from "../src/service.ts";
import { ProviderChanged } from "../src/store.ts";
import type {
  AgentRow,
  AuditEntry,
  ConnectionRow,
  GrantRow,
  ProviderRow,
  Store,
} from "../src/store.ts";
import { jiraTools } from "../../../connectors/jira/tools.ts";
import { microsoftTools } from "../../../connectors/outlook-mail/tools.ts";
import { graphTools } from "../../../connectors/sharepoint-files/tools.ts";
import { calendarTools } from "../../../connectors/outlook-calendar/tools.ts";
import { excelTools } from "../../../connectors/sharepoint-files/excel.ts";
import { toolsets } from "../../connector-kit/src/common.ts";
import { teamsTools } from "../../../connectors/teams/tools.ts";
import { normalizeSites } from "../../../connectors/sharepoint-files/vendor.ts";
import {
  ConnectError,
  type ProviderClient,
  type Vendor,
  type VendorConnection,
  type VendorRequest,
  type VendorResponse,
} from "../../connector-kit/src/types.ts";

export const ORG = "11111111-1111-1111-1111-111111111111";
export const AGENT = "22222222-2222-2222-2222-222222222222";
export const ADMIN_TOKEN = "t".repeat(40);

export const connectors = loadConnectors();
export const jira = connectors.find((c) => c.id === "jira")!;
export const microsoft = connectors.find((c) => c.id === "outlook-mail")!;
export const graph = connectors.find((c) => c.id === "sharepoint-files")!;
export const calendar = connectors.find((c) => c.id === "outlook-calendar")!;
export const teamsConnector = connectors.find((c) => c.id === "teams")!;
export const microsoftProvider = loadProviders().find((p) => p.id === "microsoft")!;
export const TENANT = "33333333-3333-3333-3333-333333333333";

// The Microsoft connection: the customer's app signs in with a secret, or with a
// certificate made before connecting. `refuse` makes Microsoft refuse a sign-in. With the
// shared app offered, a sign-in code proves the tenant it maps to in `signIns`.
export function fakeMicrosoftProvider({ approval = false } = {}): ProviderClient & {
  connected: string[];
  verified: VendorConnection[];
  signIns: Record<string, string>;
  refuse: (c: VendorConnection) => boolean;
} {
  const connected: string[] = [];
  const verified: VendorConnection[] = [];
  const signIns: Record<string, string> = {};
  const client = {
    connected,
    verified,
    signIns,
    refuse: (_c: VendorConnection) => false,
    async connect(fields: Record<string, string>, existing: VendorConnection | null) {
      const base = { tenantId: fields.tenant_id, clientId: fields.client_id };
      const conn = fields.client_secret
        ? {
            config: { auth: "client_secret", ...base },
            secret: { client_secret: fields.client_secret } as Record<string, string>,
          }
        : existing?.config.certificate
          ? {
              config: { auth: "certificate", ...base, certificate: existing.config.certificate },
              secret: { private_key: existing.secret.private_key },
            }
          : null;
      if (!conn) throw new ConnectError("Create the certificate or enter a client secret");
      await client.verify(conn);
      connected.push(fields.tenant_id);
      return {
        config: { ...conn.config, servicePrincipalId: "sp-1" },
        secret: conn.secret,
        accountLabel: fields.tenant_id,
      };
    },
    async verify(c: VendorConnection) {
      verified.push(c);
      if (client.refuse(c)) throw new ConnectError("Microsoft refused the app's sign-in");
    },
    ...(approval
      ? {
          approval: {
            clientId: "app-123",
            consentUrl: ({ state, redirectUri }: { state: string; redirectUri: string }) =>
              `https://login.example/consent?state=${state}&redirect_uri=${redirectUri}`,
            signInUrl: ({
              tenant,
              state,
              nonce,
            }: {
              tenant: string;
              state: string;
              nonce: string;
            }) =>
              `https://login.example/${tenant}/authorize?${new URLSearchParams({ state, nonce })}`,
            async proveTenant({ tenant, code }: { tenant: string; code: string }) {
              if (signIns[code] !== tenant)
                throw new ConnectError("Sign in with an account of that organization");
            },
            async connect(tenant: string) {
              connected.push(tenant);
              return {
                config: {
                  auth: "arbium_app",
                  tenantId: tenant,
                  clientId: "app-123",
                  servicePrincipalId: "sp-1",
                },
                secret: {},
                accountLabel: "X (x.com)",
              };
            },
          },
        }
      : {}),
  };
  return client;
}

// Microsoft's apps: one mailbox, agent@x.com; outside@x.com is one Exchange keeps the app out of.
export function fakeMicrosoft(): Vendor {
  return {
    async connect(_fields, provider) {
      if (!provider) throw new Error("no provider");
      return { config: {}, secret: {}, accountLabel: provider.accountLabel };
    },
    checks: { permissions: async () => ({ values: [] }) },
    async call(_conn, req): Promise<VendorResponse> {
      const path = decodeURIComponent(req.path.split("?")[0]);
      if (path === "/v1.0/organization")
        return { status: 200, data: { value: [{ displayName: "X" }] } };
      const box = /^\/v1\.0\/users\/([^/]+)\//.exec(path)?.[1];
      if (box === "outside@x.com")
        return { status: 403, data: { error: { code: "ErrorAccessDenied" } } };
      if (box === "agent@x.com") return { status: 200, data: { id: "x" } };
      return { status: 404, data: { error: { message: "Mailbox not found" } } };
    },
  };
}

export const FINANCE = "https://acme.sharepoint.com/sites/finance";

// SharePoint: Finance is the one site granted.
export function fakeGraph(): Vendor {
  return {
    async connect(fields, provider, previous) {
      if (!provider) throw new Error("no provider");
      return {
        config: {
          ...(fields.access ? { access: fields.access } : {}),
          sites: previous?.sites ?? FINANCE,
        },
        secret: {},
        accountLabel: "acme.sharepoint.com",
      };
    },
    sites: {
      normalize: normalizeSites,
      async access(_c, site) {
        const granted = site === FINANCE;
        return {
          site,
          status: granted ? "granted" : "not_granted",
          name: granted ? "Finance" : null,
          message: granted ? "Agents can open this site's files." : "Not given yet.",
        };
      },
    },
    checks: {
      permissions: async () => ({ values: ["Sites.Selected"] }),
      sites: async (c) => ({ values: c.config.sites.split("\n") }),
    },
    async call(): Promise<VendorResponse> {
      return { status: 404, data: { error: { message: "itemNotFound" } } };
    },
  };
}

// In-memory Jira with projects KAN and SECRET and one issue, KAN-1.
export function fakeJira(): Vendor & { requests: VendorRequest[] } {
  const requests: VendorRequest[] = [];
  return {
    requests,
    async connect(fields) {
      return {
        config: { site: fields.site, cloudId: fields.cloud ?? "c" },
        secret: { client_secret: fields.client_secret },
        accountLabel: fields.site,
      };
    },
    async call(_conn, req): Promise<VendorResponse> {
      requests.push(req);
      const path = req.path.split("?")[0];
      if (path === "/rest/api/3/serverInfo")
        return { status: 200, data: { baseUrl: "https://x.atlassian.net" } };
      if (path === "/rest/api/3/project/search")
        return {
          status: 200,
          data: {
            values: [
              { key: "KAN", name: "Kanban" },
              { key: "SECRET", name: "Secret" },
            ],
            isLast: true,
          },
        };
      if (path === "/rest/api/3/permissions/project")
        return {
          status: 200,
          data: { projects: ["KAN", "SECRET"].map((key) => ({ id: key, key })) },
        };
      if (path === "/rest/api/3/issue/KAN-1/comment" && req.method === "GET")
        return { status: 200, data: { comments: [], total: 0 } };
      if (path === "/rest/api/3/issue/KAN-1")
        return {
          status: 200,
          data: { key: "KAN-1", fields: { summary: "Allowed issue", project: { key: "KAN" } } },
        };
      return { status: 404, data: { errorMessages: ["Issue does not exist"] } };
    },
  };
}

export function fakeStore(key: CredentialKey) {
  const sealed = key.seal({ client_secret: "x" }, credentialAad(ORG, "jira"));
  const sealedMs = key.seal({ client_secret: "x" }, credentialAad(ORG, "provider:microsoft"));
  const state = {
    // ORG holds the seeded rows, so it is known, as an upgraded install's organizations are.
    organizations: new Map<string, { name: string | null; createdAt: string }>([
      [ORG, { name: null, createdAt: new Date(0).toISOString() }],
    ]),
    providers: [
      {
        id: "p1",
        organizationId: ORG,
        provider: "microsoft",
        status: "connected",
        config: {
          auth: "client_secret",
          tenantId: TENANT,
          clientId: TENANT,
          servicePrincipalId: "sp-1",
        },
        credential: sealedMs.ciphertext,
        credentialKeyId: sealedMs.keyId,
        credentialScope: "provider:microsoft",
        accountLabel: TENANT,
        connectedBy: "admin",
        connectedAt: "",
        checkedAt: "",
      },
    ] as ProviderRow[],
    connections: [
      {
        id: "c1",
        organizationId: ORG,
        connectorId: "jira",
        capabilities: jira.capabilities.map((c) => c.id),
        config: { cloudId: "c" },
        credential: sealed.ciphertext,
        credentialKeyId: sealed.keyId,
        providerId: null,
        status: "connected",
        accountLabel: "x.atlassian.net",
        connectedBy: "admin",
        connectedAt: "",
        checkedAt: "",
      } as ConnectionRow,
      {
        id: "c2",
        organizationId: ORG,
        connectorId: "outlook-mail",
        capabilities: microsoft.capabilities.map((c) => c.id),
        config: {},
        credential: null,
        credentialKeyId: null,
        providerId: "p1",
        status: "connected",
        accountLabel: "X (x.com)",
        connectedBy: "admin",
        connectedAt: "",
        checkedAt: "",
      } as ConnectionRow,
      {
        id: "c4",
        organizationId: ORG,
        connectorId: "sharepoint-files",
        capabilities: graph.capabilities.map((c) => c.id),
        config: { sites: FINANCE },
        credential: null,
        credentialKeyId: null,
        providerId: "p1",
        status: "connected",
        accountLabel: "acme.sharepoint.com",
        connectedBy: "admin",
        connectedAt: "",
        checkedAt: "",
      } as ConnectionRow,
      ...(
        [
          ["c5", "outlook-calendar", calendar, {}, "X (x.com)"],
          ["c8", "teams", teamsConnector, {}, TENANT],
        ] as const
      ).map(
        ([id, connectorId, connector, config, accountLabel]) =>
          ({
            id,
            organizationId: ORG,
            connectorId,
            capabilities: connector.capabilities.map((c) => c.id),
            config: { ...config },
            credential: null,
            credentialKeyId: null,
            providerId: "p1",
            status: "connected",
            accountLabel,
            connectedBy: "admin",
            connectedAt: "",
            checkedAt: "",
          }) as ConnectionRow,
      ),
    ],
    agents: [] as (AgentRow & { keyHash: string })[],
    grants: [] as GrantRow[],
    audit: [] as AuditEntry[],
    // Every audit write throws while set, as a database that refuses the insert would.
    auditDown: false,
    confirmed: [] as { org: string; step: string; by: string; at: string }[],
    checked: [] as { id: string; ok: boolean }[],
    saved: [] as {
      organizationId: string;
      connectorId: string;
      capabilities: string[];
      config: Record<string, string>;
      credential: Buffer | null;
      credentialKeyId: string | null;
      providerId: string | null;
      accountLabel: string;
      connectedBy: string;
      clearGrants: boolean;
    }[],
    savedProviders: [] as { removeApps: boolean; config: Record<string, string> }[],
  };
  const removeProviderApps = (org: string, provider: string) => {
    const p = state.providers.find((x) => x.organizationId === org && x.provider === provider);
    const gone = state.connections
      .filter((c) => p && c.providerId === p.id)
      .map((c) => c.connectorId);
    state.grants = state.grants.filter((g) => !gone.includes(g.connectorId));
    state.connections = state.connections.filter((c) => !gone.includes(c.connectorId));
  };
  const logged = (entry: AuditEntry) => {
    if (state.auditDown) throw new Error("audit insert refused");
    state.audit.push(entry);
  };
  const store = {
    listConnections: async (org: string) =>
      state.connections.filter((c) => c.organizationId === org),
    getConnection: async (org: string, id: string) =>
      state.connections.find((c) => c.organizationId === org && c.connectorId === id) ?? null,
    saveConnection: async (row: (typeof state.saved)[number]) => {
      state.saved.push(row);
      const { clearGrants: _c, ...rest } = row;
      const previous = state.connections.find(
        (c) => c.organizationId === row.organizationId && c.connectorId === row.connectorId,
      );
      const next = {
        id: previous?.id ?? randomUUID(),
        status: "connected" as const,
        connectedAt: "",
        checkedAt: "",
        ...rest,
      };
      if (previous) Object.assign(previous, next);
      else state.connections.push(next);
    },
    listProviders: async (org: string) => state.providers.filter((p) => p.organizationId === org),
    getProvider: async (org: string, provider: string) =>
      state.providers.find((p) => p.organizationId === org && p.provider === provider) ?? null,
    savePendingProvider: async (
      row: Omit<
        ProviderRow,
        "id" | "status" | "accountLabel" | "connectedBy" | "connectedAt" | "checkedAt"
      >,
    ) => {
      const previous = state.providers.find(
        (p) => p.organizationId === row.organizationId && p.provider === row.provider,
      );
      if (previous && previous.status !== "pending") return false;
      if (previous) Object.assign(previous, row);
      else
        state.providers.push({
          id: randomUUID(),
          status: "pending",
          accountLabel: null,
          connectedBy: null,
          connectedAt: "",
          checkedAt: "",
          ...row,
        });
      return true;
    },
    saveProvider: async (
      row: Omit<ProviderRow, "id" | "status" | "connectedAt" | "checkedAt"> & {
        removeApps: boolean;
        replaces: Buffer | null;
      },
    ) => {
      const { removeApps, replaces, ...rest } = row;
      const previous = state.providers.find(
        (p) => p.organizationId === row.organizationId && p.provider === row.provider,
      );
      if ((previous?.credential ?? null) !== replaces) throw new ProviderChanged();
      state.savedProviders.push({ removeApps, config: row.config });
      if (removeApps) removeProviderApps(row.organizationId, row.provider);
      const next = { status: "connected" as const, ...rest };
      if (previous) Object.assign(previous, next);
      else state.providers.push({ id: randomUUID(), connectedAt: "", checkedAt: "", ...next });
    },
    updateProviderCredential: async (
      id: string,
      config: Record<string, string>,
      credential: Buffer,
      credentialKeyId: string,
      credentialScope: string,
      replaces: Buffer,
      verified = false,
    ) => {
      const row = state.providers.find((p) => p.id === id)!;
      if (row.credential !== replaces) return false;
      Object.assign(row, {
        config,
        credential,
        credentialKeyId,
        credentialScope,
        ...(verified ? { status: "connected" as const } : {}),
      });
      return true;
    },
    markProviderChecked: async (id: string, ok: boolean, credential: Buffer) => {
      const row = state.providers.find((p) => p.id === id)!;
      if (row.status !== "pending" && row.credential === credential)
        row.status = ok ? "connected" : "needs_reconnect";
    },
    removeProvider: async (org: string, provider: string) => {
      removeProviderApps(org, provider);
      const before = state.providers.length;
      state.providers = state.providers.filter(
        (p) => !(p.organizationId === org && p.provider === provider),
      );
      return state.providers.length < before;
    },
    updateConfig: async (
      org: string,
      connectorId: string,
      config: Record<string, string>,
      replaces: Record<string, string>,
    ) => {
      const row = state.connections.find(
        (c) => c.organizationId === org && c.connectorId === connectorId,
      );
      if (!row || JSON.stringify(row.config) !== JSON.stringify(replaces)) return false;
      row.config = config;
      return true;
    },
    markChecked: async (id: string, ok: boolean) => {
      state.checked.push({ id, ok });
      const row = state.connections.find((c) => c.id === id);
      if (row) row.status = ok ? "connected" : "needs_reconnect";
    },
    reseal: async (id: string, replaces: Buffer, credential: Buffer, credentialKeyId: string) => {
      const row = state.connections.find((c) => c.id === id)!;
      if (!row.credential?.equals(replaces)) return false;
      Object.assign(row, { credential, credentialKeyId });
      return true;
    },
    updateAgent: async (
      org: string,
      id: string,
      patch: { status?: AgentRow["status"]; keyHash?: string },
      entry: AuditEntry,
    ) => {
      const a = state.agents.find((x) => x.id === id && x.organizationId === org);
      if (!a) return false;
      logged(entry);
      Object.assign(a, Object.fromEntries(Object.entries(patch).filter(([, v]) => v)));
      return true;
    },
    deleteAgent: async (org: string, id: string) => {
      const before = state.agents.length;
      state.agents = state.agents.filter((x) => !(x.id === id && x.organizationId === org));
      state.grants = state.grants.filter((g) => state.agents.some((a) => a.id === g.agentId));
      return state.agents.length < before;
    },
    removeGrant: async (_org: string, id: string, connectorId: string) => {
      const before = state.grants.length;
      state.grants = state.grants.filter(
        (g) => !(g.agentId === id && g.connectorId === connectorId),
      );
      return state.grants.length < before;
    },
    removeConnection: async (org: string, connectorId: string) => {
      const before = state.connections.length;
      state.connections = state.connections.filter(
        (c) => !(c.organizationId === org && c.connectorId === connectorId),
      );
      state.grants = state.grants.filter((g) => g.connectorId !== connectorId);
      return state.connections.length < before;
    },
    agentByKeyHash: async (h: string) => state.agents.find((a) => a.keyHash === h) ?? null,
    grantsFor: async (id: string) => state.grants.filter((g) => g.agentId === id),
    listAgents: async () =>
      state.agents.map(({ keyHash: _hash, ...a }) => ({
        ...a,
        lastUsedAt: null,
        grants: state.grants.filter((g) => g.agentId === a.id),
      })),
    createAgent: async (
      input: {
        organizationId: string;
        name: string;
        keyHash: string;
        createdBy: string;
      },
      entry: (agent: AgentRow) => AuditEntry,
    ) => {
      const agent = {
        id: state.agents.length ? randomUUID() : AGENT,
        organizationId: input.organizationId,
        name: input.name,
        status: "active" as const,
        createdBy: input.createdBy,
        createdAt: "",
        keyHash: input.keyHash,
      };
      const { keyHash: _hash, ...row } = agent;
      logged(entry(row));
      state.agents.push(agent);
      return row;
    },
    setGrant: async (
      org: string,
      agentId: string,
      connectorId: string,
      capabilities: string[],
      limits: GrantRow["limits"],
    ) => {
      const conn = state.connections.find(
        (c) => c.organizationId === org && c.connectorId === connectorId,
      );
      if (!conn || capabilities.some((c) => !conn.capabilities.includes(c))) return false;
      state.grants = state.grants.filter(
        (g) => !(g.agentId === agentId && g.connectorId === connectorId),
      );
      state.grants.push({ agentId, connectorId, capabilities, limits, updatedAt: "" });
      return true;
    },
    audit: async (entry: AuditEntry) => logged(entry),
    setCeiling: async (org: string, connectorId: string, ceiling: string[]) => {
      const row = state.connections.find(
        (c) => c.organizationId === org && c.connectorId === connectorId,
      );
      if (row) {
        row.capabilities = ceiling;
        for (const g of state.grants)
          if (g.connectorId === connectorId)
            g.capabilities = g.capabilities.filter((c) => ceiling.includes(c));
      }
      return Boolean(row);
    },
    listConfirmed: async (org: string) =>
      state.confirmed
        .filter((c) => c.org === org)
        .map(({ step, by, at }) => ({ step, by, at }))
        .sort((a, b) => a.step.localeCompare(b.step)),
    setConfirmed: async (org: string, step: string, by: string, on: boolean) => {
      state.confirmed = state.confirmed.filter((c) => !(c.org === org && c.step === step));
      if (on) state.confirmed.push({ org, step, by, at: new Date(0).toISOString() });
    },
    listOrganizations: async () => [...state.organizations].map(([id, o]) => ({ id, ...o })),
    getOrganization: async (org: string) => {
      const o = state.organizations.get(org);
      return o ? { id: org, ...o } : null;
    },
    createOrganization: async (name: string) => {
      const id = randomUUID();
      state.organizations.set(id, { name, createdAt: new Date().toISOString() });
      return { id, ...state.organizations.get(id)! };
    },
    ensureOrganization: async (org: string, name?: string) => {
      const kept = state.organizations.get(org);
      state.organizations.set(org, {
        name: name ?? kept?.name ?? null,
        createdAt: kept?.createdAt ?? new Date().toISOString(),
      });
      return { id: org, ...state.organizations.get(org)! };
    },
    recentAudit: async (
      _org: string,
      limit: number,
      filter: { agentId: string | null; connectorId: string | null },
    ) =>
      state.audit
        .filter(
          (a) =>
            (!filter.agentId || a.agentId === filter.agentId) &&
            (!filter.connectorId || a.connectorId === filter.connectorId),
        )
        .slice(-limit)
        .map(({ organizationId: _, ...row }) => ({ at: new Date(0).toISOString(), ...row })),
  };
  return { state, store: store as unknown as Store };
}

// rotated: rows are sealed with a previous key that the service still holds.
// shared: the deployment offers approving its own shared Microsoft app.

export function harness({ rotated = false, shared = false } = {}) {
  const previous = randomBytes(32).toString("base64");
  const key = rotated
    ? new CredentialKey(randomBytes(32).toString("base64"), previous)
    : new CredentialKey(previous);
  const vendor = fakeJira();
  const { state, store } = fakeStore(rotated ? new CredentialKey(previous) : key);
  const ms = fakeMicrosoft();
  const msp = fakeMicrosoftProvider({ approval: shared });
  const gr = fakeGraph();
  const apps: App[] = [
    { connector: graph, vendor: gr, toolset: toolsets(graphTools, excelTools) },
    { connector: jira, vendor, toolset: jiraTools },
    { connector: microsoft, vendor: ms, toolset: microsoftTools },
    { connector: calendar, vendor: ms, toolset: calendarTools },
    { connector: teamsConnector, vendor: ms, toolset: teamsTools },
  ];
  const service = createService({
    apps,
    providers: [{ provider: microsoftProvider, client: msp }],
    store,
    key,
  });
  const app = createApp({ service, store, adminToken: ADMIN_TOKEN, validateResponses: true });

  function addAgent(
    capabilities: string[],
    limits: GrantRow["limits"] = {},
    status: AgentRow["status"] = "active",
    connectorId = "jira",
  ) {
    const agentKey = newAgentKey();
    state.agents.push({
      id: AGENT,
      organizationId: ORG,
      name: "Triage",
      status,
      createdBy: null,
      createdAt: "",
      keyHash: hashAgentKey(agentKey),
    });
    state.grants.push({
      agentId: AGENT,
      connectorId,
      capabilities,
      limits,
      updatedAt: "",
    });
    return agentKey;
  }

  let id = 0;
  async function rpc(
    agentKey: string | null,
    method: string,
    params: unknown = {},
    headers: Record<string, string> = {},
  ) {
    const res = await app.fetch(
      new Request("http://svc/agent-connect/mcp", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          accept: "application/json, text/event-stream",
          "mcp-protocol-version": "2025-06-18",
          ...(agentKey ? { authorization: `Bearer ${agentKey}` } : {}),
          ...headers,
        },
        body: JSON.stringify({ jsonrpc: "2.0", id: ++id, method, params }),
      }),
    );
    const text = await res.text();
    const payload = text.startsWith("{")
      ? text
      : text
          .split("\n")
          .find((l) => l.startsWith("data: "))
          ?.slice(6);
    return { status: res.status, body: payload ? JSON.parse(payload) : null };
  }

  async function callTool(agentKey: string, name: string, args: unknown) {
    const { body } = await rpc(agentKey, "tools/call", { name, arguments: args });
    if (body.error) return { error: body.error };
    const r = body.result;
    return { isError: Boolean(r.isError), text: r.content[0].text as string };
  }

  return {
    app,
    state,
    store,
    vendor,
    msp,
    service,
    key,
    addAgent,
    rpc,
    callTool,
  };
}

// Admin changes carry the acting admin; test bodies get one unless they name their own.
export const withBy = (body: unknown) =>
  body && typeof body === "object" && !Array.isArray(body) ? { by: "admin@x", ...body } : body;
