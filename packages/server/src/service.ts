import type { McpServer } from "@modelcontextprotocol/server";
import { brand } from "../../connector-kit/src/branding.js";
import { describeCertificate, newCertificate } from "../../connector-kit/src/certificate.js";
import { signInNonce, signState, verifyState } from "./consent.js";
import {
  effectiveCapabilities,
  pickDisplay,
  pickLabels,
  providerSettings,
  tokenRolesFor,
  type Connector,
  type Probe,
  type Provider,
} from "../../connector-kit/src/connector.js";
import { credentialAad, hashAgentKey, looksLikeAgentKey, type CredentialKey } from "./secrets.js";
import { ProviderChanged } from "./store.js";
import { configSites } from "../../../connectors/sharepoint-files/vendor.js";
import { SHARED_APP_AUTH } from "../../../connectors/microsoft/client.js";
import type {
  AgentRow,
  ConnectionRow,
  GrantRow,
  Limits,
  ProviderRow,
  PublicConnection,
  Store,
} from "./store.js";
import { argumentsTarget, type ToolContext, type Toolset } from "../../connector-kit/src/common.js";
import type { TeamsAccess } from "./teams/runtime.js";
import type { BotRow } from "./teams/store.js";
import {
  type AgentAccess,
  ConnectError,
  type ProviderClient,
  type Vendor,
  type VendorConnection,
  type VendorRequest,
} from "../../connector-kit/src/types.js";

export type App = { connector: Connector; vendor: Vendor; toolset: Toolset };
export type ProviderApp = { provider: Provider; client: ProviderClient };

type CertificateInfo = { thumbprint: string; expiresAt: string };

export type PublicProvider = Omit<
  ProviderRow,
  "credential" | "credentialKeyId" | "credentialScope"
> & {
  certificate: CertificateInfo | null;
  // A certificate made for rotating, not yet switched to.
  nextCertificate: CertificateInfo | null;
  // The apps signing in through it.
  apps: string[];
};

export type CheckResult = {
  label: string;
  values: string[];
  warning?: string;
  labels?: Record<string, string>;
};

export class InvalidInput extends Error {}

// fetch rejects with TypeError("fetch failed") when the app's address does not answer, e.g.
// a mistyped site or domain. That is the admin's to fix, not a server fault.
async function reaching<T>(name: string, run: () => Promise<T>): Promise<T> {
  try {
    return await run();
  } catch (error) {
    if (error instanceof TypeError && error.message === "fetch failed")
      throw new ConnectError(`Could not reach ${name}. Check the address and try again.`);
    throw error;
  }
}

// Validates ticked capabilities and returns them in connector order, with the required ones
// added: they are always on.
export function validateCapabilities(connector: Connector, raw: unknown): string[] {
  if (!Array.isArray(raw)) throw new InvalidInput("capabilities must be a list");
  for (const id of raw)
    if (typeof id !== "string" || !connector.capabilities.some((c) => c.id === id))
      throw new InvalidInput(`${connector.name} has no capability ${String(id)}`);
  const ids = connector.capabilities
    .filter((c) => c.required || raw.includes(c.id))
    .map((c) => c.id);
  if (!ids.length) throw new InvalidInput(`Choose at least one thing ${connector.name} may do`);
  return ids;
}

export function publicConnection(row: ConnectionRow): PublicConnection {
  const { credential: _c, credentialKeyId: _k, ...rest } = row;
  return rest;
}

const certificateInfo = (pem: string | undefined) => (pem ? describeCertificate(pem) : null);

// `sharedClientId`: the shared app's id, for rows approved before the id was stored with them
// (moved from microsoft-365), so their setup steps still show it.
export function publicProvider(
  row: ProviderRow,
  connections: ConnectionRow[],
  sharedClientId: string | null = null,
): PublicProvider {
  const { credential: _c, credentialKeyId: _k, credentialScope: _s, ...rest } = row;
  const shared = row.config.auth === SHARED_APP_AUTH && !row.config.clientId && sharedClientId;
  return {
    ...rest,
    ...(shared ? { config: { ...row.config, clientId: sharedClientId } } : {}),
    certificate: certificateInfo(row.config.certificate),
    nextCertificate: certificateInfo(row.config.nextCertificate),
    apps: connections.filter((c) => c.providerId === row.id).map((c) => c.connectorId),
  };
}

// Same account, reached by the same app. A principal key missing on either side (rows moved
// from before it was stored) is not counted as a change.
export function sameSignIn(
  provider: Provider,
  before: Record<string, string>,
  after: Record<string, string>,
): boolean {
  if (before[provider.accountKey] !== after[provider.accountKey]) return false;
  return provider.principalKeys.every((k) => !before[k] || !after[k] || before[k] === after[k]);
}

// What a provider's credential is bound to. Rows moved from an app connection keep that
// app's id until re-sealed.
const providerScope = (provider: string) => `provider:${provider}`;

// Every tool a connector's capabilities name must exist in its toolset, and every tool in
// the toolset must belong to a capability (otherwise no agent could ever be given it). Every
// vendor check a probe names must exist too.
export function assertToolsImplemented(apps: App[]): void {
  for (const { connector, toolset, vendor } of apps) {
    for (const p of probes(connector))
      if ("check" in p && !vendor.checks?.[p.check])
        throw new Error(`${connector.id}: check not implemented: ${p.check}`);
    if (connector.siteAccess && !vendor.sites)
      throw new Error(`${connector.id}: sites are added, but its client cannot check them`);
    const named = new Set(connector.capabilities.flatMap((l) => l.tools));
    const missing = [...named].filter((t) => !toolset.names.includes(t));
    if (missing.length)
      throw new Error(`${connector.id}: tools not implemented: ${missing.join(", ")}`);
    const orphans = toolset.names.filter((t) => !named.has(t));
    if (orphans.length)
      throw new Error(`${connector.id}: tools in no capability: ${orphans.join(", ")}`);
    const destructive = destructiveTools(toolset);
    for (const c of connector.capabilities) {
      const tool = c.tools.find((t) => destructive.has(t));
      if (tool && !c.destructive)
        throw new Error(
          `${connector.id}: capability ${c.id} has ${tool}, which can destroy or overwrite, so it must be destructive`,
        );
    }
  }
}

// The tools a toolset declares destructive (MCP's destructiveHint), read by registering them
// all on a server that only records.
function destructiveTools(toolset: Toolset): Set<string> {
  const found = new Set<string>();
  const recorder = {
    registerTool(name: string, config: { annotations?: { destructiveHint?: boolean } }) {
      if (config.annotations?.destructiveHint) found.add(name);
    },
  } as unknown as McpServer;
  toolset.register(
    recorder,
    { limits: {}, config: {} } as unknown as ToolContext,
    new Set(toolset.names),
  );
  return found;
}

const probes = (connector: Connector): Probe[] => [
  connector.verify,
  ...(connector.discover ? [connector.discover] : []),
];

export function validateLimits(connector: Connector, raw: unknown): Limits {
  if (raw === undefined || raw === null) raw = {};
  if (typeof raw !== "object" || Array.isArray(raw))
    throw new InvalidInput("limits must be an object");
  const input = raw as Record<string, unknown>;
  const out: Limits = {};
  for (const key of Object.keys(input))
    if (!connector.limits.some((l) => l.id === key))
      throw new InvalidInput(`${connector.name} has no limit ${key}`);
  for (const limit of connector.limits) {
    const value = input[limit.id];
    const pattern = limit.pattern ? new RegExp(limit.pattern) : null;
    if (limit.kind === "map") {
      // Keyed by values of another list limit (checked above, in connector order).
      const keys = (out[limit.keys!] as string[] | undefined) ?? [];
      const map = value ?? {};
      if (typeof map !== "object" || Array.isArray(map))
        throw new InvalidInput(`${limit.label} must name a value for each choice`);
      const kept: Record<string, string> = {};
      for (const [k, v] of Object.entries(map as Record<string, unknown>)) {
        if (!keys.includes(k))
          throw new InvalidInput(`${limit.label}: ${k} is not one of the chosen values`);
        const text = typeof v === "string" ? v.trim() : v;
        if (typeof text !== "string" || text.length > 400)
          throw new InvalidInput(`${limit.label} must be text`);
        if (text && pattern && !pattern.test(text))
          throw new InvalidInput(`${limit.label}: ${text} is not valid`);
        if (text) kept[k] = text;
      }
      if (limit.required && Object.keys(kept).length === 0)
        throw new InvalidInput(`${limit.label} is required`);
      if (Object.keys(kept).length) out[limit.id] = kept;
    } else if (limit.kind === "list") {
      const list = value ?? [];
      if (!Array.isArray(list) || list.some((v) => typeof v !== "string" || !v || v.length > 255))
        throw new InvalidInput(`${limit.label} must be a list of values`);
      if (list.length > 500) throw new InvalidInput(`${limit.label} has too many values`);
      const bad = pattern && list.find((v) => !pattern.test(v));
      if (bad) throw new InvalidInput(`${limit.label}: ${bad} is not valid`);
      if (limit.required && list.length === 0) throw new InvalidInput(`${limit.label} is required`);
      if (list.length) out[limit.id] = [...new Set(list as string[])];
    } else {
      const text = typeof value === "string" ? value.trim() : value;
      if (text !== undefined && (typeof text !== "string" || text.length > 320))
        throw new InvalidInput(`${limit.label} must be text`);
      if (text && pattern && !pattern.test(text))
        throw new InvalidInput(`${limit.label}: ${text} is not valid`);
      if (limit.required && !text) throw new InvalidInput(`${limit.label} is required`);
      if (text) out[limit.id] = text;
    }
  }
  return out;
}

// A choice field's value must be one of its options; left empty, it is the first.
export function validateChoices(
  connector: Connector,
  fields: Record<string, string>,
): Record<string, string> {
  const out = { ...fields };
  for (const field of connector.fields) {
    if (!field.options) continue;
    const value = out[field.id]?.trim() || field.options[0].value;
    if (!field.options.some((o) => o.value === value))
      throw new InvalidInput(`${field.label}: ${value} is not one of the choices`);
    out[field.id] = value;
  }
  return out;
}

// One user by id or sign-in name, its id and name only.
const DIRECTORY_LOOKUP = /^\/v1\.0\/users\/[^/?#]+\?\$select=id,displayName$/;

export function createService(deps: {
  apps: App[];
  providers?: ProviderApp[];
  store: Store;
  key: CredentialKey;
  // Apps that sign in as the agent's own bot (connector `signsInAs: "agentBot"`).
  bots?: {
    botForAgent(org: string, agentId: string): Promise<BotRow | null>;
    access(bot: BotRow): TeamsAccess;
  };
}) {
  const { store, key } = deps;
  const stateKey = key.derive("consent-state");
  const apps = new Map(deps.apps.map((a) => [a.connector.id, a]));
  const providers = new Map((deps.providers ?? []).map((p) => [p.provider.id, p]));
  for (const { connector } of deps.apps)
    if (connector.provider && !providers.has(connector.provider))
      throw new Error(`${connector.id}: provider ${connector.provider} is not loaded`);
  for (const { provider } of providers.values())
    for (const step of provider.addAgent)
      for (const id of step.apps)
        if (apps.get(id)?.connector.provider !== provider.id)
          throw new Error(
            `${provider.id}: add-agent step "${step.title}" names ${id}, not one of its apps`,
          );

  // Each confirmable step id and who owns it (for its audit row): the provider when one of
  // its steps carries the id, else the first app that does.
  const confirmable = new Map<string, string>();
  for (const { connector } of deps.apps)
    for (const s of [...connector.setup, ...connector.afterConnect])
      if (s.confirm && !confirmable.has(s.confirm)) confirmable.set(s.confirm, connector.id);
  for (const { provider } of providers.values())
    for (const s of provider.setup) if (s.confirm) confirmable.set(s.confirm, provider.id);

  const sharedClientId = (id: string) => providers.get(id)?.client.approval?.clientId ?? null;

  function app(connectorId: string): App {
    const found = apps.get(connectorId);
    if (!found) throw new InvalidInput(`Unknown app ${connectorId}`);
    return found;
  }

  function providerApp(id: string): ProviderApp {
    const found = providers.get(id);
    if (!found) throw new InvalidInput(`Unknown connection ${id}`);
    return found;
  }

  function approvalFor(connectorId: string) {
    const target = app(connectorId);
    const approval = target.connector.provider
      ? providerApp(target.connector.provider).client.approval
      : undefined;
    if (!approval)
      throw new InvalidInput(`${target.connector.name} is connected with a form, not an approval`);
    return { target, approval, provider: providerApp(target.connector.provider!) };
  }

  function openProvider(row: ProviderRow): Record<string, string> {
    return key.open(
      { ciphertext: row.credential, keyId: row.credentialKeyId },
      credentialAad(row.organizationId, row.credentialScope),
    );
  }

  // Moves a provider credential sealed with the previous key, or under a former app's id,
  // onto the current key and its own scope.
  async function freshen(row: ProviderRow, secret: Record<string, string>) {
    const scope = providerScope(row.provider);
    if (row.credentialKeyId === key.id && row.credentialScope === scope) return;
    const sealed = key.seal(secret, credentialAad(row.organizationId, scope));
    // False when an admin changed the credential meanwhile: theirs is sealed afresh already.
    await store.updateProviderCredential(
      row.id,
      row.config,
      sealed.ciphertext,
      sealed.keyId,
      scope,
      row.credential,
    );
  }

  // Changes to a provider's sign-in, kept with the agents' tool calls (no agent; `by` is the
  // admin). Shown nowhere yet, but there for an audit.
  async function logAdmin(
    org: string,
    by: string,
    provider: string,
    action: string,
    target: string,
  ) {
    await store.audit({
      organizationId: org,
      agentId: null,
      agentName: by,
      connectorId: provider,
      tool: `admin:${action}`,
      target,
      decision: "allowed",
      reason: null,
      vendorStatus: null,
    });
  }

  // How the admin can tell sign-ins apart: the app, the method, and the certificate.
  function signInLabel(config: Record<string, string>): string {
    const cert = config.certificate
      ? ` certificate ${describeCertificate(config.certificate).thumbprint}`
      : "";
    return `${config.auth ?? "none"} app ${config.clientId ?? "-"} tenant ${config.tenantId ?? "-"}${cert}`;
  }

  // The connection an app's calls go through: its own credential, or its provider's with the
  // provider's settings under its own.
  function resolve(row: ConnectionRow, providerRows: ProviderRow[]): VendorConnection {
    if (!row.providerId)
      return {
        config: row.config,
        secret: key.open(
          { ciphertext: row.credential!, keyId: row.credentialKeyId! },
          credentialAad(row.organizationId, row.connectorId),
        ),
      };
    const provider = providerRows.find((p) => p.id === row.providerId);
    if (!provider) throw new Error(`${row.connectorId} connection has no provider row`);
    return { config: { ...provider.config, ...row.config }, secret: openProvider(provider) };
  }

  async function save(
    target: App,
    input: {
      organizationId: string;
      connectorId: string;
      capabilities: unknown;
      fields: Record<string, string>;
      connectedBy: string;
    },
  ) {
    const org = input.organizationId;
    const capabilities = validateCapabilities(target.connector, input.capabilities);
    let provider: ProviderRow | null = null;
    let account: (VendorConnection & { accountLabel: string }) | undefined;
    if (target.connector.provider) {
      const { provider: def } = providerApp(target.connector.provider);
      provider = await store.getProvider(org, def.id);
      if (!provider || provider.status === "pending")
        throw new ConnectError(`Connect ${def.name} first`);
      account = {
        config: provider.config,
        secret: openProvider(provider),
        accountLabel: provider.accountLabel ?? "",
      };
    }
    const fields = validateChoices(target.connector, input.fields);
    const previous = await store.getConnection(org, input.connectorId);
    const { conn, checks } = await reaching(target.connector.name, async () => {
      const conn = await target.vendor.connect(fields, account, previous?.config);
      const full = account
        ? { config: { ...account.config, ...conn.config }, secret: account.secret }
        : conn;
      return { conn, checks: await probe(target, full, capabilities) };
    });
    const accountKey = target.connector.accountKey;
    const accountOf = (config: Record<string, string>) =>
      ({ ...provider?.config, ...config })[accountKey];
    if (!accountOf(conn.config))
      throw new Error(`${target.connector.id} connection has no ${accountKey}`);
    const sealed = provider ? null : key.seal(conn.secret, credentialAad(org, input.connectorId));
    await store.saveConnection({
      clearGrants: previous !== null && accountOf(previous.config) !== accountOf(conn.config),
      organizationId: org,
      connectorId: input.connectorId,
      capabilities,
      config: conn.config,
      credential: sealed?.ciphertext ?? null,
      credentialKeyId: sealed?.keyId ?? null,
      providerId: provider?.id ?? null,
      accountLabel: conn.accountLabel,
      connectedBy: input.connectedBy,
    });
    const row = await store.getConnection(org, input.connectorId);
    if (!row) throw new Error("Connection vanished after saving");
    return { connection: publicConnection(row), checks };
  }

  async function saveProvider(
    org: string,
    { provider }: ProviderApp,
    conn: VendorConnection & { accountLabel: string },
    connectedBy: string,
    existing: ProviderRow | null,
  ) {
    const accountKey = provider.accountKey;
    if (!conn.config[accountKey]) throw new Error(`${provider.id} connection has no ${accountKey}`);
    const scope = providerScope(provider.id);
    const sealed = key.seal(conn.secret, credentialAad(org, scope));
    const removeApps =
      existing !== null &&
      existing.status !== "pending" &&
      !sameSignIn(provider, existing.config, conn.config);
    // What the admin entered in the apps' steps (e.g. the Azure resource group for Teams bots)
    // stays with the same sign-in.
    const settings = providerSettings(
      deps.apps.map((a) => a.connector),
      provider.id,
    );
    const kept =
      existing && !removeApps
        ? Object.fromEntries(
            Object.entries(existing.config).filter(([k]) => settings.has(k) && !(k in conn.config)),
          )
        : {};
    const removed = removeApps
      ? (await store.listConnections(org))
          .filter((c) => c.providerId === existing!.id)
          .map((c) => c.connectorId)
      : [];
    await store.saveProvider({
      organizationId: org,
      provider: provider.id,
      config: { ...kept, ...conn.config },
      credential: sealed.ciphertext,
      credentialKeyId: sealed.keyId,
      credentialScope: scope,
      accountLabel: conn.accountLabel,
      connectedBy,
      removeApps,
      replaces: existing?.credential ?? null,
    });
    const before =
      existing && existing.status !== "pending" ? signInLabel(existing.config) : "none";
    await logAdmin(
      org,
      connectedBy,
      provider.id,
      "connect",
      `${before} -> ${signInLabel(conn.config)}${removed.length ? `; disconnected ${removed.join(", ")}` : ""}`,
    );
    return providerView(org, provider.id);
  }

  async function providerView(org: string, id: string) {
    const [row, connections] = await Promise.all([
      store.getProvider(org, id),
      store.listConnections(org),
    ]);
    if (!row) throw new Error("Provider connection vanished after saving");
    return publicProvider(row, connections, sharedClientId(id));
  }

  async function probe(
    { connector, vendor }: App,
    conn: VendorConnection,
    ceiling: string[],
  ): Promise<CheckResult[]> {
    const results: CheckResult[] = [];
    for (const p of probes(connector)) {
      if ("check" in p) {
        const { values, warning, labels } = await vendor.checks![p.check](conn, {
          expected: tokenRolesFor(connector, ceiling, conn.config),
        });
        // The vendor's own warning says more than the generic "nothing yet".
        const shown = warning || (values.length === 0 ? p.ifEmpty : undefined);
        results.push({
          label: p.label,
          values,
          ...(shown ? { warning: shown } : {}),
          ...(labels ? { labels } : {}),
        });
        continue;
      }
      // {name} in a probe path is the connection's config value, e.g. the signed-in user's id.
      const path = p.path.replace(/\{([A-Za-z_]+)\}/g, (_, name: string) => {
        const value = conn.config[name];
        if (!value) throw new Error(`${connector.id} probe needs config ${name}`);
        return encodeURIComponent(value);
      });
      const res = await vendor.call(conn, { method: p.method, path, body: p.body });
      if (res.status < 200 || res.status >= 300)
        throw new ConnectError(`${p.label}: ${connector.name} answered ${res.status}`);
      const values = pickDisplay(res.data, p.show, p.match);
      results.push({
        label: p.label,
        values,
        ...(values.length === 0 && p.ifEmpty ? { warning: p.ifEmpty } : {}),
        ...(p.name ? { labels: pickLabels(res.data, p.show, p.name, p.match) } : {}),
      });
    }
    return results;
  }

  return {
    connectors: () => deps.apps.map(({ connector }) => connector),

    providers: () =>
      [...providers.values()].map(({ provider, client }) => ({
        ...provider,
        // The deployment offers approving its own shared app instead of registering one.
        approval: Boolean(client.approval),
      })),

    async listConnections(org: string) {
      return (await store.listConnections(org))
        .filter((r) => apps.has(r.connectorId))
        .map(publicConnection);
    },

    async listProviders(org: string) {
      const [rows, connections] = await Promise.all([
        store.listProviders(org),
        store.listConnections(org),
      ]);
      return rows
        .filter((r) => providers.has(r.provider))
        .map((r) => publicProvider(r, connections, sharedClientId(r.provider)));
    },

    // Setup steps the admin marked done, shared by every admin of the organization.
    async listConfirmed(org: string) {
      return store.listConfirmed(org);
    },

    async setConfirmed(org: string, step: string, by: string, on: boolean) {
      const owner = confirmable.get(step);
      if (!owner) throw new InvalidInput(`No setup step ${step} to confirm`);
      await store.setConfirmed(org, step, by, on);
      await logAdmin(org, by, owner, on ? "confirm_step" : "unconfirm_step", step);
    },

    // Guided apps: the admin's form fields. Verified before anything is stored.
    async connect(input: {
      organizationId: string;
      connectorId: string;
      capabilities: unknown;
      fields: Record<string, string>;
      connectedBy: string;
    }) {
      return save(app(input.connectorId), input);
    },

    // The provider's own sign-in, shared by its apps. Proven before it is stored. Reaching
    // another account than before disconnects its apps.
    async connectProvider(input: {
      organizationId: string;
      provider: string;
      fields: Record<string, string>;
      connectedBy: string;
    }) {
      const target = providerApp(input.provider);
      const existing = await store.getProvider(input.organizationId, input.provider);
      const conn = await reaching(target.provider.name, () =>
        target.client.connect(
          input.fields,
          existing ? { config: existing.config, secret: openProvider(existing) } : null,
        ),
      );
      return saveProvider(input.organizationId, target, conn, input.connectedBy, existing);
    },

    // A new key pair for the provider. Before it is connected, the certificate to upload
    // first; after, one to rotate to, kept apart until the admin switches to it.
    async newCertificate(org: string, id: string, by: string) {
      const { provider } = providerApp(id);
      if (!provider.certificate)
        throw new InvalidInput(`${provider.name} does not sign in with a certificate`);
      const made = newCertificate(brand().microsoft.appName);
      const scope = providerScope(id);
      const row = await store.getProvider(org, id);
      if (!row || row.status === "pending") {
        const sealed = key.seal({ private_key: made.privateKey }, credentialAad(org, scope));
        const saved = await store.savePendingProvider({
          organizationId: org,
          provider: id,
          config: { certificate: made.certificate },
          credential: sealed.ciphertext,
          credentialKeyId: sealed.keyId,
          credentialScope: scope,
        });
        if (!saved)
          throw new ConnectError(`${provider.name} was connected meanwhile. Reload the page.`);
      } else {
        if (row.config.auth === SHARED_APP_AUTH)
          throw new InvalidInput(
            `This connection uses ${brand().productName}'s shared app, which has no certificate of yours`,
          );
        const secret = { ...openProvider(row), next_private_key: made.privateKey };
        const sealed = key.seal(secret, credentialAad(org, scope));
        const saved = await store.updateProviderCredential(
          row.id,
          { ...row.config, nextCertificate: made.certificate },
          sealed.ciphertext,
          sealed.keyId,
          scope,
          row.credential,
        );
        if (!saved) throw new ProviderChanged();
      }
      await logAdmin(
        org,
        by,
        id,
        "create_certificate",
        `${!row || row.status === "pending" ? "first" : "next"} certificate ${describeCertificate(made.certificate).thumbprint}`,
      );
      return providerView(org, id);
    },

    // Signs in with the new certificate and, only if that works, makes it the one used. The
    // old one keeps working in the app until the admin removes it there.
    async activateCertificate(org: string, id: string, by: string) {
      const { provider, client } = providerApp(id);
      const row = await store.getProvider(org, id);
      if (!row || row.status === "pending")
        throw new InvalidInput(`${provider.name} is not connected`);
      const { nextCertificate, ...config } = row.config;
      const { next_private_key, client_secret: _old, ...secret } = openProvider(row);
      if (!nextCertificate || !next_private_key)
        throw new InvalidInput("There is no new certificate to switch to");
      const candidate = {
        config: { ...config, auth: "certificate", certificate: nextCertificate },
        secret: { ...secret, private_key: next_private_key },
      };
      await reaching(provider.name, () => client.verify(candidate));
      const scope = providerScope(id);
      const sealed = key.seal(candidate.secret, credentialAad(org, scope));
      const saved = await store.updateProviderCredential(
        row.id,
        candidate.config,
        sealed.ciphertext,
        sealed.keyId,
        scope,
        row.credential,
        true,
      );
      if (!saved) throw new ProviderChanged();
      await logAdmin(
        org,
        by,
        id,
        "switch_certificate",
        `${signInLabel(row.config)} -> ${signInLabel(candidate.config)}`,
      );
      return providerView(org, id);
    },

    async checkProvider(org: string, id: string) {
      const { client } = providerApp(id);
      const row = await store.getProvider(org, id);
      if (!row) return null;
      if (row.status === "pending") return { ok: false as const, error: "Not connected yet" };
      const secret = openProvider(row);
      try {
        await reaching(providerApp(id).provider.name, () =>
          client.verify({ config: row.config, secret }),
        );
        await store.markProviderChecked(row.id, true, row.credential);
        await freshen(row, secret);
        return { ok: true as const };
      } catch (error) {
        if (!(error instanceof ConnectError)) throw error;
        await store.markProviderChecked(row.id, false, row.credential);
        return { ok: false as const, error: error.message };
      }
    },

    // Values the provider's apps ask for in their steps (connector `settings`), kept in the
    // provider's config. Only over the credential read: a reconnect meanwhile wins.
    async setProviderSettings(org: string, id: string, values: Record<string, string>, by: string) {
      const { provider } = providerApp(id);
      const allowed = providerSettings(
        deps.apps.map((a) => a.connector),
        provider.id,
      );
      const clean: Record<string, string> = {};
      for (const [k, raw] of Object.entries(values)) {
        const setting = allowed.get(k);
        if (!setting) throw new InvalidInput(`${provider.name} has no setting ${k}`);
        const value = raw.trim();
        if (!new RegExp(`^(?:${setting.pattern})$`).test(value))
          throw new InvalidInput(`${setting.label}: ${setting.help ?? "not a valid value"}`);
        clean[k] = value;
      }
      const row = await store.getProvider(org, provider.id);
      if (!row || row.status === "pending")
        throw new ConnectError(`Connect ${provider.name} first`);
      const saved = await store.updateProviderCredential(
        row.id,
        { ...row.config, ...clean },
        row.credential,
        row.credentialKeyId,
        row.credentialScope,
        row.credential,
      );
      if (!saved) throw new ProviderChanged();
      await logAdmin(
        org,
        by,
        provider.id,
        "settings",
        Object.entries(clean)
          .map(([k, v]) => `${k}=${v}`)
          .join(" "),
      );
      return { settings: clean };
    },

    async disconnectProvider(org: string, id: string, by: string) {
      const { provider } = providerApp(id);
      const row = await store.getProvider(org, provider.id);
      const apps = row
        ? (await store.listConnections(org))
            .filter((c) => c.providerId === row.id)
            .map((c) => c.connectorId)
        : [];
      const removed = await store.removeProvider(org, provider.id);
      if (removed && row)
        await logAdmin(
          org,
          by,
          provider.id,
          "disconnect",
          `${signInLabel(row.config)}; disconnected ${apps.join(", ") || "no apps"}`,
        );
      return removed;
    },

    // Approval of the deployment's shared app, where offered: where the admin goes to approve. The
    // nonce goes back to the admin app, which keeps it in the admin's browser until the
    // approval returns.
    startConsent(input: {
      organizationId: string;
      connectorId: string;
      capabilities: unknown;
      redirectUri: string;
    }) {
      const { target, approval } = approvalFor(input.connectorId);
      const { state, nonce } = signState(stateKey, {
        org: input.organizationId,
        app: input.connectorId,
        ceiling: validateCapabilities(target.connector, input.capabilities),
        redirectUri: input.redirectUri,
      });
      return { url: approval.consentUrl({ state, redirectUri: input.redirectUri }), nonce };
    },

    // What the app sent back, twice. First the approval, which names the organization that
    // approved but is unsigned, so the admin is sent on to sign in there (`redirect`). Then
    // that sign-in, which proves the organization; only then is the provider connection
    // saved, and the app too when it needs no settings of its own. The ceiling and the
    // organization come from the signed state, never from the callback.
    async finishConsent(input: {
      organizationId: string;
      connectorId: string;
      state: string;
      nonce: string;
      params: Record<string, string>;
      connectedBy: string;
    }): Promise<
      { redirect: string } | { provider: PublicProvider } | Awaited<ReturnType<typeof save>>
    > {
      const { target, approval, provider } = approvalFor(input.connectorId);
      const state = verifyState(stateKey, input.state, {
        org: input.organizationId,
        app: input.connectorId,
        nonce: input.nonce,
      });
      const name = provider.provider.name;
      const { error, error_description: description } = input.params;
      if (error)
        throw new ConnectError(
          `${name} did not approve ${brand().productName}: ${(description ?? error).split(/\r?\n/)[0].slice(0, 300)}`,
        );
      if (!state.tenant) {
        if (input.params.admin_consent?.toLowerCase() !== "true")
          throw new ConnectError(`${name} did not approve ${brand().productName}`);
        const tenant = (input.params.tenant ?? "").trim().toLowerCase();
        const next = signState(
          stateKey,
          {
            org: state.org,
            app: state.app,
            ceiling: state.ceiling,
            redirectUri: state.redirectUri,
            tenant,
          },
          Date.now(),
          input.nonce,
        );
        return {
          redirect: approval.signInUrl({
            tenant,
            state: next.state,
            redirectUri: state.redirectUri,
            nonce: signInNonce(input.nonce),
          }),
        };
      }
      const code = input.params.code;
      if (!code) throw new ConnectError(`${name} did not confirm the sign-in. Start again.`);
      const conn = await reaching(name, async () => {
        await approval.proveTenant({
          tenant: state.tenant!,
          code,
          redirectUri: state.redirectUri,
          nonce: signInNonce(input.nonce),
        });
        return approval.connect(state.tenant!);
      });
      const existing = await store.getProvider(input.organizationId, provider.provider.id);
      const saved = await saveProvider(
        input.organizationId,
        provider,
        conn,
        input.connectedBy,
        existing,
      );
      if (target.connector.fields.length) return { provider: saved };
      return save(target, {
        organizationId: input.organizationId,
        connectorId: input.connectorId,
        capabilities: state.ceiling,
        fields: {},
        connectedBy: input.connectedBy,
      });
    },

    // `record: false` only reads (e.g. the choices for a key's limits): a failure there must
    // not mark the connection broken and cut every agent off. A recorded check that changes the
    // connection's status is audited under `by`, the admin who asked for it.
    async check(
      org: string,
      connectorId: string,
      { record = true, by }: { record?: boolean; by?: string } = {},
    ) {
      const target = app(connectorId);
      const row = await store.getConnection(org, connectorId);
      if (!row) return null;
      const providerRows = row.providerId ? await store.listProviders(org) : [];
      const before = row.status;
      const recordCheck = async (ok: boolean, target: string) => {
        await store.markChecked(row.id, ok);
        if (by && before !== (ok ? "connected" : "needs_reconnect"))
          await logAdmin(org, by, connectorId, "check_app", target);
      };
      try {
        const conn = resolve(row, providerRows);
        const checks = await reaching(target.connector.name, () =>
          probe(target, conn, row.capabilities),
        );
        // Moves a row sealed with the previous key onto the current one, read only or not
        // (pages only read): it changes no status.
        const provider = providerRows.find((p) => p.id === row.providerId);
        if (provider) await freshen(provider, conn.secret);
        else if (row.credentialKeyId !== key.id) {
          const sealed = key.seal(conn.secret, credentialAad(org, connectorId));
          // Only over the credential read: one saved by a reconnect meanwhile stays.
          await store.reseal(row.id, row.credential!, sealed.ciphertext, sealed.keyId);
        }
        if (record) await recordCheck(true, "ok");
        return { ok: true as const, checks };
      } catch (error) {
        if (!(error instanceof ConnectError)) throw error;
        if (record) await recordCheck(false, `failed: ${error.message}`);
        return { ok: false as const, error: error.message };
      }
    },

    // Sites added after connecting, the whole list at once (adding or removing one).
    async setSites(org: string, connectorId: string, raw: string[]) {
      const { connector, vendor } = app(connectorId);
      if (!connector.siteAccess || !vendor.sites)
        throw new InvalidInput(`${connector.name} has no sites to add`);
      const row = await store.getConnection(org, connectorId);
      if (!row) return null;
      const sites = vendor.sites.normalize(raw);
      const { sites: _old, ...rest } = row.config;
      const config = sites.length ? { ...rest, sites: sites.join("\n") } : rest;
      // Only over the config read here: a reconnect meanwhile must not get the old one back.
      if (!(await store.updateConfig(org, connectorId, config, row.config)))
        throw new ProviderChanged();
      return { sites };
    },

    // Whether Microsoft gave the app each added site (or the one asked about), read live.
    async siteAccess(org: string, connectorId: string, only?: string) {
      const { connector, vendor } = app(connectorId);
      if (!connector.siteAccess || !vendor.sites)
        throw new InvalidInput(`${connector.name} has no sites to check`);
      const row = await store.getConnection(org, connectorId);
      if (!row) return null;
      const added = configSites(row.config);
      const sites = only === undefined ? added : vendor.sites.normalize([only]);
      if (only !== undefined && !added.includes(sites[0]))
        throw new InvalidInput(`${only} is not one of the added sites`);
      const conn = resolve(row, row.providerId ? await store.listProviders(org) : []);
      const write = connector.siteAccess.write.some((id) => row.capabilities.includes(id));
      const results = [];
      for (const site of sites)
        results.push(await reaching(connector.name, () => vendor.sites!.access(conn, site, write)));
      return { sites: results };
    },

    // Whether the app can reach one agent's own address (its mailbox) yet, read live.
    // A check of an address no key holds yet is recorded, so lookups of other mailboxes in the
    // tenant leave a trace.
    async agentAccess(org: string, connectorId: string, value: string, by: string) {
      const { connector, vendor } = app(connectorId);
      const access = connector.agentAccess;
      if (!access) throw new InvalidInput(`${connector.name} has no agent address to check`);
      const limit = connector.limits.find((l) => l.id === access.limit)!;
      if (limit.pattern && !new RegExp(limit.pattern).test(value))
        throw new InvalidInput(`${value} is not a valid ${limit.label.toLowerCase()}`);
      const row = await store.getConnection(org, connectorId);
      if (!row) return null;
      const conn = resolve(row, row.providerId ? await store.listProviders(org) : []);
      const held = (await store.listAgents(org)).some((a) =>
        a.grants.some(
          (g) =>
            g.connectorId === connectorId &&
            String(g.limits[access.limit] ?? "").toLowerCase() === value.toLowerCase(),
        ),
      );
      if (!held) await logAdmin(org, by, connectorId, "agent_access_check", value);
      const res = await reaching(connector.name, () =>
        vendor.call(conn, {
          method: "GET",
          path: access.path.replace("{value}", encodeURIComponent(value)),
        }),
      );
      const status: AgentAccess["status"] =
        res.status >= 200 && res.status < 300
          ? "granted"
          : res.status === 401 || res.status === 403
            ? "not_granted"
            : res.status === 404
              ? "not_found"
              : "failed";
      const message = {
        granted: access.granted,
        not_granted: access.notGranted,
        not_found: access.notFound,
        failed: `${connector.name} answered ${res.status}. Check again later.`,
      }[status];
      if (status !== "granted") return { value, status, message } satisfies AgentAccess;
      // Opening it is not enough: each ticked capability's own role, proven by a call that
      // changes nothing.
      const missing: string[] = [];
      for (const capability of connector.capabilities) {
        const probe = access.probes[capability.id];
        if (!probe || !row.capabilities.includes(capability.id)) continue;
        const answer = await reaching(connector.name, () =>
          vendor.call(conn, {
            method: probe.method,
            path: probe.path.replace("{value}", encodeURIComponent(value)),
            body: probe.body,
          }),
        );
        if (answer.status === 401 || answer.status === 403)
          missing.push(`${capability.label} (${capability.scopes.join(", ")})`);
        else if (!probe.allowed.includes(answer.status))
          return {
            value,
            status: "failed",
            message: `${connector.name} answered ${answer.status} when checking ${capability.label}. Check again later.`,
          } satisfies AgentAccess;
      }
      if (missing.length)
        return {
          value,
          status: "not_granted",
          message: access.missing!.replace("{missing}", missing.join(", ")),
        } satisfies AgentAccess;
      return { value, status, message } satisfies AgentAccess;
    },

    // The saved ceiling, or null when the app is not connected.
    async setCeiling(org: string, connectorId: string, ceiling: unknown) {
      const caps = validateCapabilities(app(connectorId).connector, ceiling);
      return (await store.setCeiling(org, connectorId, caps)) ? caps : null;
    },

    async setGrant(
      org: string,
      agentId: string,
      connectorId: string,
      capabilities: unknown,
      limits: unknown,
    ) {
      const { connector, vendor } = app(connectorId);
      const caps = validateCapabilities(connector, capabilities);
      const checked = validateLimits(connector, limits);
      const row = await store.getConnection(org, connectorId);
      if (!row)
        throw new InvalidInput(`Connect ${connector.name} before giving a key access to it`);
      // Never stored above the app's ceiling: raising the ceiling later must not turn it on.
      const above = (ceiling: string[]) => {
        const over = caps.filter((c) => !ceiling.includes(c));
        if (over.length)
          throw new InvalidInput(
            `${connector.name} does not allow ${over
              .map((id) => connector.capabilities.find((c) => c.id === id)!.label)
              .join(", ")}. Turn it on for the app first, or untick it.`,
          );
      };
      above(row.capabilities);
      if (connector.siteAccess && vendor.sites) {
        const added = configSites(row.config);
        const missing = ((checked.sites as string[] | undefined) ?? []).find(
          (s) => !added.includes(vendor.sites!.normalize([s])[0]),
        );
        if (missing)
          throw new InvalidInput(`${missing} is not one of the sites added to ${connector.name}`);
      }
      const saved = await store.setGrant(org, agentId, connectorId, caps, checked);
      // Lost to a ceiling lowered meanwhile: say so rather than "not found".
      if (!saved) {
        const now = await store.getConnection(org, connectorId);
        if (!now)
          throw new InvalidInput(`Connect ${connector.name} before giving a key access to it`);
        above(now.capabilities);
        return null;
      }
      return { capabilities: caps, limits: checked };
    },

    /** A call with the organization's own connection, no agent's (an event poll or a webhook
     *  registration). Null when the app is not connected. */
    async appCall(org: string, connectorId: string, request: VendorRequest) {
      const target = app(connectorId);
      const row = await store.getConnection(org, connectorId);
      if (!row || row.status !== "connected") return null;
      const providerRows = row.providerId ? await store.listProviders(org) : [];
      return target.vendor.call(resolve(row, providerRows), request);
    },

    // Resolves an agent key to the tools that agent may call right now.
    async agentSession(rawKey: string) {
      if (!looksLikeAgentKey(rawKey)) return null;
      const agent = await store.agentByKeyHash(hashAgentKey(rawKey));
      if (!agent || agent.status !== "active") return null;
      const [grants, connections, providerRows] = await Promise.all([
        store.grantsFor(agent.id),
        store.listConnections(agent.organizationId),
        store.listProviders(agent.organizationId),
      ]);
      const needsBot = grants.some((g) => apps.get(g.connectorId)?.connector.signsInAs);
      if (needsBot && !deps.bots)
        throw new Error("An app signs in as the agent's bot, but bots are not loaded");
      const bot = needsBot ? await deps.bots!.botForAgent(agent.organizationId, agent.id) : null;
      const toolsets = sessionToolsets(agent, grants, connections, providerRows, bot);
      const granted = new Set(toolsets.flatMap((t) => [...t.allowed]));
      return {
        agent,
        toolsets,
        // A call to one of our tools the agent was not given never reaches a handler (the MCP
        // server answers "not found"), so it is recorded here.
        async refuseUngranted(tool: string, args: unknown) {
          if (granted.has(tool)) return;
          const owner = deps.apps.find((a) => a.toolset.names.includes(tool));
          if (!owner) return;
          await store.audit({
            organizationId: agent.organizationId,
            agentId: agent.id,
            agentName: agent.name,
            connectorId: owner.connector.id,
            tool,
            target: argumentsTarget(args),
            decision: "denied",
            reason: "This agent is not allowed to do this",
            vendorStatus: null,
          });
        },
      };
    },
  };

  function sessionToolsets(
    agent: AgentRow,
    grants: GrantRow[],
    connections: ConnectionRow[],
    providerRows: ProviderRow[],
    bot: BotRow | null,
  ) {
    return grants.flatMap((grant) => {
      const target = apps.get(grant.connectorId);
      const row = connections.find((c) => c.connectorId === grant.connectorId);
      if (!target || !row || row.status !== "connected") return [];
      const provider = row.providerId ? providerRows.find((p) => p.id === row.providerId) : null;
      if (provider !== null && provider?.status !== "connected") return [];
      const capabilities = effectiveCapabilities(
        target.connector,
        grant.capabilities,
        row.capabilities,
      );
      // Its own bot, once that is set up; until then the agent gets none of these tools.
      const asBot = target.connector.signsInAs === "agentBot";
      if (asBot && (bot?.status !== "connected" || !bot.botAppId)) return [];
      let conn: VendorConnection | null = null;
      const ctx: ToolContext = {
        limits: grant.limits,
        config: provider ? { ...provider.config, ...row.config } : row.config,
        call: (request) => {
          // An app signing in as the agent's bot reaches the Microsoft connection only to look
          // a person up in the directory (Teams' approver), never to read or write anything else.
          if (asBot && !(request.method === "GET" && DIRECTORY_LOOKUP.test(request.path)))
            throw new Error(`${grant.connectorId} calls go through the agent's bot`);
          return target.vendor.call((conn ??= resolve(row, providerRows)), request);
        },
        forbidden: Object.fromEntries(
          capabilities.flatMap((c) =>
            c.forbidden
              ? c.tools.map((t) => [t, c.forbidden!.replaceAll("{scopes}", c.scopes.join(", "))])
              : [],
          ),
        ),
        ...(asBot ? { teams: deps.bots!.access(bot!) } : {}),
        audit: (entry) =>
          store.audit({
            organizationId: agent.organizationId,
            agentId: agent.id,
            agentName: agent.name,
            connectorId: grant.connectorId,
            ...entry,
          }),
      };
      return [
        { toolset: target.toolset, ctx, allowed: new Set(capabilities.flatMap((c) => c.tools)) },
      ];
    });
  }
}

export type Service = ReturnType<typeof createService>;
