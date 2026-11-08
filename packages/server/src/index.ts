import { existsSync, readdirSync } from "node:fs";
import { pathToFileURL } from "node:url";
import { createApp } from "./app.js";
import {
  parseBranding,
  setBranding,
  type BrandingInput,
} from "../../connector-kit/src/branding.js";
import { builtinConnectors, type BuiltinModule, type ConnectorModule } from "./builtin.js";
import httpConnector from "../../../connectors/http/index.js";
import odataConnector from "../../../connectors/odata/index.js";
import soapConnector from "../../../connectors/soap/index.js";
import sqlConnector from "../../../connectors/sql/index.js";
import {
  CONNECTORS_DIR,
  loadConnectors,
  loadProviders,
  parseConnector,
  type Connector,
  type ConnectorJson,
} from "../../connector-kit/src/connector.js";
import { migrate as runMigrations } from "./migrate.js";
import { CredentialKey } from "./secrets.js";
import { assertToolsImplemented, createService, type App, type ProviderApp } from "./service.js";
import { errorReporter, reportingVendorRefusals, type ErrorReporting } from "./report.js";
import { assertSchemaName, connectDatabase, createStore, type Sql } from "./store.js";
import { createEventsAdmin, resealEventSecrets } from "./events/admin.js";
import { createEventsRuntime } from "./events/runtime.js";
import { createEventsStore } from "./events/store.js";
import { createTeamsAdmin, resealTeamsCredentials } from "./teams/admin.js";
import { botClient } from "./teams/bot.js";
import { claudeClient } from "./teams/claude.js";
import { botFrameworkAuth } from "./teams/jwt.js";
import { LIVE_CHECKS } from "./teams/provision.js";
import {
  claudeManagedAgents,
  MAX_RETRY_WINDOW_MS,
  RETRY_WINDOW_MS,
  signedWebhook,
  type RunnerFactory,
} from "./teams/runner.js";
import { createTeamsRuntime } from "./teams/runtime.js";
import { createTeamsStore } from "./teams/store.js";
import { microsoftClient } from "../../../connectors/microsoft/client.js";

export type { BrandingInput } from "../../connector-kit/src/branding.js";
export type { ConnectorContext, ConnectorModule } from "./builtin.js";
export type { ConnectorJson } from "../../connector-kit/src/connector.js";
export type { ErrorReporting } from "./report.js";
export { builtinConnectors } from "./builtin.js";
export type { EventRunInput, RunInput, RunnerFactory, TriggerRunner } from "./teams/runner.js";
export {
  claudeManagedAgents,
  DELIVERY_HEADER,
  RetryLater,
  signature,
  signedWebhook,
} from "./teams/runner.js";

export type AgentConnectOptions = {
  // A postgres.js instance (left open on stop), or how to make one.
  database: Sql | { url: string };
  schema?: string;
  // Apply pending migrations before anything else (default true).
  migrate?: boolean;
  encryptionKey: string;
  // Only while rotating the key; see CredentialKey.
  previousEncryptionKey?: string;
  adminToken: string;
  // The deployment's public origin; the service is under <origin>/agent-connect. Without one,
  // Teams bots cannot be set up.
  publicBaseUrl?: string | null;
  branding?: BrandingInput;
  // The apps this deployment offers (default: every built-in one). A built-in app is offered
  // only while its module is here, so a subset of builtinConnectors offers only those. Your own
  // apps: a ConnectorModule carrying its definition as `connector`, or a definition with an
  // `http`, `soap`, `odata` or `sql` block alone (no code).
  connectors?: (ConnectorModule | BuiltinModule | ConnectorJson)[];
  // Folders laid out like connectors/ (<id>/connector.json), loaded after it; each needs a module
  // in `connectors` unless it has an `http`, `soap`, `odata` or `sql` block.
  extraConnectorDirs?: string[];
  // Approving the deployment's own multi-tenant Microsoft app instead of registering one: only
  // for a deployment that runs it, since every holder of its secret could reach every tenant
  // that approved it.
  microsoftSharedApp?: { clientId: string; clientSecret: string } | null;
  // What answers Teams triggers and app events (default: Claude Managed Agents and the signed
  // webhook).
  runners?: RunnerFactory[];
  // How long a run its runner did not take (no answer, 5xx, 408, 409, 429) is handed over again,
  // in seconds: 0 to 3600, default 900; 0 never tries again.
  runnerRetrySeconds?: number;
  log?: Pick<Console, "log" | "error">;
  // Reports failures to Sentry (its DSN) as they happen; off when left out.
  errorReporting?: ErrorReporting | null;
};

/** The whole service, checked and wired: fetch serves /agent-connect/*, start runs the Teams
 *  and app events queue workers, watchers and polls. Throws on anything misconfigured. */
export async function createAgentConnect(options: AgentConnectOptions) {
  const log = options.log ?? console;
  setBranding(parseBranding(options.branding));
  const schema = options.schema ?? "agent_connect";
  assertSchemaName(schema);
  const retryWindowMs = (options.runnerRetrySeconds ?? RETRY_WINDOW_MS / 1000) * 1000;
  if (!Number.isInteger(retryWindowMs) || retryWindowMs < 0 || retryWindowMs > MAX_RETRY_WINDOW_MS)
    throw new Error("runnerRetrySeconds must be a whole number of seconds from 0 to 3600");

  const microsoft = microsoftClient({ sharedApp: options.microsoftSharedApp ?? null });
  const modules = new Map<string, BuiltinModule>();
  const inline: Connector[] = [];
  for (const entry of options.connectors ?? builtinConnectors) {
    if (!("create" in entry)) {
      inline.push(parseConnector(entry, `Connector ${entry.id}`));
      continue;
    }
    if (modules.has(entry.id)) throw new Error(`Two connector modules have the id ${entry.id}`);
    modules.set(entry.id, entry);
    if (!entry.connector) continue;
    const connector = parseConnector(entry.connector, `Connector ${entry.id}`);
    if (connector.id !== entry.id)
      throw new Error(`Connector module ${entry.id} carries the definition of ${connector.id}`);
    inline.push(connector);
  }
  const builtin = loadConnectors(CONNECTORS_DIR, (id) => modules.has(id));
  const connectors = [
    ...builtin,
    ...inline,
    ...(options.extraConnectorDirs ?? []).flatMap((dir) => {
      const found = loadConnectors(pathToFileURL(`${dir.replace(/\/+$/, "")}/`));
      if (!found.length) throw new Error(`No connector.json in a subfolder of ${dir}`);
      return found;
    }),
  ];
  for (const id of modules.keys())
    if (!connectors.some((c) => c.id === id))
      throw new Error(
        `Connector module ${id} has no definition: give it as its connector, or its folder in extraConnectorDirs`,
      );
  // An app described entirely in its connector.json runs on a generic module.
  const generic = (connector: Connector): BuiltinModule | undefined =>
    connector.http
      ? httpConnector
      : connector.soap
        ? soapConnector
        : connector.odata
          ? odataConnector
          : connector.sql
            ? sqlConnector
            : undefined;
  // Every app, loaded here or an organization's own, passes the same checks.
  const build = (connector: Connector, module: BuiltinModule): App => {
    const app = { connector, ...module.create({ microsoft, connector }) };
    assertToolsImplemented([app]);
    // Every `live` step check a connector names is one the Teams admin runs.
    for (const step of [...connector.setup, ...connector.afterConnect])
      if (step.live && !(LIVE_CHECKS as readonly string[]).includes(step.live))
        throw new Error(`Connector ${connector.id}: no live check ${step.live}`);
    return app;
  };
  const apps: App[] = connectors.map((connector) => {
    if (connectors.filter((c) => c.id === connector.id).length > 1)
      throw new Error(`Two connectors have the id ${connector.id}`);
    const module = modules.get(connector.id) ?? generic(connector);
    if (!module) throw new Error(`Connector ${connector.id} has no code (ConnectorModule)`);
    return build(connector, module);
  });
  if (!apps.length) throw new Error("No connectors: pass at least one app in connectors");
  const providerClients: Record<string, ProviderApp["client"]> = { microsoft: microsoft.provider };
  // A built-in app left out drops out of its provider's add-agent steps.
  const left = new Set(
    readdirSync(CONNECTORS_DIR).filter(
      (id) => existsSync(new URL(`${id}/connector.json`, CONNECTORS_DIR)) && !modules.has(id),
    ),
  );
  const providers: ProviderApp[] = loadProviders().map((provider) => {
    const client = providerClients[provider.id];
    if (!client) throw new Error(`connectors/${provider.id}/provider.json has no client`);
    const addAgent = provider.addAgent
      .map((step) => ({ ...step, apps: step.apps.filter((id) => !left.has(id)) }))
      .filter((step) => step.apps.length);
    return { provider: { ...provider, addAgent }, client };
  });
  const teamsConnector = apps.find((a) => a.connector.id === "teams")?.connector ?? null;

  const ownsSql = !("unsafe" in options.database);
  const sql =
    "unsafe" in options.database ? options.database : connectDatabase(options.database.url);
  if (options.migrate ?? true) await runMigrations(sql, schema);
  const report = errorReporter(options.errorReporting, { log });
  const store = reportingVendorRefusals(createStore(sql, schema), report);
  const key = new CredentialKey(options.encryptionKey, options.previousEncryptionKey);
  const teamsStore = createTeamsStore(sql, schema);
  const bots = botClient();
  const base = options.publicBaseUrl?.trim().replace(/\/+$/, "");
  const publicUrl = base ? `${base}/agent-connect` : null;
  const runners = new Map(
    (options.runners ?? [claudeManagedAgents(), signedWebhook()])
      .map((make) => make({ teams: teamsStore, key, publicUrl }))
      .map((r) => [r.id, r]),
  );
  const runtime = createTeamsRuntime({
    store,
    teams: teamsStore,
    key,
    auth: botFrameworkAuth(),
    bots,
    runners,
    connector: teamsConnector,
    publicUrl,
    report,
    retryWindowMs,
  });
  const service = createService({
    apps,
    internalApp: (connector) => build(connector, generic(connector)!),
    // Built-in folders stay reserved even when this deployment does not offer them.
    reservedIds: readdirSync(CONNECTORS_DIR, { withFileTypes: true })
      .filter((d) => d.isDirectory())
      .map((d) => d.name),
    providers,
    store,
    key,
    bots: { botForAgent: teamsStore.botForAgent, access: runtime.access },
  });
  const eventsStore = createEventsStore(sql, schema);
  const eventApps = new Map(
    apps.filter((a) => a.connector.events).map((a) => [a.connector.id, a.connector]),
  );
  const eventDeps = {
    store,
    events: eventsStore,
    key,
    connectors: eventApps,
    appCall: service.appCall,
    runners,
    publicUrl,
    report,
    retryWindowMs,
  };
  const events = createEventsRuntime(eventDeps);
  const app = createApp({
    service,
    store,
    adminToken: options.adminToken,
    events: { runtime: events, admin: createEventsAdmin({ ...eventDeps, runtime: events }) },
    teams: {
      runtime,
      admin: createTeamsAdmin({
        store,
        teams: teamsStore,
        key,
        bots,
        claude: claudeClient,
        runners,
        publicUrl,
        microsoft,
      }),
    },
  });

  return {
    fetch: (req: Request) => app.fetch(req),
    connectors: apps.map((a) => a.connector.id),
    start() {
      runtime.start();
      events.start();
      // While rotating the key, move what Teams sealed onto the new one.
      if (options.previousEncryptionKey)
        resealTeamsCredentials(teamsStore, key).then(
          (n) =>
            log.log(
              `[agent-connect] re-sealed ${n.bots} bot credentials, ${n.claudeKeys} Claude keys and ${n.triggerSecrets} trigger secrets`,
            ),
          (error) => log.error("[agent-connect] re-sealing Teams credentials failed", error),
        );
      if (options.previousEncryptionKey)
        resealEventSecrets(eventsStore, key).then(
          (n) => log.log(`[agent-connect] re-sealed ${n} app event secrets`),
          (error) => log.error("[agent-connect] re-sealing app event secrets failed", error),
        );
    },
    async stop() {
      await Promise.all([runtime.stop(), events.stop()]);
      if (ownsSql) await sql.end();
    },
  };
}

export type AgentConnect = Awaited<ReturnType<typeof createAgentConnect>>;
