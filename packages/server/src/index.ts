import { pathToFileURL } from "node:url";
import { createApp } from "./app.js";
import {
  parseBranding,
  setBranding,
  type BrandingInput,
} from "../../connector-kit/src/branding.js";
import { builtinConnectors, type ConnectorModule } from "./builtin.js";
import httpConnector from "../../../connectors/http/index.js";
import { loadConnectors, loadProviders } from "../../connector-kit/src/connector.js";
import { migrate as runMigrations } from "./migrate.js";
import { CredentialKey } from "./secrets.js";
import { assertToolsImplemented, createService, type App, type ProviderApp } from "./service.js";
import { assertSchemaName, connectDatabase, createStore, type Sql } from "./store.js";
import { createTeamsAdmin, resealTeamsCredentials } from "./teams/admin.js";
import { botClient } from "./teams/bot.js";
import { claudeClient } from "./teams/claude.js";
import { botFrameworkAuth } from "./teams/jwt.js";
import { LIVE_CHECKS } from "./teams/provision.js";
import { claudeManagedAgents, signedWebhook, type RunnerFactory } from "./teams/runner.js";
import { createTeamsRuntime } from "./teams/runtime.js";
import { createTeamsStore } from "./teams/store.js";
import { microsoftClient } from "../../../connectors/microsoft/client.js";

export type { BrandingInput } from "../../connector-kit/src/branding.js";
export type { ConnectorModule } from "./builtin.js";
export { builtinConnectors } from "./builtin.js";
export type { RunInput, RunnerFactory, TriggerRunner } from "./teams/runner.js";
export { claudeManagedAgents, signature, signedWebhook } from "./teams/runner.js";

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
  // The code behind each connector JSON (default: the built-in ones).
  connectors?: ConnectorModule[];
  // Folders laid out like connectors/ (<id>/connector.json), loaded after it; each needs a module.
  extraConnectorDirs?: string[];
  // Approving the deployment's own multi-tenant Microsoft app instead of registering one: only
  // for a deployment that runs it, since every holder of its secret could reach every tenant
  // that approved it.
  microsoftSharedApp?: { clientId: string; clientSecret: string } | null;
  // What answers Teams triggers (default: Claude Managed Agents and the signed webhook).
  runners?: RunnerFactory[];
  log?: Pick<Console, "log" | "error">;
};

/** The whole service, checked and wired: fetch serves /agent-connect/*, start runs the Teams
 *  queue worker and watcher. Throws on anything misconfigured. */
export async function createAgentConnect(options: AgentConnectOptions) {
  const log = options.log ?? console;
  setBranding(parseBranding(options.branding));
  const schema = options.schema ?? "agent_connect";
  assertSchemaName(schema);

  const microsoft = microsoftClient({ sharedApp: options.microsoftSharedApp ?? null });
  const modules = new Map((options.connectors ?? builtinConnectors).map((m) => [m.id, m]));
  const connectors = [
    ...loadConnectors(),
    ...(options.extraConnectorDirs ?? []).flatMap((dir) => {
      const found = loadConnectors(pathToFileURL(`${dir.replace(/\/+$/, "")}/`));
      if (!found.length) throw new Error(`No connector.json in a subfolder of ${dir}`);
      return found;
    }),
  ];
  const apps: App[] = connectors.map((connector) => {
    if (connectors.filter((c) => c.id === connector.id).length > 1)
      throw new Error(`Two connectors have the id ${connector.id}`);
    // An app described entirely in its connector.json runs on the generic http module.
    const module = modules.get(connector.id) ?? (connector.http ? httpConnector : undefined);
    if (!module) throw new Error(`Connector ${connector.id} has no code (ConnectorModule)`);
    return { connector, ...module.create({ microsoft, connector }) };
  });
  assertToolsImplemented(apps);
  // Every `live` step check a connector names is one the Teams admin runs.
  for (const { connector } of apps)
    for (const step of [...connector.setup, ...connector.afterConnect])
      if (step.live && !(LIVE_CHECKS as readonly string[]).includes(step.live))
        throw new Error(`Connector ${connector.id}: no live check ${step.live}`);
  const providerClients: Record<string, ProviderApp["client"]> = { microsoft: microsoft.provider };
  const providers: ProviderApp[] = loadProviders().map((provider) => {
    const client = providerClients[provider.id];
    if (!client) throw new Error(`connectors/${provider.id}/provider.json has no client`);
    return { provider, client };
  });
  const teamsConnector = apps.find((a) => a.connector.id === "teams")?.connector;
  if (!teamsConnector) throw new Error("The teams connector is missing");

  const ownsSql = !("unsafe" in options.database);
  const sql =
    "unsafe" in options.database ? options.database : connectDatabase(options.database.url);
  if (options.migrate ?? true) await runMigrations(sql, schema);
  const store = createStore(sql, schema);
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
  });
  const service = createService({
    apps,
    providers,
    store,
    key,
    bots: { botForAgent: teamsStore.botForAgent, access: runtime.access },
  });
  const app = createApp({
    service,
    store,
    adminToken: options.adminToken,
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
      // While rotating the key, move what Teams sealed onto the new one.
      if (options.previousEncryptionKey)
        resealTeamsCredentials(teamsStore, key).then(
          (n) =>
            log.log(
              `[agent-connect] re-sealed ${n.bots} bot credentials, ${n.claudeKeys} Claude keys and ${n.triggerSecrets} trigger secrets`,
            ),
          (error) => log.error("[agent-connect] re-sealing Teams credentials failed", error),
        );
    },
    async stop() {
      await runtime.stop();
      if (ownsSql) await sql.end();
    },
  };
}

export type AgentConnect = Awaited<ReturnType<typeof createAgentConnect>>;
