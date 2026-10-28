import { serve } from "@hono/node-server";
import { builtinConnectors, createAgentConnect } from "./index.js";
import { migrate } from "./migrate.js";
import { assertSchemaName, connectDatabase } from "./store.js";

// The container's entry: reads the environment, builds the service with createAgentConnect
// and listens. `node dist/server.js migrate` only applies pending migrations and exits, for
// installs that run schema changes as their own step (MIGRATE_ON_START=false).

function required(name: string): string {
  const value = process.env[name];
  if (!value) throw new Error(`Missing required env var: ${name}`);
  return value;
}

// The name a setting had before the service had its own; read for one release so an install
// whose environment still sets only the old name keeps working, saying so in its log.
function renamed(name: string, old: string): string | undefined {
  const value = process.env[name]?.trim();
  if (value) return value;
  const legacy = process.env[old]?.trim();
  if (legacy) console.log(`[agent-connect] ${old} is deprecated; set ${name} instead`);
  return legacy || undefined;
}

const schema = process.env.DATABASE_SCHEMA?.trim() || "agent_connect";
assertSchemaName(schema);
const migrateOnStart = process.env.MIGRATE_ON_START?.trim() ?? "true";
if (migrateOnStart !== "true" && migrateOnStart !== "false")
  throw new Error("MIGRATE_ON_START must be true or false");
if (process.argv[2] === "migrate") {
  const sql = connectDatabase(required("DATABASE_URL"));
  await migrate(sql, schema);
  await sql.end();
  process.exit(0);
}

const sharedApp = process.env.MICROSOFT_SHARED_APP?.trim() ?? "";
if (sharedApp !== "" && sharedApp !== "enabled")
  throw new Error("MICROSOFT_SHARED_APP must be unset or enabled");
const sharedAppSetting = (name: string, old: string) => {
  const value = renamed(name, old);
  if (!value) throw new Error(`Missing required env var: ${name}`);
  return value;
};

// Which built-in apps this deployment offers (comma-separated ids); unset offers every one.
// Apps from CONNECTOR_DIRS are always offered.
function builtinApps() {
  const ids = (process.env.BUILTIN_APPS ?? "").split(",").map((id) => id.trim()).filter(Boolean);
  if (!ids.length) return undefined;
  const known = builtinConnectors.map((m) => m.id);
  const unknown = ids.filter((id) => !known.includes(id));
  if (unknown.length)
    throw new Error(`BUILTIN_APPS: no built-in app ${unknown.join(", ")} (known: ${known.join(", ")})`);
  return builtinConnectors.filter((m) => ids.includes(m.id));
}

const service = await createAgentConnect({
  database: { url: required("DATABASE_URL") },
  schema,
  migrate: migrateOnStart === "true",
  encryptionKey: required("CONNECT_ENCRYPTION_KEY"),
  previousEncryptionKey: process.env.CONNECT_ENCRYPTION_KEY_PREVIOUS || undefined,
  adminToken: required("AGENT_CONNECT_ADMIN_TOKEN"),
  publicBaseUrl: renamed("PUBLIC_BASE_URL", "SELF_HOST_PUBLIC_BASE_URL") ?? null,
  branding: process.env.BRANDING?.trim() ? JSON.parse(process.env.BRANDING) : undefined,
  connectors: builtinApps(),
  extraConnectorDirs: (process.env.CONNECTOR_DIRS ?? "").split(":").filter(Boolean),
  microsoftSharedApp:
    sharedApp === "enabled"
      ? {
          clientId: sharedAppSetting(
            "MICROSOFT_SHARED_APP_CLIENT_ID",
            "ARBIUM_MICROSOFT_CLIENT_ID",
          ),
          clientSecret: sharedAppSetting(
            "MICROSOFT_SHARED_APP_CLIENT_SECRET",
            "ARBIUM_MICROSOFT_CLIENT_SECRET",
          ),
        }
      : null,
});
service.start();
process.on("SIGTERM", () => {
  service.stop().finally(() => process.exit(0));
});

const port = Number(process.env.PORT || 3000);
serve({ fetch: service.fetch, port });
console.log(`[agent-connect] listening on :${port} with ${service.connectors.join(", ")}`);
