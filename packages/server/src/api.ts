import { z } from "zod";
import { connectorSchema, providerSchema, setupStep } from "../../connector-kit/src/connector.js";
import { brandingSchema } from "../../connector-kit/src/branding.js";
import { eventTriggerSchema, subscriptionSchema } from "./events/admin.js";
import { triggerSchema } from "./teams/admin.js";

// The admin API's contract: every route once, with its body and the shapes it answers. The
// router (app.ts) parses bodies with these, tests check every JSON answer against them, and
// openapi.ts turns them into openapi.json, from which the client package's types are made.

const name = z.string().trim().min(1).max(200);
const by = z.object({ by: name });
const fields = z.record(z.string(), z.string().max(4000));
const redirectUri = z
  .url()
  .refine(
    (u) => u.startsWith("https://") || /^http:\/\/localhost(:\d+)?\//.test(u),
    "redirectUri must be https",
  )
  .describe("The admin app's own callback page: https, or http://localhost while developing");
const consentFinish = z.object({
  state: z.string().max(4000),
  nonce: z.string().min(1).max(200),
  params: z.record(z.string(), z.string().max(4000)),
  connectedBy: name,
});

const iso = z.string().describe("ISO 8601, UTC");
const config = z.record(z.string(), z.string());

// Shared shapes, each a named schema in openapi.json (`id`).
const named = <T extends z.ZodType>(id: string, schema: T) => schema.meta({ id });

export const schemas = (() => {
  const Organization = named(
    "Organization",
    z.strictObject({ id: z.guid(), name: z.string().nullable(), createdAt: iso }),
  );
  const Check = named(
    "Check",
    z.strictObject({
      label: z.string(),
      values: z.array(z.string()),
      warning: z.string().optional(),
    }),
  );
  const Ok = z.strictObject({ ok: z.literal(true) });
  const Failed = z.strictObject({ ok: z.literal(false), error: z.string() });
  const CertificateInfo = named(
    "CertificateInfo",
    z.strictObject({ thumbprint: z.string(), expiresAt: iso }),
  );
  const Connection = named(
    "Connection",
    z.strictObject({
      id: z.string(),
      organizationId: z.string(),
      connectorId: z.string(),
      providerId: z.string().nullable(),
      capabilities: z.array(z.string()),
      config,
      status: z.enum(["connected", "needs_reconnect"]),
      accountLabel: z.string().nullable(),
      connectedBy: z.string().nullable(),
      connectedAt: iso,
      checkedAt: iso,
    }),
  );
  const ProviderConnection = named(
    "ProviderConnection",
    z.strictObject({
      id: z.string(),
      organizationId: z.string(),
      provider: z.string(),
      status: z.enum(["pending", "connected", "needs_reconnect"]),
      config,
      accountLabel: z.string().nullable(),
      connectedBy: z.string().nullable(),
      connectedAt: iso,
      checkedAt: iso,
      certificate: CertificateInfo.nullable(),
      nextCertificate: CertificateInfo.nullable(),
      apps: z.array(z.string()),
    }),
  );
  const Limits = named(
    "Limits",
    z.record(
      z.string(),
      z.union([z.array(z.string()), z.string(), z.record(z.string(), z.string())]),
    ),
  );
  const Grant = named(
    "Grant",
    z.strictObject({
      agentId: z.string(),
      connectorId: z.string(),
      capabilities: z.array(z.string()),
      limits: Limits,
      updatedAt: iso,
    }),
  );
  const agentFields = {
    id: z.string(),
    organizationId: z.string(),
    name: z.string(),
    status: z.enum(["active", "disabled"]),
    createdBy: z.string().nullable(),
    createdAt: iso,
  };
  const NewAgent = named("NewAgent", z.strictObject(agentFields));
  const Agent = named(
    "Agent",
    z.strictObject({ ...agentFields, lastUsedAt: iso.nullable(), grants: z.array(Grant) }),
  );
  const access = z.enum(["granted", "not_granted", "not_found", "failed"]);
  const SiteAccess = named(
    "SiteAccess",
    z.strictObject({
      site: z.string(),
      status: access,
      name: z.string().nullable(),
      message: z.string(),
    }),
  );
  const LiveCheck = named(
    "LiveCheck",
    z.strictObject({ state: z.enum(["done", "failing", "todo"]), message: z.string() }),
  );
  const botFields = {
    id: z.string(),
    organizationId: z.string(),
    agentId: z.string(),
    botAppId: z.string().nullable(),
    teamsAppId: z.string(),
    tenantId: z.string(),
    displayName: z.string(),
    manifestVersion: z.number(),
    status: z.enum(["awaiting_setup", "connected", "needs_reconnect"]),
    createdBy: z.string().nullable(),
    createdAt: iso,
    lastRefusedAt: iso.nullable(),
    lastRefusedReason: z.string().nullable(),
    certificate: CertificateInfo.nullable(),
    nextCertificate: CertificateInfo.nullable(),
    madeByService: z.boolean(),
    publishedVersion: z.string().nullable(),
    packageVersion: z.string(),
  };
  const TeamsBotRecord = named("TeamsBotRecord", z.strictObject(botFields));
  const TeamsPlace = named(
    "TeamsPlace",
    z.strictObject({
      kind: z.enum(["personal", "groupChat", "team"]),
      name: z.string().nullable(),
      id: z.string(),
      conversationId: z.string(),
    }),
  );
  const TeamsBot = named(
    "TeamsBot",
    z.strictObject({
      ...botFields,
      messagingEndpoint: z.string().nullable(),
      mcpUrl: z.string().nullable(),
      installedIn: z.array(TeamsPlace),
    }),
  );
  const TeamsTrigger = named(
    "TeamsTrigger",
    z.strictObject({
      agentId: z.string(),
      enabled: z.boolean(),
      runner: z.string(),
      settings: z.record(z.string(), z.unknown()),
      hasSigningSecret: z.boolean(),
      // Set for the Claude Managed Agents runner, null for others.
      managedAgentId: z.string().nullable(),
      environmentId: z.string().nullable(),
      vaultIds: z.array(z.string()),
      budgetCents: z.number(),
      maxRunsPerHour: z.number(),
      maxCentsPerDay: z.number(),
      maxConcurrent: z.number(),
      replyMode: z.enum(["mentions", "all"]),
      allowedTeams: z.array(z.string()),
      allowedPeople: z.array(z.string()),
      anyoneInOrganization: z.boolean(),
      updatedBy: z.string().nullable(),
      updatedAt: iso,
    }),
  );
  const TeamsRun = named(
    "TeamsRun",
    z.strictObject({
      id: z.string(),
      agentId: z.string(),
      botId: z.string(),
      conversationId: z.string(),
      threadId: z.string(),
      triggerActivityId: z.string(),
      runner: z.string(),
      sessionId: z.string().nullable(),
      state: z.enum(["starting", "running", "idle", "ended", "failed", "no_reply"]),
      error: z.string().nullable(),
      note: z.string().nullable(),
      requestedBy: z.string().nullable(),
      replyCount: z.number(),
      repliedAt: iso.nullable(),
      costCents: z.number().nullable(),
      turns: z.number(),
      startedAt: iso,
      finishedAt: iso.nullable(),
      replyTokenAt: iso,
    }),
  );
  const EventSubscription = named(
    "EventSubscription",
    z.strictObject({
      id: z.string(),
      connectorId: z.string(),
      source: z.enum(["webhook", "poll"]),
      filter: z.string(),
      intervalSeconds: z.number(),
      cursor: z.string().nullable(),
      polledAt: iso.nullable(),
      lastError: z.string().nullable(),
      registered: z.string().nullable(),
      // Where the app sends a webhook's events; null for a poll.
      url: z.string().nullable(),
      createdBy: z.string(),
      createdAt: iso,
    }),
  );
  const EventTrigger = named(
    "EventTrigger",
    z.strictObject({
      id: z.string(),
      agentId: z.string(),
      connectorId: z.string(),
      enabled: z.boolean(),
      eventTypes: z.array(z.string()),
      conditions: z.array(
        z.union([
          z.strictObject({ field: z.string(), in: z.array(z.string()) }),
          z.strictObject({ field: z.string(), notIn: z.array(z.string()) }),
        ]),
      ),
      prompt: z.string(),
      runner: z.string(),
      settings: z.record(z.string(), z.unknown()),
      hasSigningSecret: z.boolean(),
      managedAgentId: z.string().nullable(),
      environmentId: z.string().nullable(),
      vaultIds: z.array(z.string()),
      budgetCents: z.number(),
      maxRunsPerHour: z.number(),
      maxCentsPerDay: z.number(),
      maxConcurrent: z.number(),
      updatedBy: z.string().nullable(),
      updatedAt: iso,
    }),
  );
  const EventRun = named(
    "EventRun",
    z.strictObject({
      id: z.string(),
      agentId: z.string(),
      triggerId: z.string(),
      connectorId: z.string(),
      eventType: z.string(),
      deliveryId: z.string(),
      runner: z.string(),
      sessionId: z.string().nullable(),
      state: z.enum(["starting", "running", "ended", "failed"]),
      error: z.string().nullable(),
      costCents: z.number().nullable(),
      startedAt: iso,
      finishedAt: iso.nullable(),
    }),
  );
  const keySet = z.strictObject({
    set: z.literal(true),
    ending: z.string(),
    by: z.string().nullable(),
    at: iso,
  });
  const ClaudeKey = named(
    "ClaudeKey",
    z.union([z.strictObject({ set: z.literal(false) }), keySet]),
  );
  const ClaudeAgent = named("ClaudeAgent", z.strictObject({ id: z.string(), name: z.string() }));
  const ClaudeKeySaved = named("ClaudeKeySaved", keySet.extend({ agents: z.array(ClaudeAgent) }));
  const SignIn = named("SignIn", z.strictObject({ url: z.string(), nonce: z.string() }));
  const ConnectResult = named(
    "ConnectResult",
    z.strictObject({ connection: Connection, checks: z.array(Check) }),
  );
  return {
    Error: named("Error", z.strictObject({ error: z.string() })),
    Organization,
    Connector: named("Connector", connectorSchema),
    Branding: named("Branding", brandingSchema),
    Provider: named("Provider", providerSchema.extend({ approval: z.boolean() })),
    ProviderConnection,
    Connection,
    Check,
    CheckOutcome: named(
      "CheckOutcome",
      z.union([z.strictObject({ ok: z.literal(true), checks: z.array(Check) }), Failed]),
    ),
    SignInCheck: named("SignInCheck", z.union([Ok, Failed])),
    ConnectResult,
    SignIn,
    SiteAccess,
    AgentAccess: named(
      "AgentAccess",
      z.strictObject({ value: z.string(), status: access, message: z.string() }),
    ),
    Limits,
    Grant,
    NewAgent,
    Agent,
    AuditRow: named(
      "AuditRow",
      z.strictObject({
        at: iso,
        agentId: z.string().nullable(),
        agentName: z.string().nullable(),
        connectorId: z.string(),
        tool: z.string(),
        target: z.string().nullable(),
        decision: z.enum(["allowed", "denied", "failed"]),
        reason: z.string().nullable(),
        vendorStatus: z.number().nullable(),
      }),
    ),
    Confirmed: named("Confirmed", z.strictObject({ step: z.string(), by: z.string(), at: iso })),
    LiveCheck,
    TeamsBotRecord,
    TeamsPlace,
    TeamsBot,
    AddStep: named(
      "AddStep",
      z.strictObject({
        id: z.string(),
        label: z.string(),
        state: z.enum(["done", "failed", "waiting", "todo"]),
        message: z.string().optional(),
      }),
    ),
    InstallTarget: named(
      "InstallTarget",
      z.strictObject({ id: z.string(), name: z.string(), detail: z.string().nullable() }),
    ),
    TeamsTrigger,
    TeamsRun,
    EventSubscription,
    EventSubscribed: named(
      "EventSubscribed",
      z.strictObject({
        subscription: EventSubscription,
        // Shown once: what the app signs deliveries with. Null for a poll.
        secret: z.string().nullable(),
        // Where to paste the address and secret in the app ({url} filled in).
        setup: z.array(setupStep),
        // Registering through the app's API, when asked.
        registration: z.strictObject({ ok: z.boolean(), message: z.string() }).nullable(),
      }),
    ),
    EventTrigger,
    EventRun,
    ClaudeKey,
    ClaudeKeySaved,
    ClaudeAgent,
    Ok,
  };
})();

const s = schemas;

/** A route's answer: a JSON shape, a file of this content type, or no body (null). */
export type Answer = z.ZodType | { file: string } | null;

export type RouteSpec = {
  method: "GET" | "POST" | "PUT" | "PATCH" | "DELETE";
  /** Parameters: {org} and {agent} are uuids, {app}, {provider} and {step} ids. */
  path: `/${string}`;
  summary: string;
  query?: Record<string, string>;
  body?: z.ZodType;
  responses: Record<number, Answer>;
  /** Served only where Teams is set up. */
  teams?: true;
  /** Served only where app events are set up. */
  events?: true;
};

export const PARAMS = {
  org: "[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}",
  agent: "[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}",
  app: "[a-z0-9-]+",
  provider: "[a-z0-9-]+",
  step: "[a-z0-9-]+",
  subscription: "[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}",
} as const;

const notFound = { 404: s.Error } as const;
const ORG = "/orgs/{org}";
const APP = `${ORG}/connections/{app}`;
const PROVIDER = `${ORG}/providers/{provider}`;
const AGENT = `${ORG}/agents/{agent}`;
const BOT = `${AGENT}/teams-bot`;

/** Every admin route, by the client method that calls it. */
export const routes = {
  listOrganizations: {
    method: "GET",
    path: "/orgs",
    summary: "The organizations",
    responses: { 200: z.array(s.Organization) },
  },
  createOrganization: {
    method: "POST",
    path: "/orgs",
    summary: "Makes an organization; its id is made here",
    body: z.object({ name }),
    responses: { 201: s.Organization },
  },
  getOrganization: {
    method: "GET",
    path: ORG,
    summary: "One organization",
    responses: { 200: s.Organization, ...notFound },
  },
  ensureOrganization: {
    method: "PUT",
    path: ORG,
    summary: "Makes the organization with this id if missing, and names it when a name is given",
    body: z.object({ name: name.optional() }),
    responses: { 200: s.Organization },
  },
  branding: {
    method: "GET",
    path: "/branding",
    summary: "The deployment's names and logo, for a UI to show",
    responses: { 200: s.Branding },
  },
  connectors: {
    method: "GET",
    path: "/connectors",
    summary: "The apps agents can use",
    responses: { 200: z.array(s.Connector) },
  },
  providers: {
    method: "GET",
    path: "/providers",
    summary: "The sign-ins several apps share",
    responses: { 200: z.array(s.Provider) },
  },
  providerConnections: {
    method: "GET",
    path: `${ORG}/providers`,
    summary: "The organization's provider connections",
    responses: { 200: z.array(s.ProviderConnection) },
  },
  connectProvider: {
    method: "PUT",
    path: PROVIDER,
    summary: "Connects a provider, proving the sign-in first",
    body: z.object({ fields, connectedBy: name }),
    responses: { 200: s.ProviderConnection },
  },
  disconnectProvider: {
    method: "DELETE",
    path: PROVIDER,
    summary: "Disconnects a provider and every app signing in through it",
    body: by,
    responses: { 204: null, ...notFound },
  },
  checkProvider: {
    method: "POST",
    path: `${PROVIDER}/check`,
    summary: "Signs in again",
    responses: { 200: s.SignInCheck, ...notFound },
  },
  newCertificate: {
    method: "POST",
    path: `${PROVIDER}/certificate`,
    summary: "Before connecting, the certificate to upload; after, one to rotate to",
    body: by,
    responses: { 200: s.ProviderConnection },
  },
  activateCertificate: {
    method: "POST",
    path: `${PROVIDER}/certificate/activate`,
    summary: "Signs in with the new certificate and, if that works, switches to it",
    body: by,
    responses: { 200: s.ProviderConnection },
  },
  setProviderSettings: {
    method: "PUT",
    path: `${PROVIDER}/settings`,
    summary: "Values the provider's apps ask for in their steps",
    body: z.object({ values: z.record(z.string(), z.string().max(400)), by: name }),
    responses: { 200: z.strictObject({ settings: config }) },
  },
  confirmed: {
    method: "GET",
    path: `${ORG}/confirmed`,
    summary: "Setup steps an admin marked done",
    responses: { 200: z.array(s.Confirmed) },
  },
  confirm: {
    method: "PUT",
    path: `${ORG}/confirmed/{step}`,
    summary: "Marks a setup step done",
    body: by,
    responses: { 204: null },
  },
  unconfirm: {
    method: "DELETE",
    path: `${ORG}/confirmed/{step}`,
    summary: "Marks a setup step not done",
    body: by,
    responses: { 204: null },
  },
  connections: {
    method: "GET",
    path: `${ORG}/connections`,
    summary: "The organization's apps that are on",
    responses: { 200: z.array(s.Connection) },
  },
  connect: {
    method: "PUT",
    path: APP,
    summary: "Turns an app on, checking it before saving",
    body: z.object({ capabilities: z.array(z.string()), fields, connectedBy: name }),
    responses: { 200: s.ConnectResult },
  },
  setCeiling: {
    method: "PATCH",
    path: APP,
    summary: "The most any agent may do in the app",
    body: z.object({ capabilities: z.array(z.string()), by: name }),
    responses: { 204: null, ...notFound },
  },
  disconnect: {
    method: "DELETE",
    path: APP,
    summary: "Turns an app off",
    body: by,
    responses: { 204: null, ...notFound },
  },
  startConsent: {
    method: "POST",
    path: `${APP}/consent`,
    summary: "Where the admin approves the shared app",
    body: z.object({ capabilities: z.array(z.string()), redirectUri }),
    responses: { 200: s.SignIn },
  },
  finishConsent: {
    method: "POST",
    path: `${APP}/consent/callback`,
    summary:
      "The approval came back: `redirect` is the sign-in that proves the organization, " +
      "`provider` the provider connection when the app still needs its own settings",
    body: consentFinish,
    responses: {
      200: z.union([
        z.strictObject({ redirect: z.string() }),
        z.strictObject({ provider: s.ProviderConnection }),
        s.ConnectResult,
      ]),
    },
  },
  check: {
    method: "POST",
    path: `${APP}/check`,
    summary:
      "Checks the app. Recorded (the default) marks a failing app as needing reconnecting " +
      "and needs `by`; record=false only reads",
    query: { record: "false: read only" },
    body: by.optional(),
    responses: { 200: s.CheckOutcome, ...notFound },
  },
  setSites: {
    method: "PUT",
    path: `${APP}/sites`,
    summary: "The sites added to the app, the whole list at once",
    body: z.object({ sites: z.array(z.string().max(500)).max(500), by: name }),
    responses: { 200: z.strictObject({ sites: z.array(z.string()) }) },
  },
  checkSites: {
    method: "POST",
    path: `${APP}/sites/check`,
    summary: "Whether the app can open each added site, or just `site`",
    body: z.object({ site: z.string().max(500).optional() }),
    responses: { 200: z.strictObject({ sites: z.array(s.SiteAccess) }) },
  },
  agentAccess: {
    method: "POST",
    path: `${APP}/agent-access`,
    summary: "Whether the app can reach one agent's own address yet",
    body: z.object({ value: z.string().min(1).max(320), by: name }),
    responses: { 200: s.AgentAccess },
  },
  agents: {
    method: "GET",
    path: `${ORG}/agents`,
    summary: "The organization's agents (API keys) and their grants",
    responses: { 200: z.array(s.Agent) },
  },
  createAgent: {
    method: "POST",
    path: `${ORG}/agents`,
    summary: "Makes an agent; its key is returned this once",
    body: z.object({ name, createdBy: name }),
    responses: { 201: z.strictObject({ agent: s.NewAgent, key: z.string() }), 409: s.Error },
  },
  setAgentStatus: {
    method: "PATCH",
    path: AGENT,
    summary: "Turns an agent's key on or off",
    body: z.object({ status: z.enum(["active", "disabled"]), by: name }),
    responses: { 204: null, ...notFound },
  },
  deleteAgent: {
    method: "DELETE",
    path: AGENT,
    summary: "Deletes an agent",
    body: by,
    responses: { 204: null, ...notFound },
  },
  newKey: {
    method: "POST",
    path: `${AGENT}/key`,
    summary: "Replaces an agent's key; the new one is returned this once",
    body: by,
    responses: { 200: z.strictObject({ key: z.string() }), ...notFound },
  },
  setGrant: {
    method: "PUT",
    path: `${AGENT}/grants/{app}`,
    summary: "What the agent may do in one app",
    body: z.object({ capabilities: z.array(z.string()), limits: z.unknown().optional(), by: name }),
    responses: { 204: null, ...notFound },
  },
  removeGrant: {
    method: "DELETE",
    path: `${AGENT}/grants/{app}`,
    summary: "Takes one app away from the agent",
    body: by,
    responses: { 204: null, ...notFound },
  },
  audit: {
    method: "GET",
    path: `${ORG}/audit`,
    summary: "Recent requests and admin changes, newest first",
    query: { limit: "1 to 500, default 50", agent: "an agent id", app: "an app id" },
    responses: { 200: z.array(s.AuditRow) },
  },
  teamsBots: {
    method: "GET",
    path: `${ORG}/teams/bots`,
    summary: "Each agent's Teams bot",
    responses: { 200: z.array(s.TeamsBot) },
    teams: true,
  },
  teamsSetupState: {
    method: "GET",
    path: `${ORG}/teams/setup`,
    summary: "The Teams page's set-up-once steps, checked live",
    responses: { 200: z.record(z.string(), s.LiveCheck) },
    teams: true,
  },
  finishTeamsPublish: {
    method: "POST",
    path: `${ORG}/teams/publish/callback`,
    summary: "The publishing sign-in came back: publishes the bot's package",
    body: consentFinish.omit({ connectedBy: true }).extend({ by: name }),
    responses: {
      200: z.strictObject({ agentId: z.string(), version: z.string(), changed: z.boolean() }),
    },
    teams: true,
  },
  createTeamsBot: {
    method: "POST",
    path: BOT,
    summary: "Makes the agent's bot's key pair, for the setup script",
    body: z.object({ displayName: z.string(), by: name }),
    responses: { 200: s.TeamsBotRecord },
    teams: true,
  },
  connectTeamsBot: {
    method: "PUT",
    path: BOT,
    summary: "Proves the application ID the setup script printed by signing in as it",
    body: z.object({ botAppId: z.string().max(100), by: name }),
    responses: { 200: s.TeamsBotRecord },
    teams: true,
  },
  removeTeamsBot: {
    method: "DELETE",
    path: BOT,
    summary:
      "Forgets the bot. What Add to Teams made is deleted (`failed`: what could not be); " +
      "`removal` is the commands for the rest",
    body: by,
    responses: {
      200: z.strictObject({ removal: z.string(), failed: z.array(z.string()) }),
      ...notFound,
    },
    teams: true,
  },
  addToTeams: {
    method: "POST",
    path: `${BOT}/add`,
    summary: "Makes the agent's bot in the tenant, each part only if missing",
    body: z.object({ displayName: z.string().max(100), by: name }),
    responses: { 200: z.strictObject({ steps: z.array(s.AddStep), bot: s.TeamsBotRecord }) },
    teams: true,
  },
  startTeamsPublish: {
    method: "POST",
    path: `${BOT}/publish`,
    summary: "Where the administrator signs in to publish the bot",
    body: z.object({ redirectUri }),
    responses: { 200: s.SignIn },
    teams: true,
  },
  teamsTargets: {
    method: "GET",
    path: `${BOT}/targets`,
    summary: "Teams or people to add the bot to, matching `q`",
    query: { kind: "team or person", q: "what the admin typed" },
    responses: { 200: z.array(s.InstallTarget) },
    teams: true,
  },
  installTeamsBot: {
    method: "POST",
    path: `${BOT}/install`,
    summary: "Adds the published bot to a team, a group chat or a person",
    body: z.object({
      kind: z.enum(["team", "chat", "person"]),
      id: z.string().max(2000),
      by: name,
    }),
    responses: { 200: z.strictObject({ result: z.enum(["added", "already"]) }) },
    teams: true,
  },
  teamsKit: {
    method: "GET",
    path: `${BOT}/kit`,
    summary: "The bot's setup script",
    responses: { 200: { file: "text/plain" } },
    teams: true,
  },
  checkTeamsBot: {
    method: "POST",
    path: `${BOT}/check`,
    summary: "Signs in as the bot again",
    responses: { 200: s.SignInCheck },
    teams: true,
  },
  teamsTest: {
    method: "POST",
    path: `${BOT}/test`,
    summary: "A test message from the bot in one place it was added to",
    body: z.object({ conversationId: z.string().max(400), by: name }),
    responses: { 200: s.Ok },
    teams: true,
  },
  teamsCertificate: {
    method: "GET",
    path: `${BOT}/certificate`,
    summary: "The bot certificate's public half",
    responses: { 200: { file: "application/x-pem-file" } },
    teams: true,
  },
  teamsPackage: {
    method: "GET",
    path: `${BOT}/package`,
    summary: "The bot's Teams app package",
    responses: { 200: { file: "application/zip" } },
    teams: true,
  },
  newTeamsCertificate: {
    method: "POST",
    path: `${BOT}/certificate`,
    summary: "A certificate for the bot to rotate to",
    body: by,
    responses: {
      200: z.strictObject({
        bot: s.TeamsBotRecord,
        command: z.string().nullable(),
        certificate: z.string(),
      }),
    },
    teams: true,
  },
  activateTeamsCertificate: {
    method: "POST",
    path: `${BOT}/certificate/activate`,
    summary: "Signs in with the bot's new certificate and, if that works, switches to it",
    body: by,
    responses: { 200: s.TeamsBotRecord },
    teams: true,
  },
  teamsTrigger: {
    method: "GET",
    path: `${AGENT}/teams-trigger`,
    summary: "How the agent answers in Teams",
    responses: { 200: s.TeamsTrigger, ...notFound },
    teams: true,
  },
  setTeamsTrigger: {
    method: "PUT",
    path: `${AGENT}/teams-trigger`,
    summary: "Sets how the agent answers in Teams, checked by its runner first",
    body: triggerSchema,
    responses: { 200: s.TeamsTrigger },
    teams: true,
  },
  teamsRuns: {
    method: "GET",
    path: `${AGENT}/teams-runs`,
    summary: "The agent's recent answers in Teams, newest first",
    query: { limit: "1 to 200, default 20" },
    responses: { 200: z.array(s.TeamsRun) },
    teams: true,
  },
  eventSubscriptions: {
    method: "GET",
    path: `${ORG}/event-subscriptions`,
    summary: "The apps whose events reach this organization, by webhook or by polling",
    responses: { 200: z.array(s.EventSubscription) },
    events: true,
  },
  subscribeEvents: {
    method: "POST",
    path: `${ORG}/event-subscriptions`,
    summary:
      "Starts receiving an app's events: a webhook (answers its address and secret, shown " +
      "once) or a poll on an interval",
    body: subscriptionSchema,
    responses: { 200: s.EventSubscribed },
    events: true,
  },
  rotateEventSecret: {
    method: "POST",
    path: `${ORG}/event-subscriptions/{subscription}/secret`,
    summary: "A new secret for a webhook, shown once; the old one stops working at once",
    body: by,
    responses: {
      200: z.strictObject({ subscription: s.EventSubscription, secret: z.string() }),
      ...notFound,
    },
    events: true,
  },
  unsubscribeEvents: {
    method: "DELETE",
    path: `${ORG}/event-subscriptions/{subscription}`,
    summary: "Stops receiving an app's events (remove the webhook in the app too)",
    body: by,
    responses: { 204: null, ...notFound },
    events: true,
  },
  eventTriggers: {
    method: "GET",
    path: `${AGENT}/event-triggers`,
    summary: "The app events that start the agent",
    responses: { 200: z.array(s.EventTrigger) },
    events: true,
  },
  setEventTrigger: {
    method: "PUT",
    path: `${AGENT}/event-triggers/{app}`,
    summary: "Sets which of an app's events start the agent and how, checked by its runner first",
    body: eventTriggerSchema,
    responses: { 200: s.EventTrigger },
    events: true,
  },
  removeEventTrigger: {
    method: "DELETE",
    path: `${AGENT}/event-triggers/{app}`,
    summary: "Stops an app's events starting the agent, stopping its live runs",
    body: by,
    responses: { 204: null, ...notFound },
    events: true,
  },
  eventRuns: {
    method: "GET",
    path: `${AGENT}/event-runs`,
    summary: "The agent's recent runs started by app events, newest first",
    query: { limit: "1 to 200, default 20" },
    responses: { 200: z.array(s.EventRun) },
    events: true,
  },
  claudeKey: {
    method: "GET",
    path: `${ORG}/claude-key`,
    summary: "The organization's Claude API key, by its last four characters",
    responses: { 200: s.ClaudeKey },
    teams: true,
  },
  setClaudeKey: {
    method: "PUT",
    path: `${ORG}/claude-key`,
    summary: "Saves the organization's Claude API key, proven by listing its agents",
    body: z.object({ apiKey: z.string().max(400), by: name }),
    responses: { 200: s.ClaudeKeySaved },
    teams: true,
  },
  removeClaudeKey: {
    method: "DELETE",
    path: `${ORG}/claude-key`,
    summary: "Removes the organization's Claude API key",
    body: by,
    responses: { 204: null, ...notFound },
    teams: true,
  },
  claudeAgents: {
    method: "GET",
    path: `${ORG}/claude-key/agents`,
    summary: "The Claude Managed Agents the organization's key can use",
    responses: { 200: z.array(s.ClaudeAgent) },
    teams: true,
  },
} as const satisfies Record<string, RouteSpec>;

export type Routes = typeof routes;
