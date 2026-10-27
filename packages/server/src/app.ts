import { createMcpHandler, McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";
import { PARAMS, routes, type RouteSpec, type Routes } from "./api.js";
import { brand } from "../../connector-kit/src/branding.js";
import { openapi } from "./openapi.js";
import { hashAgentKey, newAgentKey, tokensEqual } from "./secrets.js";
import { InvalidInput, type Service } from "./service.js";
import type { Store } from "./store.js";
import { argumentsTarget, type ToolContext, type Toolset } from "../../connector-kit/src/common.js";
import type { TeamsAdmin } from "./teams/admin.js";
import { readCapped, type TeamsRuntime } from "./teams/runtime.js";
import { ConnectError } from "../../connector-kit/src/types.js";

const BASE = "/agent-connect";

type Session = { toolsets: { toolset: Toolset; ctx: ToolContext; allowed: Set<string> }[] };

// The MCP SDK's own body limit (uploads are sized to fit it), and the most messages one
// request may batch: each refused call writes an audit row.
const MAX_MCP_BODY = 4 * 1024 * 1024;
const MAX_BATCH = 20;

// Messages in a request body (single message or batch); unparsable JSON is left to the SDK.
function messages(text: string): unknown[] {
  let body: unknown;
  try {
    body = JSON.parse(text);
  } catch {
    return [];
  }
  return Array.isArray(body) ? body : [body];
}

// Tool names in tools/call messages.
function calledTools(sent: unknown[]): { name: string; args: unknown }[] {
  return sent.flatMap((m: any) =>
    m?.method === "tools/call" && typeof m.params?.name === "string"
      ? [{ name: m.params.name, args: m.params.arguments }]
      : [],
  );
}

// The SDK checks a call's arguments against the tool's schema before the tool runs, so a
// rejected call never reaches audited(). Each schema is wrapped to record that as denied.
function auditInvalidInput(server: McpServer, ctx: ToolContext): McpServer {
  const register = server.registerTool.bind(server);
  return Object.assign(Object.create(server) as McpServer, {
    registerTool: ((name: string, config: { inputSchema?: z.ZodType }, handler: unknown) => {
      const schema = config.inputSchema;
      if (!schema) return register(name, config as never, handler as never);
      const audited = schema.clone();
      const std = schema["~standard"];
      Object.defineProperty(audited, "~standard", {
        value: {
          ...std,
          validate: async (value: unknown) => {
            const result = await std.validate(value);
            if (result.issues?.length)
              await ctx.audit({
                tool: name,
                target: argumentsTarget(value),
                decision: "denied",
                reason: `Invalid arguments: ${result.issues
                  .map((i) => `${i.path?.join(".") || "input"}: ${i.message}`)
                  .join("; ")}`.slice(0, 500),
                vendorStatus: null,
              });
            return result;
          },
        },
      });
      return register(name, { ...config, inputSchema: audited } as never, handler as never);
    }) as McpServer["registerTool"],
  });
}

const reply = (status: number, body: unknown, headers: Record<string, string> = {}) =>
  new Response(body === null ? null : JSON.stringify(body), {
    status,
    headers: body === null ? headers : { "content-type": "application/json", ...headers },
  });

// The body a handler reads, the answers it may give, and its handler (api.ts declares each).
type BodyOf<S> = S extends { body: infer B extends z.ZodType } ? z.output<B> : undefined;
type Payload<A> = A extends z.ZodType
  ? z.input<A>
  : A extends { file: string }
    ? { data: string | Uint8Array; filename?: string }
    : null;
type Reply<S extends RouteSpec> = {
  [K in keyof S["responses"] & number]: { status: K; body: Payload<S["responses"][K]> };
}[keyof S["responses"] & number];
type Handler<S extends RouteSpec> = (input: {
  params: string[];
  body: BodyOf<S>;
  url: URL;
}) => Promise<Reply<S>>;
type Handlers<Teams extends boolean> = {
  [
    K in keyof Routes as Routes[K] extends { teams: true }
      ? Teams extends true
        ? K
        : never
      : Teams extends true
        ? never
        : K
  ]: Handler<Routes[K]>;
};

// Literal answers, typed to the route's declared ones.
const ok = <T>(body: T) => ({ status: 200 as const, body });
const none = { status: 204 as const, body: null };
const missing = { status: 404 as const, body: { error: "Not found" } };
const found = (yes: boolean) => (yes ? none : missing);
const some = <T>(value: T | null) => (value === null ? missing : ok(value));

const UUID = PARAMS.org;
// The admin API's prefixes. /admin/v1 is the contract (openapi.json); the unversioned
// /admin is deprecated, kept for one release so an admin app of the previous release keeps
// working during an upgrade, and creates an organization it does not know (v1 answers 404).
const ADMIN = `${BASE}/admin`;

// The deprecated prefix also answers madeByService under its old name, at any depth. Not in
// openapi.json.
function legacyFields(body: unknown): unknown {
  if (Array.isArray(body)) return body.map(legacyFields);
  if (!body || Object.getPrototypeOf(body) !== Object.prototype) return body;
  const out = Object.fromEntries(Object.entries(body).map(([k, v]) => [k, legacyFields(v)]));
  return "madeByService" in out ? { ...out, madeByArbium: out.madeByService } : out;
}

export function createApp(deps: {
  service: Service;
  store: Store;
  adminToken: string;
  teams?: { admin: TeamsAdmin; runtime: TeamsRuntime };
  /** Tests: check every JSON answer against its declared shape, failing loudly. */
  validateResponses?: boolean;
}) {
  const { service, store } = deps;
  if (deps.adminToken.length < 32)
    throw new Error("AGENT_CONNECT_ADMIN_TOKEN must be at least 32 characters");

  const mcp = createMcpHandler(
    ({ authInfo }) => {
      const server = new McpServer(
        { name: `${brand().resourcePrefix}-agent-connect`, version: "0.1.0" },
        // An agent with no grants still gets an empty tool list rather than an error.
        { capabilities: { tools: {} } },
      );
      const session = authInfo?.extra?.session as Session | undefined;
      if (!session) throw new Error("MCP request reached the handler without an agent");
      for (const { toolset, ctx, allowed } of session.toolsets)
        toolset.register(auditInvalidInput(server, ctx), ctx, allowed);
      return server;
    },
    { onerror: (error) => console.error("[mcp]", error) },
  );

  // An admin's change to apps, keys or access, kept with the agents' tool calls (no agent;
  // `agent_name` is the admin). Key changes are filed under "agents".
  const adminEntry = (org: string, by: string, app: string, action: string, target: string) => ({
    organizationId: org,
    agentId: null,
    agentName: by,
    connectorId: app,
    tool: `admin:${action}`,
    target,
    decision: "allowed" as const,
    reason: null,
    vendorStatus: null,
  });
  async function logAdmin(org: string, by: string, app: string, action: string, target: string) {
    await store.audit(adminEntry(org, by, app, action, target));
  }

  const handlers: Handlers<false> = {
    listOrganizations: async () => ok(await store.listOrganizations()),
    createOrganization: async ({ body }) => ({
      status: 201,
      body: await store.createOrganization(body.name),
    }),
    getOrganization: async ({ params: [org] }) => some(await store.getOrganization(org)),
    ensureOrganization: async ({ params: [org], body }) =>
      ok(await store.ensureOrganization(org, body.name)),
    branding: async () => ok(brand()),
    connectors: async () => ok(service.connectors()),
    providers: async () => ok(service.providers()),
    providerConnections: async ({ params: [org] }) => ok(await service.listProviders(org)),
    connectProvider: async ({ params: [org, provider], body }) =>
      ok(await service.connectProvider({ organizationId: org, provider, ...body })),
    checkProvider: async ({ params: [org, provider] }) =>
      some(await service.checkProvider(org, provider)),
    newCertificate: async ({ params: [org, provider], body }) =>
      ok(await service.newCertificate(org, provider, body.by)),
    activateCertificate: async ({ params: [org, provider], body }) =>
      ok(await service.activateCertificate(org, provider, body.by)),
    setProviderSettings: async ({ params: [org, provider], body }) =>
      ok(await service.setProviderSettings(org, provider, body.values, body.by)),
    disconnectProvider: async ({ params: [org, provider], body }) =>
      found(await service.disconnectProvider(org, provider, body.by)),
    confirmed: async ({ params: [org] }) => ok(await service.listConfirmed(org)),
    confirm: async ({ params: [org, step], body }) => {
      await service.setConfirmed(org, step, body.by, true);
      return none;
    },
    unconfirm: async ({ params: [org, step], body }) => {
      await service.setConfirmed(org, step, body.by, false);
      return none;
    },
    connections: async ({ params: [org] }) => ok(await service.listConnections(org)),
    connect: async ({ params: [org, app], body }) => {
      const result = await service.connect({ organizationId: org, connectorId: app, ...body });
      await logAdmin(
        org,
        body.connectedBy,
        app,
        "connect_app",
        `capabilities ${result.connection.capabilities.join(", ")}`,
      );
      return ok(result);
    },
    startConsent: async ({ params: [org, app], body }) =>
      ok(service.startConsent({ organizationId: org, connectorId: app, ...body })),
    finishConsent: async ({ params: [org, app], body }) => {
      const result = await service.finishConsent({
        organizationId: org,
        connectorId: app,
        ...body,
      });
      if ("connection" in result && result.connection)
        await logAdmin(
          org,
          body.connectedBy,
          app,
          "connect_app",
          `capabilities ${result.connection.capabilities.join(", ")}`,
        );
      return ok(result);
    },
    setCeiling: async ({ params: [org, app], body }) => {
      const saved = await service.setCeiling(org, app, body.capabilities);
      if (saved)
        await logAdmin(org, body.by, app, "set_ceiling", `capabilities ${saved.join(", ")}`);
      return found(saved !== null);
    },
    check: async ({ params: [org, app], body, url }) => {
      const record = url.searchParams.get("record") !== "false";
      // A recorded check can cut every agent off, so it names the admin who asked.
      if (record && !body) throw new InvalidInput("Body must be JSON");
      const result = await service.check(org, app, { record, by: body?.by });
      return result ? ok(result) : { status: 404, body: { error: "Not connected" } };
    },
    setSites: async ({ params: [org, app], body }) => {
      const saved = await service.setSites(org, app, body.sites);
      if (!saved) throw new ConnectError("Turn the app on first");
      await logAdmin(org, body.by, app, "set_sites", saved.sites.join(" ") || "none");
      return ok(saved);
    },
    checkSites: async ({ params: [org, app], body }) => {
      const result = await service.siteAccess(org, app, body.site);
      if (!result) throw new ConnectError("Turn the app on first");
      return ok(result);
    },
    agentAccess: async ({ params: [org, app], body }) => {
      const result = await service.agentAccess(org, app, body.value, body.by);
      if (!result) throw new ConnectError("Turn the app on first");
      return ok(result);
    },
    disconnect: async ({ params: [org, app], body }) => {
      const removed = await store.removeConnection(org, app);
      if (removed) await logAdmin(org, body.by, app, "disconnect_app", app);
      return found(removed);
    },
    agents: async ({ params: [org] }) => ok(await store.listAgents(org)),
    createAgent: async ({ params: [org], body }) => {
      const key = newAgentKey();
      const agent = await store.createAgent(
        {
          organizationId: org,
          name: body.name,
          keyHash: hashAgentKey(key),
          createdBy: body.createdBy,
        },
        (a) => adminEntry(org, body.createdBy, "agents", "create_key", `agent ${a.id} "${a.name}"`),
      );
      if (!agent)
        return { status: 409, body: { error: `An agent named ${body.name} already exists` } };
      return { status: 201, body: { agent, key } };
    },
    setAgentStatus: async ({ params: [org, id], body }) =>
      found(
        await store.updateAgent(
          org,
          id,
          { status: body.status },
          adminEntry(org, body.by, "agents", `key_${body.status}`, `agent ${id}`),
        ),
      ),
    newKey: async ({ params: [org, id], body }) => {
      const key = newAgentKey();
      const saved = await store.updateAgent(
        org,
        id,
        { keyHash: hashAgentKey(key) },
        adminEntry(org, body.by, "agents", "rotate_key", `agent ${id}`),
      );
      return saved ? ok({ key }) : missing;
    },
    deleteAgent: async ({ params: [org, id], body }) => {
      const deleted = await store.deleteAgent(org, id);
      if (deleted) await logAdmin(org, body.by, "agents", "delete_key", `agent ${id}`);
      return found(deleted);
    },
    setGrant: async ({ params: [org, id, app], body }) => {
      const saved = await service.setGrant(org, id, app, body.capabilities, body.limits);
      if (saved)
        await logAdmin(
          org,
          body.by,
          app,
          "set_grant",
          `agent ${id} capabilities ${saved.capabilities.join(", ")} limits ${JSON.stringify(saved.limits)}`,
        );
      return found(saved !== null);
    },
    removeGrant: async ({ params: [org, id, app], body }) => {
      const removed = await store.removeGrant(org, id, app);
      if (removed) await logAdmin(org, body.by, app, "remove_grant", `agent ${id}`);
      return found(removed);
    },
    audit: async ({ params: [org], url }) => {
      const limit = Math.min(Math.max(Number(url.searchParams.get("limit") ?? 50) || 50, 1), 500);
      const agentId = url.searchParams.get("agent");
      if (agentId !== null && !new RegExp(`^${UUID}$`).test(agentId))
        throw new InvalidInput("agent must be an agent id");
      const connectorId = url.searchParams.get("app");
      if (connectorId !== null && !new RegExp(`^${PARAMS.app}$`).test(connectorId))
        throw new InvalidInput("app must be an app id");
      return ok(await store.recentAudit(org, limit, { agentId, connectorId }));
    },
  };

  const teams = deps.teams?.admin;
  const teamsHandlers: Handlers<true> | null = teams
    ? {
        teamsBots: async ({ params: [org] }) => ok(await teams.listBots(org)),
        teamsSetupState: async ({ params: [org] }) => ok(await teams.setupState(org)),
        finishTeamsPublish: async ({ params: [org], body }) =>
          ok(await teams.finishPublish(org, body)),
        createTeamsBot: async ({ params: [org, id], body }) => {
          const bot = await teams.createBot(org, id, body.displayName, body.by);
          if (!bot) throw new InvalidInput("No such agent");
          return ok(bot);
        },
        connectTeamsBot: async ({ params: [org, id], body }) =>
          ok(await teams.connectBot(org, id, body.botAppId, body.by)),
        removeTeamsBot: async ({ params: [org, id], body }) =>
          some(await teams.removeBot(org, id, body.by)),
        addToTeams: async ({ params: [org, id], body }) =>
          ok(await teams.addToTeams(org, id, body.displayName, body.by)),
        startTeamsPublish: async ({ params: [org, id], body }) =>
          ok(await teams.startPublish(org, id, body.redirectUri)),
        teamsTargets: async ({ params: [org, id], url }) => {
          const kind = url.searchParams.get("kind");
          if (kind !== "team" && kind !== "person")
            throw new InvalidInput("kind must be team or person");
          return ok(await teams.installTargets(org, id, kind, url.searchParams.get("q") ?? ""));
        },
        installTeamsBot: async ({ params: [org, id], body }) =>
          ok(await teams.install(org, id, body.kind, body.id, body.by)),
        teamsKit: async ({ params: [org, id] }) => ok({ data: await teams.setupScript(org, id) }),
        checkTeamsBot: async ({ params: [org, id] }) => ok(await teams.checkBot(org, id)),
        teamsTest: async ({ params: [org, id], body }) =>
          ok(await teams.sendTest(org, id, body.conversationId, body.by)),
        teamsCertificate: async ({ params: [org, id] }) =>
          ok({
            data: await teams.certificate(org, id),
            filename: `${brand().resourcePrefix}-bot.cer`,
          }),
        teamsPackage: async ({ params: [org, id] }) =>
          ok({ data: new Uint8Array(await teams.packageZip(org, id)), filename: "teams-app.zip" }),
        newTeamsCertificate: async ({ params: [org, id], body }) =>
          ok(await teams.newCertificate(org, id, body.by)),
        activateTeamsCertificate: async ({ params: [org, id], body }) =>
          ok(await teams.activateCertificate(org, id, body.by)),
        teamsTrigger: async ({ params: [org, id] }) => some(await teams.trigger(org, id)),
        setTeamsTrigger: async ({ params: [org, id], body }) => {
          const saved = await teams.setTrigger(org, id, body);
          if (!saved) throw new InvalidInput("No such agent");
          return ok(saved);
        },
        teamsRuns: async ({ params: [org, id], url }) => {
          const limit = Math.min(
            Math.max(Number(url.searchParams.get("limit") ?? 20) || 20, 1),
            200,
          );
          return ok(await teams.runs(org, id, limit));
        },
        claudeKey: async ({ params: [org] }) => ok(await teams.claudeKey(org)),
        setClaudeKey: async ({ params: [org], body }) =>
          ok(await teams.setClaudeKey(org, body.apiKey, body.by)),
        removeClaudeKey: async ({ params: [org], body }) =>
          found(await teams.removeClaudeKey(org, body.by)),
        claudeAgents: async ({ params: [org] }) => ok(await teams.claudeAgents(org)),
      }
    : null;

  const table = (Object.entries(routes) as [keyof Routes, RouteSpec][]).flatMap(([id, spec]) => {
    const run = (spec.teams ? teamsHandlers : handlers)?.[id as never] as
      Handler<RouteSpec> | undefined;
    if (!run) return [];
    const pattern = new RegExp(
      `^${spec.path.replace(/\{(\w+)\}/g, (_, p: keyof typeof PARAMS) => `(${PARAMS[p]})`)}$`,
    );
    return [{ spec, pattern, run }];
  });

  async function readBody(req: Request, schema: z.ZodType): Promise<unknown> {
    const text = await req.text();
    let raw: unknown;
    try {
      raw = text === "" ? undefined : JSON.parse(text);
    } catch {
      throw new InvalidInput("Body must be JSON");
    }
    const parsed = schema.safeParse(raw);
    if (parsed.success) return parsed.data;
    if (raw === undefined) throw new InvalidInput("Body must be JSON");
    throw new InvalidInput(z.prettifyError(parsed.error));
  }

  function send(
    spec: RouteSpec,
    { status, body }: { status: number; body: unknown },
    legacy: boolean,
  ): Response {
    const answer = spec.responses[status];
    if (answer === undefined) throw new Error(`${spec.path} answered an undeclared ${status}`);
    if (answer === null) return new Response(null, { status });
    if ("file" in answer) {
      const { data, filename } = body as { data: string | Uint8Array; filename?: string };
      return new Response(data, {
        status,
        headers: {
          "content-type": answer.file === "text/plain" ? "text/plain; charset=utf-8" : answer.file,
          ...(filename ? { "content-disposition": `attachment; filename="${filename}"` } : {}),
        },
      });
    }
    if (deps.validateResponses) {
      const checked = answer.safeParse(body);
      if (!checked.success)
        throw new Error(
          `${spec.method} ${spec.path} answered ${status} off its schema: ${z.prettifyError(checked.error)}`,
        );
    }
    return reply(status, legacy ? legacyFields(body) : body);
  }

  async function admin(req: Request, url: URL): Promise<Response> {
    // The admin API is for the admin app inside the cluster. Anything that came through
    // the public load balancer carries X-Forwarded-For and is refused outright.
    if (req.headers.has("x-forwarded-for")) return reply(404, { error: "Not found" });
    const auth = req.headers.get("authorization") ?? "";
    const token = auth.startsWith("Bearer ") ? auth.slice(7) : "";
    if (!token || !tokensEqual(token, deps.adminToken))
      return reply(401, { error: "Unauthorized" });
    const v1 = url.pathname.startsWith(`${ADMIN}/v1/`);
    const path = url.pathname.slice((v1 ? `${ADMIN}/v1` : ADMIN).length);
    if (v1 && path === "/openapi.json" && req.method === "GET") return reply(200, openapi());
    for (const { spec, pattern, run } of table) {
      const m = pattern.exec(path);
      if (!m || spec.method !== req.method) continue;
      try {
        const org = new RegExp(`^/orgs/(${UUID})/`).exec(path)?.[1];
        // v1 answers only for an organization made through it (PUT or POST /orgs). The
        // deprecated prefix makes one it does not know, as it did before organizations
        // were kept here.
        if (org && !v1) await store.ensureOrganization(org);
        if (org && v1 && !(await store.getOrganization(org)))
          return reply(404, { error: "No such organization" });
        const body = spec.body ? await readBody(req, spec.body) : undefined;
        return send(spec, await run({ params: m.slice(1), body: body as never, url }), !v1);
      } catch (error) {
        if (error instanceof InvalidInput) return reply(400, { error: error.message });
        if (error instanceof ConnectError) return reply(422, { error: error.message });
        throw error;
      }
    }
    return reply(404, { error: "Not found" });
  }

  // `only`: the per-agent address (/mcp/<agent id>), which takes that agent's key alone, so a
  // Claude vault holding a key for it holds that agent's key.
  async function agentMcp(req: Request, only?: string): Promise<Response> {
    // Agents call from servers, never browsers.
    if (req.headers.has("origin"))
      return reply(403, { error: "Browser requests are not accepted" });
    const auth = req.headers.get("authorization") ?? "";
    const session = auth.startsWith("Bearer ") ? await service.agentSession(auth.slice(7)) : null;
    if (!session || (only !== undefined && session.agent.id !== only))
      return reply(
        401,
        { error: "A valid agent key is required" },
        {
          "www-authenticate": `Bearer realm="${brand().resourcePrefix}-agent-connect"`,
        },
      );
    if (req.method === "POST") {
      const text = await readCapped(req, MAX_MCP_BODY);
      if (text === null) return reply(413, { error: "The request is too large" });
      const sent = messages(text);
      if (sent.length > MAX_BATCH)
        return reply(400, { error: `Send at most ${MAX_BATCH} messages in one request` });
      for (const { name, args } of calledTools(sent)) await session.refuseUngranted(name, args);
      const headers = new Headers(req.headers);
      headers.delete("content-length");
      req = new Request(req.url, { method: "POST", headers, body: text });
    }
    return mcp.fetch(req, {
      authInfo: { token: "", clientId: session.agent.id, scopes: [], extra: { session } },
    });
  }

  return {
    async fetch(req: Request): Promise<Response> {
      const url = new URL(req.url);
      try {
        if (url.pathname === `${BASE}/healthz`) return new Response("ok");
        if (url.pathname === `${BASE}/mcp`) return await agentMcp(req);
        const own = new RegExp(`^${BASE}/mcp/(${UUID})$`).exec(url.pathname);
        if (own) return await agentMcp(req, own[1]);
        // Teams, for an agent's bot: public, proven by Microsoft's token for that bot.
        const bot = new RegExp(`^${BASE}/teams/bots/(${UUID})/messages$`).exec(url.pathname);
        if (bot && deps.teams && req.method === "POST")
          return await deps.teams.runtime.inbound(bot[1], req);
        if (url.pathname.startsWith(`${ADMIN}/`)) return await admin(req, url);
        return reply(404, { error: "Not found" });
      } catch (error) {
        console.error("[agent-connect]", req.method, url.pathname, error);
        return reply(500, { error: "Internal error" });
      }
    },
  };
}
