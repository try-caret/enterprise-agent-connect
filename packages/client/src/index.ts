import { chosenSteps, type AppLive } from "./steps.js";
import type {
  Agent,
  Answer,
  CheckOutcome,
  Connection,
  Connector,
  Limits,
  SiteAccess,
  TeamsTrigger,
  TeamsTriggerInput,
} from "./types.js";

export * from "./steps.js";
export type * from "./types.js";

/** A refusal the admin should read (4xx with an `error`): bad input, a failed connection
 *  check, a name in use. A 5xx or an answer without an `error` throws a plain Error. */
export class AgentConnectError extends Error {
  constructor(
    message: string,
    readonly status: number,
  ) {
    super(message);
  }
}

export type ClientOptions = {
  /** Where the service answers, e.g. http://agent-connect:3000. */
  baseUrl: string;
  /** The admin token (AGENT_CONNECT_ADMIN_TOKEN). */
  token: string;
  /** Connecting Microsoft retries while its approval settles (up to ~20s). Default 45s. */
  timeoutMs?: number;
  /** Default: the global fetch at the time of each call. */
  fetch?: typeof fetch;
};

export type AgentConnectClient = ReturnType<typeof createClient>;

/** The admin API (v1), server to server: the token must never reach a browser. */
export function createClient(options: ClientOptions) {
  const base = `${options.baseUrl.replace(/\/+$/, "")}/agent-connect/admin/v1`;

  async function send(method: string, path: string, body?: unknown): Promise<Response> {
    return (options.fetch ?? globalThis.fetch)(`${base}${path}`, {
      method,
      headers: {
        authorization: `Bearer ${options.token}`,
        ...(body === undefined ? {} : { "content-type": "application/json" }),
      },
      body: body === undefined ? undefined : JSON.stringify(body),
      cache: "no-store",
      signal: AbortSignal.timeout(options.timeoutMs ?? 45_000),
    });
  }

  async function refusal(res: Response, method: string, path: string): Promise<never> {
    const data: unknown = await res.json().catch(() => null);
    const message = (data as { error?: unknown } | null)?.error;
    if (res.status < 500 && typeof message === "string")
      throw new AgentConnectError(message, res.status);
    throw new Error(`agent-connect ${method} ${path} answered ${res.status}`);
  }

  async function call<T>(method: string, path: string, body?: unknown): Promise<T> {
    const res = await send(method, path, body);
    if (res.status === 204) return undefined as T;
    if (res.ok) return (await res.json()) as T;
    return refusal(res, method, path);
  }

  /** A row that may not exist yet: 404 is null, not an error. */
  async function maybe<T>(path: string): Promise<T | null> {
    const res = await send("GET", path);
    if (res.status === 404) return null;
    if (res.ok) return (await res.json()) as T;
    return refusal(res, "GET", path);
  }

  const seg = encodeURIComponent;
  const org = (id: string) => `/orgs/${seg(id)}`;
  const app = (o: string, a: string) => `${org(o)}/connections/${seg(a)}`;
  const provider = (o: string, p: string) => `${org(o)}/providers/${seg(p)}`;
  const agent = (o: string, id: string) => `${org(o)}/agents/${seg(id)}`;
  const bot = (o: string, id: string) => `${agent(o, id)}/teams-bot`;

  const client = {
    organizations: {
      list: () => call<Answer<"listOrganizations">>("GET", "/orgs"),
      get: (id: string) => maybe<Answer<"getOrganization">>(org(id)),
      /** Makes the organization with this id if missing; `name` replaces the kept one. */
      ensure: (id: string, name?: string) =>
        call<Answer<"ensureOrganization">>("PUT", org(id), name === undefined ? {} : { name }),
      create: (name: string) => call<Answer<"createOrganization">>("POST", "/orgs", { name }),
    },
    branding: () => call<Answer<"branding">>("GET", "/branding"),
    connectors: () => call<Answer<"connectors">>("GET", "/connectors"),
    providers: () => call<Answer<"providers">>("GET", "/providers"),
    providerConnections: (o: string) =>
      call<Answer<"providerConnections">>("GET", `${org(o)}/providers`),
    connectProvider: (
      o: string,
      p: string,
      body: { fields: Record<string, string>; connectedBy: string },
    ) => call<Answer<"connectProvider">>("PUT", provider(o, p), body),
    checkProvider: (o: string, p: string) =>
      call<Answer<"checkProvider">>("POST", `${provider(o, p)}/check`),
    // `by`: the admin, recorded in the service's audit log with each change.
    /** Before connecting, the certificate to upload; after, one to rotate to. */
    newCertificate: (o: string, p: string, by: string) =>
      call<Answer<"newCertificate">>("POST", `${provider(o, p)}/certificate`, { by }),
    /** Signs in with the new certificate and, if that works, switches to it. */
    activateCertificate: (o: string, p: string, by: string) =>
      call<Answer<"activateCertificate">>("POST", `${provider(o, p)}/certificate/activate`, {
        by,
      }),
    /** Disconnects the provider and every app signing in through it. */
    disconnectProvider: (o: string, p: string, by: string) =>
      call<void>("DELETE", provider(o, p), { by }),
    /** Values an app's steps ask for, kept on the provider connection. */
    setProviderSettings: (o: string, p: string, values: Record<string, string>, by: string) =>
      call<Answer<"setProviderSettings">>("PUT", `${provider(o, p)}/settings`, { values, by }),
    connections: (o: string) => call<Answer<"connections">>("GET", `${org(o)}/connections`),
    connect: (
      o: string,
      a: string,
      body: { capabilities: string[]; fields: Record<string, string>; connectedBy: string },
    ) => call<Answer<"connect">>("PUT", app(o, a), body),
    startConsent: (o: string, a: string, body: { capabilities: string[]; redirectUri: string }) =>
      call<Answer<"startConsent">>("POST", `${app(o, a)}/consent`, body),
    /** `redirect`: the approval came back; the admin signs in to that organization next.
     *  `provider`: the approval connected the provider; the app still needs its own settings. */
    finishConsent: (
      o: string,
      a: string,
      body: { state: string; nonce: string; params: Record<string, string>; connectedBy: string },
    ) => call<Answer<"finishConsent">>("POST", `${app(o, a)}/consent/callback`, body),
    setCeiling: (o: string, a: string, capabilities: string[], by: string) =>
      call<void>("PATCH", app(o, a), { capabilities, by }),
    /** Recorded: a failure marks the app as needing reconnecting, so it names the admin. */
    check: (o: string, a: string, by: string) =>
      call<CheckOutcome>("POST", `${app(o, a)}/check`, { by }),
    /** The same checks, read only: a failure here does not mark the app. */
    peek: (o: string, a: string) => call<CheckOutcome>("POST", `${app(o, a)}/check?record=false`),
    /** The sites added to an app, the whole list at once. */
    setSites: (o: string, a: string, sites: string[], by: string) =>
      call<Answer<"setSites">>("PUT", `${app(o, a)}/sites`, { sites, by }),
    /** Whether the app can open each added site, or just `site`. */
    checkSites: (o: string, a: string, site?: string) =>
      call<Answer<"checkSites">>(
        "POST",
        `${app(o, a)}/sites/check`,
        site === undefined ? {} : { site },
      ),
    /** Whether the app can reach one agent's own address (its mailbox) yet. */
    agentAccess: (o: string, a: string, value: string, by: string) =>
      call<Answer<"agentAccess">>("POST", `${app(o, a)}/agent-access`, { value, by }),
    disconnect: (o: string, a: string, by: string) => call<void>("DELETE", app(o, a), { by }),
    agents: (o: string) => call<Answer<"agents">>("GET", `${org(o)}/agents`),
    createAgent: (o: string, body: { name: string; createdBy: string }) =>
      call<Answer<"createAgent">>("POST", `${org(o)}/agents`, body),
    setAgentStatus: (o: string, id: string, status: Agent["status"], by: string) =>
      call<void>("PATCH", agent(o, id), { status, by }),
    newKey: (o: string, id: string, by: string) =>
      call<Answer<"newKey">>("POST", `${agent(o, id)}/key`, { by }),
    deleteAgent: (o: string, id: string, by: string) => call<void>("DELETE", agent(o, id), { by }),
    setGrant: (
      o: string,
      id: string,
      a: string,
      capabilities: string[],
      limits: Limits,
      by: string,
    ) => call<void>("PUT", `${agent(o, id)}/grants/${seg(a)}`, { capabilities, limits, by }),
    removeGrant: (o: string, id: string, a: string, by: string) =>
      call<void>("DELETE", `${agent(o, id)}/grants/${seg(a)}`, { by }),
    /** Recent requests, newest first, of one key (`agent`), one app, or both. */
    audit: (o: string, filter: { agent?: string; app?: string }, limit = 50) =>
      call<Answer<"audit">>(
        "GET",
        `${org(o)}/audit?${new URLSearchParams({ limit: String(limit), ...filter })}`,
      ),
    // Setup steps an admin marked done (they cannot be checked), shared by every admin.
    confirmed: async (o: string) =>
      (await call<Answer<"confirmed">>("GET", `${org(o)}/confirmed`)).map((c) => c.step),
    setConfirmed: (o: string, step: string, on: boolean, by: string) =>
      call<void>(on ? "PUT" : "DELETE", `${org(o)}/confirmed/${seg(step)}`, { by }),

    // ---- Teams: each agent's own bot ----
    teamsBots: (o: string) => call<Answer<"teamsBots">>("GET", `${org(o)}/teams/bots`),
    createTeamsBot: (o: string, id: string, displayName: string, by: string) =>
      call<Answer<"createTeamsBot">>("POST", bot(o, id), { displayName, by }),
    /** Proves the application ID the setup script printed by signing in as it. */
    connectTeamsBot: (o: string, id: string, botAppId: string, by: string) =>
      call<Answer<"connectTeamsBot">>("PUT", bot(o, id), { botAppId, by }),
    checkTeamsBot: (o: string, id: string) =>
      call<Answer<"checkTeamsBot">>("POST", `${bot(o, id)}/check`),
    teamsTest: (o: string, id: string, conversationId: string, by: string) =>
      call<Answer<"teamsTest">>("POST", `${bot(o, id)}/test`, { conversationId, by }),
    /** Forgets the bot. What Add to Teams made is deleted (`failed`: what could not be);
     *  `removal` is the commands for the rest. */
    removeTeamsBot: (o: string, id: string, by: string) =>
      call<Answer<"removeTeamsBot">>("DELETE", bot(o, id), { by }),
    newTeamsCertificate: (o: string, id: string, by: string) =>
      call<Answer<"newTeamsCertificate">>("POST", `${bot(o, id)}/certificate`, { by }),
    activateTeamsCertificate: (o: string, id: string, by: string) =>
      call<Answer<"activateTeamsCertificate">>("POST", `${bot(o, id)}/certificate/activate`, {
        by,
      }),
    /** The Teams page's set-up-once steps, checked live. */
    teamsSetupState: (o: string) => call<Answer<"teamsSetupState">>("GET", `${org(o)}/teams/setup`),
    /** Makes the agent's bot in the tenant, each part only if missing. */
    addToTeams: (o: string, id: string, displayName: string, by: string) =>
      call<Answer<"addToTeams">>("POST", `${bot(o, id)}/add`, { displayName, by }),
    startTeamsPublish: (o: string, id: string, redirectUri: string) =>
      call<Answer<"startTeamsPublish">>("POST", `${bot(o, id)}/publish`, { redirectUri }),
    finishTeamsPublish: (
      o: string,
      body: { state: string; nonce: string; params: Record<string, string>; by: string },
    ) => call<Answer<"finishTeamsPublish">>("POST", `${org(o)}/teams/publish/callback`, body),
    teamsTargets: (o: string, id: string, kind: "team" | "person", q: string) =>
      call<Answer<"teamsTargets">>(
        "GET",
        `${bot(o, id)}/targets?${new URLSearchParams({ kind, q })}`,
      ),
    installTeamsBot: (
      o: string,
      id: string,
      kind: "team" | "chat" | "person",
      target: string,
      by: string,
    ) => call<Answer<"installTeamsBot">>("POST", `${bot(o, id)}/install`, { kind, id: target, by }),
    /** A file the service makes (the setup script, certificate or Teams app package), as is. */
    teamsFile: async (o: string, id: string, file: "kit" | "package" | "certificate") => {
      const path = `${bot(o, id)}/${file}`;
      const res = await send("GET", path);
      return res.ok ? res : refusal(res, "GET", path);
    },
    teamsTrigger: (o: string, id: string) => maybe<TeamsTrigger>(`${agent(o, id)}/teams-trigger`),
    setTeamsTrigger: (o: string, id: string, body: TeamsTriggerInput) =>
      call<TeamsTrigger>("PUT", `${agent(o, id)}/teams-trigger`, body),
    teamsRuns: (o: string, id: string, limit = 20) =>
      call<Answer<"teamsRuns">>("GET", `${agent(o, id)}/teams-runs?limit=${limit}`),
    // The organization's Claude API key: write-only, shown by its last four characters.
    claudeKey: (o: string) => call<Answer<"claudeKey">>("GET", `${org(o)}/claude-key`),
    setClaudeKey: (o: string, apiKey: string, by: string) =>
      call<Answer<"setClaudeKey">>("PUT", `${org(o)}/claude-key`, { apiKey, by }),
    removeClaudeKey: (o: string, by: string) =>
      call<void>("DELETE", `${org(o)}/claude-key`, { by }),
    claudeAgents: (o: string) => call<Answer<"claudeAgents">>("GET", `${org(o)}/claude-key/agents`),
  };

  /** One app's live site check, or null when the app cannot run it (e.g. the vendor refuses
   *  the sign-in): the page then shows the app as needing attention instead of failing. */
  async function checkedSites(o: string, a: string): Promise<SiteAccess[] | null> {
    try {
      return (await client.checkSites(o, a)).sites;
    } catch (error) {
      if (!(error instanceof AgentConnectError)) throw error;
      console.error(`[agent-connect] site check for ${a}`, error.message);
      return null;
    }
  }

  const peeked = (o: string, a: string) =>
    client.peek(o, a).catch((error: unknown) => {
      if (!(error instanceof AgentConnectError)) throw error;
      console.error(`[agent-connect] check for ${a}`, error.message);
      return null;
    });

  return {
    ...client,
    checkedSites,

    /** The read-only live check of each app of its own that is on, for warnings no stored
     *  status holds (e.g. an app can see nothing yet). Null for an app whose check could not
     *  run. */
    async appChecks(
      o: string,
      connectors: Connector[],
      connections: Connection[],
    ): Promise<Record<string, CheckOutcome | null>> {
      const apps = connectors.filter(
        (c) =>
          !c.provider &&
          connections.some((x) => x.connectorId === c.id && x.status === "connected"),
      );
      const results = await Promise.all(apps.map((c) => peeked(o, c.id)));
      return Object.fromEntries(apps.map((c, i) => [c.id, results[i]]));
    },

    /** Each provider app's live state for its steps (`appState`), read only: its own check
     *  and its sites while on, the organization's Claude key where a capability needs one,
     *  and its chosen steps' own live checks (none while those steps are not chosen, so no
     *  vendor call). An app's page, its provider's page and the Integrations tab all read it
     *  here, so they agree. The Claude key and the step checks are read once for all apps. */
    async appStates(
      o: string,
      connectors: Connector[],
      connections: Connection[],
    ): Promise<Record<string, AppLive>> {
      const apps = connectors.map((connector) => {
        const connection = connections.find((c) => c.connectorId === connector.id) ?? null;
        const liveSteps = chosenSteps(
          connector,
          [...connector.setup, ...connector.afterConnect],
          connection?.config ?? {},
        ).filter((s) => s.live);
        return { connector, connection, liveSteps };
      });
      const liveSteps = apps.flatMap((a) => a.liveSteps);
      const [claudeKey, live, perApp] = await Promise.all([
        apps.some((a) => a.connector.capabilities.some((c) => c.claudeKey))
          ? client.claudeKey(o)
          : null,
        liveSteps.length
          ? client.teamsSetupState(o).catch((error: unknown) => {
              if (!(error instanceof AgentConnectError)) {
                console.error("[agent-connect] step checks", error);
                return null;
              }
              // The service's own answer (e.g. what the connection lacks) goes on each step.
              const failing = { state: "failing" as const, message: error.message };
              return Object.fromEntries(liveSteps.map((s) => [s.live, failing]));
            })
          : null,
        Promise.all(
          apps.map(({ connector, connection }) =>
            Promise.all([
              connection ? peeked(o, connector.id) : null,
              connection && connector.siteAccess ? checkedSites(o, connector.id) : null,
            ]),
          ),
        ),
      ]);
      return Object.fromEntries(
        apps.map((a, i) => [
          a.connector.id,
          {
            health: perApp[i][0],
            sites: perApp[i][1],
            claudeKey: claudeKey?.set === true,
            live: a.liveSteps.length ? live : null,
          },
        ]),
      );
    },
  };
}
