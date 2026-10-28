import type {
  Agent,
  AgentStep,
  Capability,
  CertificateInfo,
  Check,
  CheckOutcome,
  ChosenOption,
  Connection,
  Connector,
  Field,
  LiveCheck,
  Limits,
  Provider,
  ProviderConnection,
  SiteAccess,
  Step,
  TeamsBot,
  TeamsPlace,
  TeamsRun,
  TeamsTrigger,
  TeamsTriggerInput,
} from "./types.js";

// What the admin pages show, worked out from the admin API's answers: setup steps and their
// placeholders, statuses, the Integrations tab's grouping, Teams setup and new key setup.
// Pure functions, no fetch, so a browser bundle can use them.

/** A copy block's text: a line with {scope} repeats once per scope, {scopes} becomes the
 *  scopes space-separated, and any other {key} is filled from the connection's settings.
 *  Unknown placeholders stay visible so a missing value is obvious. */
export function expandCopy(
  lines: string[],
  scopes: string[],
  config: Record<string, string> = {},
  /** How {scopes} is joined, e.g. ", " in click steps people read. */
  joiner = " ",
): string {
  return lines
    .flatMap((line) =>
      line.includes("{scope}") ? scopes.map((s) => line.replaceAll("{scope}", s)) : [line],
    )
    .map((line) =>
      line
        .replaceAll("{scopes}", scopes.join(joiner))
        .replace(/\{([A-Za-z_]+)\}/g, (whole, key: string) => config[key] ?? whole),
    )
    .join("\n");
}

/** {name} values the ticked capabilities fill into steps, later ones winning, e.g. the
 *  site role: read, or write once adding files is ticked. */
export function fillsFor(connector: Connector, ticked: readonly string[]): Record<string, string> {
  return Object.assign(
    {},
    ...connector.capabilities.filter((c) => ticked.includes(c.id)).map((c) => c.fill ?? {}),
  );
}

/** The choice fields' values: the stored or typed one, else each field's first option, which
 *  also holds while the earlier choice a field `requires` is not made. */
export function choiceValues(
  connector: { fields: Field[] },
  values: Record<string, string>,
): Record<string, string> {
  const chosen: Record<string, string> = {};
  for (const f of connector.fields)
    if (f.options)
      chosen[f.id] =
        (isChosen(f.requires, chosen) && f.options.find((o) => o.value === values[f.id])?.value) ||
        f.options[0].value;
  return chosen;
}

/** True with no condition, or while its option is the chosen one. */
export const isChosen = (option: ChosenOption | undefined, chosen: Record<string, string>) =>
  !option || chosen[option.field] === option.value;

/** The choice fields shown: one that `requires` an earlier choice only while it is made. */
export const askedFields = (fields: Field[], values: Record<string, string>) => {
  const chosen = choiceValues({ fields }, values);
  return fields.filter((f) => isChosen(f.requires, chosen));
};

/** The steps that apply under the chosen options (a step with `option` only while chosen)
 *  and, given `ticked`, the ticked capabilities (a step with `capabilities` only while one of
 *  them is). */
export function chosenSteps(
  connector: Connector,
  steps: Step[],
  values: Record<string, string>,
  ticked?: readonly string[],
): Step[] {
  const chosen = choiceValues(connector, values);
  return steps.filter(
    (s) =>
      isChosen(s.option, chosen) &&
      (!ticked || !s.capabilities || s.capabilities.some((c) => ticked.includes(c))),
  );
}

/** Choices that change the permissions the setup steps ask for: made in step 1, with what
 *  agents may do, so the steps after it already show those permissions. */
export function permissionChoices(connector: Connector): Field[] {
  return connector.fields.filter((f) => f.options?.some((o) => o.scopes?.length));
}

/** Permissions the chosen options need, e.g. Sites.Read.All for reading every site. */
export function optionScopes(connector: Connector, values: Record<string, string>): string[] {
  const chosen = choiceValues(connector, values);
  return connector.fields.flatMap(
    (f) => f.options?.find((o) => o.value === chosen[f.id])?.scopes ?? [],
  );
}

/** {site} and {siteRef} (how Graph and PowerShell name it: host, or host:/path) for a site's
 *  grant steps and note. */
export function siteFills(site: string): Record<string, string> {
  const { hostname, pathname } = new URL(site);
  return { site, siteRef: pathname === "/" ? hostname : `${hostname}:${pathname}` };
}

/** The note to send whoever manages SharePoint: the connector's request with the site, the
 *  app and the role filled, and its first grant step's commands where it says
 *  {grantCommands}. */
/** The site grant steps for the site role the ticks need (`site_role`). */
export function grantSteps(connector: Connector, siteRole: string | undefined): Step[] {
  return (connector.siteAccess?.grant ?? []).filter((s) => !s.siteRole || s.siteRole === siteRole);
}

export function siteRequest(
  connector: Connector,
  site: string,
  scopes: string[],
  config: Record<string, string>,
): string {
  const fills = { ...config, ...siteFills(site) };
  const first = connector.siteAccess?.grant[0]?.copy ?? [];
  return expandCopy([connector.siteAccess?.request ?? ""], scopes, {
    ...fills,
    grantCommands: expandCopy(first, scopes, fills),
  });
}

/** The apps a key uses, in the order the apps are listed. */
export function keyApps(agent: Agent, connectors: Connector[]): Connector[] {
  return connectors.filter((c) => agent.grants.some((g) => g.connectorId === c.id));
}

/** How many active keys use an app (a disabled key uses nothing). */
export function keysUsing(agents: Agent[], connectorId: string): number {
  return agents.filter(
    (a) => a.status === "active" && a.grants.some((g) => g.connectorId === connectorId),
  ).length;
}

/** The values a connected app reported for its discover probe, e.g. Jira project keys. */
export function discovered(connector: Connector, checks: Check[] | null): string[] {
  return discoveredChoices(connector, checks).map((c) => c.value);
}

/** The same, each with the label to show for it (the value itself when the app gave none),
 *  e.g. a NetSuite subsidiary's name for its id. */
export function discoveredChoices(
  connector: Connector,
  checks: Check[] | null,
): { value: string; label: string }[] {
  if (!connector.discover || !checks) return [];
  const check = checks.find((c) => c.label === connector.discover!.label);
  return (check?.values ?? []).map((value) => ({ value, label: check?.labels?.[value] ?? value }));
}

/** Parses free-typed list input ("KAN, SAM1") into values. */
export function splitList(raw: string): string[] {
  return [
    ...new Set(
      raw
        .split(/[\s,]+/)
        .map((v) => v.trim())
        .filter(Boolean),
    ),
  ];
}

/** Shown under Click through for a step the vendor offers only as commands. */
export const CLICK_UNAVAILABLE =
  "Microsoft has no admin center page for this step: it is done only with commands. Open Commands, or send them to whoever manages it.";

/** What a free-typed list box shows: the admin's own text while it still parses to the
 *  saved values (so a typed comma or space stays put), else the values themselves. */
export function listText(typed: string, values: string[]): string {
  const parsed = splitList(typed);
  return parsed.length === values.length && parsed.every((v, i) => v === values[i])
    ? typed
    : values.join(", ");
}

/** Scopes or in-app permissions the ticked capabilities need, in connector order, each once. */
export function neededFor(
  connector: Connector,
  ticked: readonly string[],
  kind: "scopes" | "permissions",
): string[] {
  return [
    ...new Set([
      ...(kind === "scopes" ? (connector.scopes ?? []) : []),
      ...connector.capabilities.filter((c) => ticked.includes(c.id)).flatMap((c) => c[kind]),
    ]),
  ];
}

/** The capabilities after ticking or unticking one, in connector order. Required ones stay on;
 *  a destructive one is only turned on when `confirm` accepts its warning. */
export function toggleCapability(
  connector: Connector,
  current: readonly string[],
  id: string,
  on: boolean,
  confirm: (warning: string) => boolean,
): string[] {
  const capability = connector.capabilities.find((c) => c.id === id);
  if (!capability || capability.required) return [...current];
  if (
    on &&
    capability.destructive &&
    !current.includes(id) &&
    !confirm(destructiveWarning(connector, capability))
  )
    return [...current];
  return connector.capabilities
    .filter((c) => c.required || (c.id === id ? on : current.includes(c.id)))
    .map((c) => c.id);
}

export function destructiveWarning(connector: Connector, capability: Capability): string {
  return `Turn on ${capability.label} in ${connector.name}? ${capability.description} Agents can do this without asking anyone first.`;
}

/** needs_steps: on, but a step on its page is not done yet. needs_attention: on, but a live
 *  check of it could not run, so whether it works is not known. */
export type AppStatus =
  | "connected"
  | "needs_reconnect"
  | "not_connected"
  | "needs_steps"
  | "needs_attention";

/** The provider form's values to send: only the fields on show, so a secret typed and then
 *  hidden by going back to the certificate is not sent (a secret sent wins the sign-in). */
export function providerFields(
  fields: readonly { id: string; secret?: boolean }[],
  useSecret: boolean,
  values: Record<string, string>,
): Record<string, string> {
  return Object.fromEntries(
    fields
      .filter((f) => (!f.secret || useSecret) && values[f.id] !== undefined)
      .map((f) => [f.id, values[f.id]]),
  );
}

/** Whether a provider connection is usable by its apps. */
export function providerReady(connection: ProviderConnection | null): boolean {
  return connection !== null && connection.status !== "pending";
}

/** A provider connection's status, as its card and page show it. */
export function providerStatus(connection: ProviderConnection | null): AppStatus {
  if (!providerReady(connection)) return "not_connected";
  return connection!.status === "connected" ? "connected" : "needs_reconnect";
}

/** The Integrations tab's layout: apps of their own first, then each provider's apps
 *  together under it, by their `order`, then name. Providers without apps are left out. */
export function groupByProvider<T extends { provider?: string; order?: number; name: string }>(
  apps: T[],
  providers: Provider[],
): { provider: Provider | null; apps: T[] }[] {
  const byOrder = (a: T, b: T) =>
    (a.order ?? Infinity) - (b.order ?? Infinity) || a.name.localeCompare(b.name);
  const own = apps.filter((a) => !a.provider).sort(byOrder);
  return [
    ...(own.length ? [{ provider: null, apps: own }] : []),
    ...providers
      .map((provider) => ({
        provider,
        apps: apps.filter((a) => a.provider === provider.id).sort(byOrder),
      }))
      .filter((g) => g.apps.length > 0),
  ];
}

/** What the Integrations tab says about a provider connection: its certificate running out
 *  with no new one waiting (every app of it stops that day). */
export function certificateNotice(connection: ProviderConnection | null, now = Date.now()) {
  const certificate = connection?.certificate;
  if (!certificate || connection.nextCertificate || !expiresSoon(certificate, now)) return null;
  return { expiresAt: certificate.expiresAt };
}

/** A certificate expiring within this many days is called out. */
export const CERTIFICATE_WARNING_DAYS = 30;

export function expiresSoon(info: CertificateInfo, now = Date.now()): boolean {
  return new Date(info.expiresAt).getTime() - now < CERTIFICATE_WARNING_DAYS * 86_400_000;
}

/** How a provider connection signs in, in the admin's words. */
export function signInMethod(connection: ProviderConnection): string {
  const auth = connection.config.auth;
  if (auth === "certificate") return "Certificate";
  if (auth === "client_secret") return "Client secret";
  if (auth === "arbium_app") return "The shared app, approved by your administrator";
  throw new Error(`Unknown sign-in ${auth}`);
}

/** A saved connection whose last check failed needs reconnecting, as does one marked so; one
 *  whose check passed with a warning (e.g. Jira can browse no project yet) needs steps. */
export function appStatus(
  connection: Connection | null,
  healthy = true,
  warned = false,
): AppStatus {
  if (!connection) return "not_connected";
  if (connection.status !== "connected" || !healthy) return "needs_reconnect";
  return warned ? "needs_steps" : "connected";
}

/** A check that passed but found something to do. */
export function checkWarned(health: CheckOutcome | null | undefined): boolean {
  return health?.ok === true && health.checks.some((c) => c.warning);
}

/** What an app's own page shows, so it never contradicts its badge: a failed check, a failed
 *  recorded check, or a check that passed with a warning each offer Check again, and a warning
 *  leaves Prepare and "What agents can see" not done. */
export function appPageState(
  connection: Connection | null,
  health: CheckOutcome | null | undefined,
) {
  const warned = checkWarned(health);
  const notice =
    health && !health.ok
      ? ("failed" as const)
      : connection?.status === "needs_reconnect"
        ? ("recorded" as const)
        : warned
          ? ("warned" as const)
          : null;
  return { notice, prepareDone: connection !== null && !warned, seeDone: !warned };
}

/** An app of its own on the Integrations tab: its stored status, then its live read-only
 *  check (failed: needs reconnecting; could not run: needs attention; warned: needs steps).
 *  Before the check answers (`checks` null in `integrationCards`), the stored status alone. */
export function cardStatus(connection: Connection | null, check: CheckOutcome | null): AppStatus {
  const base = appStatus(connection, check?.ok !== false, checkWarned(check));
  return base === "connected" && check === null ? "needs_attention" : base;
}

/** A provider's status as its card and its page's header show it: the sign-in's, then its
 *  apps' (reconnecting, attention, steps), with how many apps are on and how many need steps
 *  or attention, each app counted by `appState`, as its row and its own page show it. */
export function providerSummary(
  definition: Provider,
  connection: ProviderConnection | null,
  apps: { connector: Connector; connection: Connection | null; state: AppLive }[],
  confirmed: string[],
): { status: AppStatus; on: number; steps: number; attention: number } {
  const rows = apps.map(
    (a) =>
      appState({
        definition,
        provider: connection,
        connector: a.connector,
        connection: a.connection,
        state: a.state,
        marks: confirmed,
      }).status,
  );
  const steps = rows.filter((r) => r === "needs_steps").length;
  const attention = rows.filter((r) => r === "needs_attention" || r === "needs_reconnect").length;
  const signIn = providerStatus(connection);
  const status: AppStatus =
    signIn !== "connected"
      ? signIn
      : rows.includes("needs_reconnect")
        ? "needs_reconnect"
        : attention
          ? "needs_attention"
          : steps
            ? "needs_steps"
            : "connected";
  return { status, on: apps.filter((a) => a.connection).length, steps, attention };
}

// ---- Teams: each agent's own bot (AgentConnect src/teams/admin.ts) ----

export type TeamsStep = { label: string; done: boolean; detail?: string };

/** Where a bot's setup stands, in the order the admin does it. The service sees what Add to Teams
 *  made and what it published; Teams telling it the bot was added somewhere proves the rest. */
export function teamsSetup(bot: TeamsBot | null): TeamsStep[] {
  const signedIn = bot?.status === "connected";
  const places = bot?.installedIn ?? [];
  const published = bot?.publishedVersion ?? null;
  return [
    {
      label: "Bot made in your Microsoft tenant",
      done: Boolean(bot?.botAppId),
      detail: bot?.botAppId ?? undefined,
    },
    {
      label: "Bot signs in to Microsoft",
      done: signedIn,
      detail: bot?.status === "needs_reconnect" ? "The last sign-in failed" : undefined,
    },
    {
      label: "Published to your organization's Teams apps",
      // A bot set up with the script publishes itself; being added somewhere proves it.
      done: published !== null || places.length > 0,
      detail:
        published && bot && published !== bot.packageVersion
          ? "Its package changed: publish it again"
          : undefined,
    },
    {
      label: "Added to a team, chat or person",
      done: places.length > 0,
      detail: places.length
        ? places.map(placeName).join(", ")
        : (bot?.lastRefusedReason ?? undefined),
    },
  ];
}

export function placeName(place: TeamsPlace): string {
  if (place.name) return place.name;
  return place.kind === "personal"
    ? "A one-to-one chat"
    : place.kind === "groupChat"
      ? "A group chat"
      : "A team";
}

/** The trigger form as typed: dollars per answer and per day, comma or space separated ids,
 *  the teams it answers in as ticked group ids. */
export type TriggerForm = {
  enabled: boolean;
  /** Kept as loaded, with its settings: the form edits a Claude trigger's fields only. */
  runner: string;
  settings: Record<string, unknown>;
  managedAgentId: string;
  environmentId: string;
  vaultIds: string;
  budgetDollars: string;
  dailyDollars: string;
  maxRunsPerHour: string;
  replyMode: "mentions" | "all";
  allowedTeams: string[];
  allowedPeople: string;
  anyoneInOrganization: boolean;
};

/** The service's default runner, the only one whose fields the trigger form edits. */
export const CLAUDE_RUNNER = "claude-managed-agents";

export function triggerForm(trigger: TeamsTrigger | null): TriggerForm {
  return {
    enabled: trigger?.enabled ?? false,
    runner: trigger?.runner ?? CLAUDE_RUNNER,
    settings: trigger?.settings ?? {},
    managedAgentId: trigger?.managedAgentId ?? "",
    environmentId: trigger?.environmentId ?? "",
    vaultIds: trigger?.vaultIds.join(", ") ?? "",
    budgetDollars: trigger ? (trigger.budgetCents / 100).toFixed(2) : "1.00",
    dailyDollars: trigger ? (trigger.maxCentsPerDay / 100).toFixed(2) : "5.00",
    maxRunsPerHour: String(trigger?.maxRunsPerHour ?? 20),
    replyMode: trigger?.replyMode ?? "mentions",
    allowedTeams: trigger?.allowedTeams ?? [],
    allowedPeople: trigger?.allowedPeople.join(", ") ?? "",
    anyoneInOrganization: trigger?.anyoneInOrganization ?? false,
  };
}

const ID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/** The trigger to save, or what to fix. The service checks again, and that the Claude agent
 *  and environment exist under the organization's key. */
export function triggerInput(
  form: TriggerForm,
): { ok: true; value: Omit<TeamsTriggerInput, "by"> } | { ok: false; error: string } {
  const claude = form.runner === CLAUDE_RUNNER;
  const managedAgentId = form.managedAgentId.trim();
  const environmentId = form.environmentId.trim();
  if (claude && !managedAgentId) return { ok: false, error: "Enter the Claude agent's ID" };
  if (claude && !environmentId) return { ok: false, error: "Enter the Claude environment's ID" };
  const money = /^\d+(\.\d{1,2})?$/;
  const dollars = Number(form.budgetDollars);
  if (!money.test(form.budgetDollars.trim()) || dollars < 0.01 || dollars > 20)
    return { ok: false, error: "The spending limit per answer is between $0.01 and $20" };
  const daily = Number(form.dailyDollars);
  if (!money.test(form.dailyDollars.trim()) || daily < dollars || daily > 1000)
    return {
      ok: false,
      error: "The daily spending limit is between the limit per answer and $1,000",
    };
  const runs = Number(form.maxRunsPerHour);
  if (!Number.isInteger(runs) || runs < 1 || runs > 100)
    return { ok: false, error: "Answers per hour is a whole number from 1 to 100" };
  const allowedTeams = form.allowedTeams;
  const allowedPeople = splitList(form.allowedPeople);
  const bad = allowedPeople.find((v) => !ID.test(v));
  if (bad) return { ok: false, error: `${bad} is not a Microsoft ID (like 0f3e…-…)` };
  if (form.enabled && !form.anyoneInOrganization && !allowedPeople.length)
    return {
      ok: false,
      error: "Name the people who may ask it, or allow anyone in your organization",
    };
  return {
    ok: true,
    value: {
      enabled: form.enabled,
      runner: form.runner,
      settings: form.settings,
      ...(claude ? { managedAgentId, environmentId, vaultIds: splitList(form.vaultIds) } : {}),
      budgetCents: Math.round(dollars * 100),
      maxCentsPerDay: Math.round(daily * 100),
      maxRunsPerHour: runs,
      replyMode: form.replyMode,
      allowedTeams,
      allowedPeople,
      anyoneInOrganization: form.anyoneInOrganization,
    },
  };
}

/** One answer in Recent Teams activity, in the admin's words. */
export function runSummary(run: TeamsRun): string {
  const state = {
    starting: "Starting",
    running: "Answering",
    idle: "Answered",
    ended: "Finished",
    failed: "Failed",
    no_reply: "Finished without replying",
  }[run.state];
  const cost = run.costCents === null ? null : `$${(run.costCents / 100).toFixed(2)}`;
  const turns = run.turns === 1 ? "1 message" : `${run.turns} messages`;
  return [state, turns, cost].filter(Boolean).join(" · ");
}

/** An agent's address, limited to characters that are inert inside the commands it fills. */
export const AGENT_EMAIL = /^[A-Za-z0-9._+-]+@[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)+$/;

/** A Teams bot's name: the setup script hands it to az, which runs through cmd.exe on Windows,
 *  so only characters inert there (the same rule as the service's). */
export const BOT_NAME = /^[A-Za-z0-9 ._-]{1,30}$/;

/** A key's name as a starting bot name: other characters become spaces. */
export function botNameFrom(name: string): string {
  return name
    .replace(/[^A-Za-z0-9 ._-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 30)
    .trim();
}

/** {agentEmail} and {agentName} (the part before the @) for the add-an-agent steps. */
export function agentFills(email: string): Record<string, string> {
  const at = email.trim();
  return AGENT_EMAIL.test(at) ? { agentEmail: at, agentName: at.split("@")[0] } : {};
}

/** The add-an-agent steps for the apps ticked on a new key. */
export function agentSteps(provider: Provider, on: readonly string[]): AgentStep[] {
  return provider.addAgent.filter((s) => s.apps.some((a) => on.includes(a)));
}

/** What a new key's agent needs outside the service for the apps ticked on it: its email where an
 *  app opens the agent's own address (its mailbox), each provider's steps for those apps, and
 *  the apps that give each agent its own bot. Apps with nothing per agent add nothing. */
export function newKeySetup(ticked: Connector[], providers: Provider[]) {
  const ids = ticked.map((c) => c.id);
  return {
    addressed: ticked.filter((c) => c.agentAccess),
    steps: providers.flatMap((provider) =>
      agentSteps(provider, ids).map((step) => ({ provider: provider.id, step })),
    ),
    bots: ticked.filter((c) => c.signsInAs === "agentBot"),
  };
}

/** The list limits a new key must be given for an app before it is made (e.g. SharePoint's
 *  sites): required, and not the agent's own address, which comes from its email. */
export function newKeyLimits(c: Connector): Connector["limits"] {
  return c.limits.filter((l) => l.kind === "list" && l.required);
}

/** "App: Limit" for each required list limit still empty on a new key's ticked apps. */
export function missingNewKeyLimits(ticked: Connector[], chosen: Record<string, Limits>): string[] {
  return ticked.flatMap((c) =>
    newKeyLimits(c)
      .filter((l) => {
        const v = chosen[c.id]?.[l.id];
        return !Array.isArray(v) || v.length === 0;
      })
      .map((l) => `${c.name}: ${l.label}`),
  );
}

/** The grants a new key starts with: each ticked app at what its page allows (its ceiling),
 *  with the limits chosen on the form, and the agent's email as its address where the app
 *  has one. */
export function newKeyGrants(
  ticked: { connector: Connector; ceiling: string[] }[],
  email: string,
  chosen: Record<string, Limits>,
): { app: string; capabilities: string[]; limits: Limits }[] {
  return ticked.map(({ connector: c, ceiling }) => ({
    app: c.id,
    capabilities: ceiling,
    limits: {
      ...Object.fromEntries(newKeyLimits(c).map((l) => [l.id, chosen[c.id]?.[l.id] ?? []])),
      ...(c.agentAccess ? { [c.agentAccess.limit]: email.trim() } : {}),
    },
  }));
}

/** The keys whose `limit` (an agent's own address) is `value`, in any app. */
export function keysWithAddress(agents: Agent[], connectors: Connector[], value: string): Agent[] {
  const wanted = value.trim().toLowerCase();
  return agents.filter((a) =>
    a.grants.some((g) => {
      const limit = connectors.find((c) => c.id === g.connectorId)?.agentAccess?.limit;
      const held = limit ? g.limits[limit] : undefined;
      return typeof held === "string" && held.toLowerCase() === wanted;
    }),
  );
}

/** The warning that active keys already hold an address, or null. */
export function holdersNotice(holders: Agent[]): string | null {
  const active = holders.filter((a) => a.status === "active").map((a) => a.name);
  if (!active.length) return null;
  return `${active.join(", ")} already ${active.length === 1 ? "uses" : "use"} this address.`;
}

/** A numbered setup step: done (checked live), failing, not done yet, or marked done by the
 *  admin. `confirm`: the service cannot check it, so the admin marks it. `locked`: an earlier step
 *  it needs is not done yet. */
export type StepStatus = {
  /** checking: its live check has not answered yet (the page shows before it does). */
  state: "done" | "failing" | "todo" | "marked" | "checking";
  confirm: boolean;
  detail: string;
  locked: boolean;
  /** Why it is locked, when that is more than the step before it. */
  why?: string;
};

const status = (
  state: StepStatus["state"],
  detail: string,
  extra: Partial<StepStatus> = {},
): StepStatus => ({ state, detail, confirm: false, locked: false, ...extra });

/** Connecting a provider, before any app: its setup steps (registering its app, which only the
 *  sign-in can prove, so the admin marks it; the certificate, made here, its upload proven by
 *  the sign-in), then connecting, a live sign-in. Each step opens once the one before it
 *  passes. `marks`: the steps the admin marked done; `secret`: the admin chose a client secret
 *  over the certificate. */
export function providerSteps(
  setup: Step[],
  connection: ProviderConnection | null,
  marks: string[],
  secret: boolean,
): StepStatus[] {
  const signedIn = connection?.status === "connected";
  const certificate = Boolean(connection?.config.certificate);
  const proven = `Proven by the sign-in in step ${setup.length + 1}.`;
  const out: StepStatus[] = [];
  let open = true;
  for (const step of setup) {
    let next: StepStatus;
    let passed = true;
    if (signedIn)
      next = status(
        "done",
        step.certificate && connection.config.auth === "client_secret"
          ? "Signs in with a client secret instead."
          : proven,
      );
    else if (step.certificate) {
      passed = certificate || secret;
      next = status(
        "todo",
        certificate
          ? "Certificate created. Upload it, then connect: the sign-in proves it."
          : secret
            ? "Using a client secret instead: enter it when you connect."
            : "Create the certificate here, then upload it to the app.",
      );
    } else if (step.confirm) {
      // Making the certificate came after registering, so that counts as the mark.
      passed = marks.includes(step.confirm) || certificate;
      next = status(
        passed ? "marked" : "todo",
        passed
          ? "Marked done. The sign-in proves it."
          : "The app can be seen here only once you connect. Mark it done to go on.",
        { confirm: true },
      );
    } else next = status("todo", "");
    out.push({ ...next, locked: !open });
    open = open && passed;
  }
  out.push(
    status(
      signedIn ? "done" : connection?.status === "needs_reconnect" ? "failing" : "todo",
      signedIn
        ? "Signs in."
        : connection?.status === "needs_reconnect"
          ? "The last sign-in failed."
          : "Not connected yet.",
      { locked: !open },
    ),
  );
  return out;
}

/** One numbered step on a provider app's page, in order: what agents may do, the app's setup
 *  steps (all before turning on: they need only the provider's app ids), the Claude key a
 *  ticked capability needs, turning it on, then what needs it on: its sites, and with
 *  the shared app the steps the service cannot check. Turning on stays locked until every
 *  step before it the admin confirms is marked, and reads as done only once every step is. */
export type AppStep = {
  id: string;
  kind: "capabilities" | "setup" | "turnOn" | "sites" | "claudeKey";
  title: string;
  step?: Step;
  status: StepStatus;
};

const settled = (s: StepStatus) => s.state === "done" || s.state === "marked";

const numbered = (ns: number[]) =>
  ns.length === 1
    ? `step ${ns[0]}`
    : `steps ${ns.slice(0, -1).join(", ")} and ${ns[ns.length - 1]}`;

export function appSteps(input: {
  connector: Connector;
  connection: Connection | null;
  /** The provider connection signs in, or the shared app is approved on turning on. */
  providerReady: boolean;
  /** Turning on approves the shared app, which asks for the app's permissions itself:
   *  only the steps the service cannot check stay, after turning on (they need the app's ids). */
  approval?: boolean;
  /** The app's live check, run when its page loads. */
  health: CheckOutcome | null;
  /** The ceiling as ticked on the page. */
  ceiling: string[];
  /** Each added site's live access, for an app with `siteAccess`. */
  sites: SiteAccess[] | null;
  claudeKey: boolean;
  /** The `confirm` ids the admin marked done. */
  marks: string[];
  /** The set-up-once steps' own live checks (`Step.live`), null when they could not run. */
  live?: Record<string, LiveCheck> | null;
  /** The live checks have not answered yet: the steps they decide show as checking. */
  checking?: boolean;
  /** Values saved at steps with `settings` (the provider connection's config). */
  saved?: Record<string, string>;
  /** The choice fields as typed or stored: a step with `option` shows only while chosen. */
  choices?: Record<string, string>;
}): AppStep[] {
  const { connector, connection, health, ceiling, marks } = input;
  const setup = chosenSteps(connector, connector.setup, input.choices ?? {}, ceiling);
  const afterConnect = chosenSteps(connector, connector.afterConnect, input.choices ?? {}, ceiling);
  const on = connection !== null;
  const changed = on && connection.capabilities.join(",") !== ceiling.join(",");
  const before = input.providerReady ? {} : { locked: true };
  const after = on ? {} : { locked: true };
  const checks = health?.ok ? health.checks : null;
  const permissions = checks?.find((c) => c.label === connector.verify.label);
  const confirmStep = (step: Step, extra: Partial<StepStatus>): StepStatus =>
    marks.includes(step.confirm!)
      ? status("marked", "You marked this done.", { confirm: true, ...extra })
      : status("todo", "This cannot be checked from here. Mark it done once it is.", {
          confirm: true,
          ...extra,
        });
  const checking = (extra: Partial<StepStatus>) => status("checking", "Checking…", extra);
  const ownCheck = (id: string, extra: Partial<StepStatus>): StepStatus => {
    if (input.checking) return checking(extra);
    const found = input.live?.[id];
    if (!found) return status("failing", "This could not be checked. Reload the page.", extra);
    return status(
      found.state === "done" ? "done" : found.state === "failing" ? "failing" : "todo",
      found.message,
      extra,
    );
  };
  const liveStep = (step: Step, extra: Partial<StepStatus>): StepStatus =>
    step.live
      ? ownCheck(step.live, extra)
      : step.confirm
        ? confirmStep(step, extra)
        : !on
          ? status("todo", "Checked live when you turn it on.", extra)
          : input.checking
            ? checking(extra)
            : health && !health.ok
              ? status("failing", `The check failed: ${health.error}`, extra)
              : permissions?.warning
                ? status("failing", permissions.warning, extra)
                : status("done", "Checked live: the app's permissions match.", extra);
  const setupStep = (step: Step, i: number, extra: Partial<StepStatus>, prefix: string) => ({
    id: `${prefix}-${i}`,
    kind: "setup" as const,
    title: step.title,
    step,
    status: liveStep(step, extra),
  });

  const steps: AppStep[] = [
    {
      id: "capabilities",
      kind: "capabilities",
      title: "Choose what agents may do",
      status: changed
        ? status("todo", "Changed. Save it.", before)
        : on
          ? status("done", "Saved.", before)
          : status("todo", "Saved when you turn it on.", before),
    },
    ...(input.approval ? [] : [...setup, ...afterConnect]).map((s, i) =>
      setupStep(s, i, before, "setup"),
    ),
  ];
  const needsKey = connector.capabilities.filter((c) => c.claudeKey && ceiling.includes(c.id));
  if (needsKey.length)
    steps.push({
      id: "claude-key",
      kind: "claudeKey",
      title: "Add your organization's Claude API key",
      status: input.checking
        ? checking(before)
        : input.claudeKey
          ? status("done", "Set.", before)
          : status(
              "todo",
              `Needed because ${needsKey.map((c) => c.label).join(" and ")} is ticked: each answer uses this key.`,
              before,
            ),
    });

  const later: AppStep[] = input.approval
    ? [...setup.filter((s) => s.confirm), ...afterConnect].map((s, i) =>
        setupStep(s, i, after, "after"),
      )
    : [];
  if (connector.siteAccess) {
    const sites = input.sites ?? [];
    const open = sites.filter((s) => s.status === "granted").length;
    later.push({
      id: "sites",
      kind: "sites",
      title: "Add the sites agents may use",
      status:
        on && input.checking
          ? checking(after)
          : sites.length === 0
            ? status("todo", "No sites yet.", after)
            : open === sites.length
              ? status("done", `${many(sites.length, "site")}, all open to the app.`, after)
              : status("todo", `${open} of ${many(sites.length, "site")} open to the app.`, after),
    });
  }

  const turnOnAt = steps.length + 1;
  // Confirm steps before turning on hold it back; live ones are checked by turning on.
  const waiting = steps.flatMap((s, i) => (s.status.confirm && !settled(s.status) ? [i + 1] : []));
  const unfinished = [
    ...steps.flatMap((s, i) => (s.kind !== "capabilities" && !settled(s.status) ? [i + 1] : [])),
    ...later.flatMap((s, i) => (settled(s.status) ? [] : [turnOnAt + 1 + i])),
  ];
  const turnOn: StepStatus = !on
    ? waiting.length && input.providerReady
      ? status("todo", "Not on yet.", {
          locked: true,
          why: `Finish ${numbered(waiting)} first.`,
        })
      : status("todo", "Not on yet.", before)
    : connection.status !== "connected"
      ? status("failing", "Needs reconnecting.")
      : input.checking
        ? status("checking", "On. Checking…")
        : health && !health.ok
          ? status("failing", `The check failed: ${health.error}`)
          : permissions?.warning
            ? status("failing", `On, but the permission check fails: ${permissions.warning}`)
            : unfinished.length
              ? status(
                  "todo",
                  `On, but ${numbered(unfinished)} ${unfinished.length === 1 ? "is" : "are"} not done yet.`,
                )
              : status("done", "On. Every step is done and the check passed.");
  return waitForSettings(
    [
      ...steps,
      { id: "turn-on", kind: "turnOn", title: `Turn on ${connector.name}`, status: turnOn },
      ...later,
    ],
    input.saved ?? {},
  );
}

/** A step whose text fills in a value saved at another step waits until it is saved: its
 *  words would show the raw {name} and its commands would not run. */
export function waitForSettings(steps: AppStep[], saved: Record<string, string>): AppStep[] {
  const at = new Map<string, { n: number; label: string }>();
  steps.forEach((s, i) =>
    s.step?.settings?.forEach((x) => at.set(x.id, { n: i + 1, label: x.label })),
  );
  return steps.map((s) => {
    if (!s.step || s.step.settings) return s;
    const text = [s.step.body, ...(s.step.click ?? []), ...(s.step.copy ?? [])].join("\n");
    const missing = [...at].filter(([id]) => text.includes(`{${id}}`) && !saved[id]?.trim());
    if (!missing.length) return s;
    const labels = missing.map(([, x]) => x.label);
    const where = [...new Set(missing.map(([, x]) => x.n))];
    const what =
      labels.length === 1 ? labels[0] : `${labels.slice(0, -1).join(", ")} and ${labels.at(-1)}`;
    return {
      ...s,
      status: status("todo", `Waiting for ${what}.`, {
        locked: true,
        why: `Save the ${what} in ${numbered(where)} first.`,
      }),
    };
  });
}

/** The status an app's page shows by its name: needing steps while it is on but any step is
 *  not done. */
export function appPageStatus(
  connection: Connection | null,
  healthy: boolean,
  steps: AppStep[],
): AppStatus {
  const base = appStatus(connection, healthy);
  if (base !== "connected") return base;
  return steps.every((s) => s.kind === "capabilities" || settled(s.status))
    ? "connected"
    : "needs_steps";
}

/** What an app's steps are checked against, read live and read only when a page loads
 *  (`appStates` in agent-connect.ts). `health` or `sites` null: that check could not run. */
export type AppLive = {
  health: CheckOutcome | null;
  sites: SiteAccess[] | null;
  claudeKey: boolean;
  live: Record<string, LiveCheck> | null;
  /** Not read yet: the page shows before its live checks answer (`CHECKING`). */
  checking?: true;
};

/** An app's live state before its checks answer: steps they decide show as checking, and its
 *  status is the stored one. */
export const CHECKING: AppLive = {
  health: null,
  sites: null,
  claudeKey: false,
  live: null,
  checking: true,
};

/** How a provider's app signs in: through the provider connection, or (nothing connected yet,
 *  where the shared app is offered) by approving that app when turning it on. */
export function signInRoute(definition: Provider, provider: ProviderConnection | null) {
  const signsIn = providerReady(provider);
  return { signsIn, approve: !signsIn && definition.approval };
}

/** A provider app's steps and status: the one computation behind its page's badge, its row on
 *  the provider's page and the provider's header and card, so none contradicts another. A
 *  check that could not run needs attention. `ceiling` and `choices`: as ticked and typed on
 *  the app's page, else as saved. */
export function appState(input: {
  definition: Provider;
  provider: ProviderConnection | null;
  connector: Connector;
  connection: Connection | null;
  state: AppLive;
  marks: string[];
  ceiling?: string[];
  choices?: Record<string, string>;
}): { steps: AppStep[]; status: AppStatus } {
  const { connector, connection, state } = input;
  const { signsIn, approve } = signInRoute(input.definition, input.provider);
  const steps = appSteps({
    connector,
    connection,
    providerReady: signsIn || approve,
    approval: approve,
    health: state.health,
    ceiling:
      input.ceiling ??
      connection?.capabilities ??
      connector.capabilities.filter((c) => c.required).map((c) => c.id),
    sites: state.sites,
    claudeKey: state.claudeKey,
    marks: input.marks,
    live: state.live,
    checking: state.checking,
    saved: input.provider?.config ?? {},
    choices: input.choices ?? connection?.config ?? {},
  });
  const healthy = input.provider?.status === "connected" && state.health?.ok !== false;
  if (state.checking) return { steps, status: appStatus(connection, healthy) };
  const unknown = state.health === null || (Boolean(connector.siteAccess) && state.sites === null);
  const status =
    unknown && appStatus(connection, healthy) === "connected"
      ? "needs_attention"
      : appPageStatus(connection, healthy, steps);
  return { steps, status };
}

const many = (n: number, noun: string) => `${n} ${noun}${n === 1 ? "" : "s"}`;

/** A key's own addresses (its mailbox), per app that has one, e.g. for its access checks. */
export function agentAddresses(
  agent: Agent,
  connectors: Connector[],
): { connector: Connector; value: string }[] {
  return connectors.flatMap((connector) => {
    const limit = connector.agentAccess?.limit;
    const held = agent.grants.find((g) => g.connectorId === connector.id)?.limits[limit ?? ""];
    return limit && typeof held === "string" && held ? [{ connector, value: held }] : [];
  });
}
