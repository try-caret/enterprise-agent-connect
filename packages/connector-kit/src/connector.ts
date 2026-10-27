import { existsSync, readdirSync, readFileSync } from "node:fs";
import { z } from "zod";
import type { MicrosoftClient } from "../../../connectors/microsoft/client.js";
import { fillBrandDeep } from "./branding.js";
import type { Toolset } from "./common.js";
import type { Vendor } from "./types.js";

// A connector describes one app agents can use: how an admin connects it, the capabilities
// they can grant (each a set of tools), the value limits a grant can carry, and how a
// connection is checked. Each app is a folder in connectors/: its connector.json, and the
// ConnectorModule (index.ts) with its vendor client and tools.

const text = z.string().min(1);
const slug = z.string().regex(/^[a-z][a-z0-9_]*$/);

const settingSchema = z
  .object({
    id: z.string().regex(/^[a-z][A-Za-z]+$/),
    label: text,
    help: text.optional(),
    placeholder: text.optional(),
    // The whole value must match.
    pattern: text,
  })
  .strict();
export type Setting = z.infer<typeof settingSchema>;

// One option of a choice field (`options`), e.g. add_to_teams = on.
const chosenSchema = z.object({ field: slug, value: slug }).strict();

const stepSchema = z
  .object({
    title: text,
    body: text,
    link: z.object({ label: text, href: z.url() }).strict().optional(),
    // A line containing {scope} repeats once per scope; {scopes} becomes them space-separated.
    copy: z.array(text).optional(),
    // The shell `copy` runs in, named above the commands.
    shell: z.enum(["bash", "powershell"]).optional(),
    // Lists the in-app permissions the ticked capabilities need, e.g. Jira project permissions.
    permissions: z.boolean().optional(),
    // Shows the certificate the service made for the provider connection, to download.
    certificate: z.boolean().optional(),
    // The same step in the vendor's admin portals, one screen or choice per item. A step with
    // both `click` and `copy` is offered both ways.
    click: z.array(text).min(1).optional(),
    // The vendor has no portal page for this step, only commands (`copy`): the admin page says
    // so in place of clicks. Never with `click`, so a step cannot pass as clickable by a
    // "use the commands" line.
    clickUnavailable: z.literal(true).optional(),
    // A site grant step shown only while the site role the ceiling needs is this one, e.g.
    // raising a read grant to write.
    siteRole: z.enum(["read", "write"]).optional(),
    // Shows the live check of the agent's own address (`agentAccess`) after this step.
    agentCheck: z.boolean().optional(),
    // The service cannot check this step live, so the admin marks it done. Steps of different apps
    // with the same id share that mark (e.g. one group both Outlook apps use).
    confirm: z
      .string()
      .regex(/^[a-z0-9-]+$/)
      .optional(),
    // Checked live by the app's own setup check under this id (e.g. Teams' Azure role), not by
    // the permissions check.
    live: z
      .string()
      .regex(/^[a-z0-9-]+$/)
      .optional(),
    // Values the admin enters at this step, kept in the provider connection's config under
    // their ids and filled into steps as {id} (e.g. the Azure subscription for Teams bots).
    settings: z.array(settingSchema).min(1).optional(),
    // Shown, and holding back turning on, only while this option is chosen (e.g. Teams' Azure
    // steps only while the service makes the bots).
    option: chosenSchema.optional(),
  })
  .strict();

// A step of adding one more agent on a provider's page, shown while one of `apps` is on.
const agentStepSchema = stepSchema
  .extend({ apps: z.array(z.string().regex(/^[a-z0-9-]+$/)).min(1) })
  .strict();

const clickRule = [
  (s: { click?: string[]; copy?: string[]; clickUnavailable?: true }) =>
    !s.clickUnavailable || (!s.click && Boolean(s.copy)),
  { message: "A step with clickUnavailable has commands (copy) and no click" },
] as const;
const step = stepSchema.refine(...clickRule);
const agentStep = agentStepSchema.refine(...clickRule);

// A field with `options` is a choice (the first is the default), stored in the app's config
// under its id. An option's `scopes` are vendor permissions the app then needs too, e.g.
// reading every site.
const optionSchema = z
  .object({ value: slug, label: text, description: text, scopes: z.array(text).optional() })
  .strict();

const fieldSchema = z
  .object({
    id: slug,
    label: text,
    help: text.optional(),
    placeholder: text.optional(),
    secret: z.boolean(),
    options: z.array(optionSchema).min(2).optional(),
    // Asked only while an earlier choice is this option; otherwise its first option holds.
    requires: chosenSchema.optional(),
  })
  .strict();

// "map": one optional value per value of the list limit `keys`, e.g. a folder per site.
const limitSchema = z
  .object({
    id: slug,
    label: text,
    help: text,
    kind: z.enum(["list", "text", "map"]),
    required: z.boolean(),
    // Each value must match this, e.g. a Jira project key or a mailbox address.
    pattern: z.string().optional(),
    keys: slug.optional(),
  })
  .strict()
  .refine((l) => (l.kind === "map") === Boolean(l.keys), {
    message: "A map limit names the list limit it is keyed by (keys), and only it",
  });

const httpProbeSchema = z
  .object({
    method: z.enum(["GET", "POST"]),
    path: z.string().startsWith("/"),
    label: text,
    show: text,
    // Keeps only list items whose fields equal these values, e.g. objects agents can query.
    match: z.record(z.string(), z.union([z.string(), z.number(), z.boolean()])).optional(),
    body: z.record(z.string(), z.unknown()).optional(),
    // Shown to the admin when the probe finds nothing, e.g. a missing project permission.
    ifEmpty: text.optional(),
  })
  .strict();

// A check no single API call answers, e.g. the permissions inside a token. The vendor client
// implements it under this name (Vendor.checks).
const vendorProbeSchema = z
  .object({
    label: text,
    check: slug,
    ifEmpty: text.optional(),
  })
  .strict();

const probeSchema = z.union([httpProbeSchema, vendorProbeSchema]);
export type Probe = z.infer<typeof probeSchema>;

// An app described entirely here, e.g. a company's internal REST API: the generic http module
// (connectors/http) runs it, so it needs no code. Auth values are the ids of connect form fields.
const httpAuthSchema = z.discriminatedUnion("type", [
  z
    .object({ type: z.literal("header"), header: z.string().regex(/^[A-Za-z0-9-]+$/), value: slug })
    .strict(),
  z.object({ type: z.literal("bearer"), token: slug }).strict(),
  z.object({ type: z.literal("basic"), username: slug, password: slug }).strict(),
  z
    .object({
      type: z.literal("oauth2"),
      tokenUrl: z.url({ protocol: /^https?$/ }),
      clientId: slug,
      clientSecret: slug,
    })
    .strict(),
]);

const httpParamSchema = z
  .object({
    type: z.enum(["string", "number", "integer", "boolean"]),
    description: text,
    // Where it goes; a {name} in the path template is always "path".
    in: z.enum(["path", "query", "body"]),
    required: z.boolean().default(true),
    // The value must be one of the agent's values of this list limit, while it has any.
    limit: slug.optional(),
  })
  .strict();

const httpToolSchema = z
  .object({
    name: slug,
    description: text,
    method: z.enum(["GET", "POST", "PUT", "PATCH", "DELETE"]),
    // Under baseUrl, with {param} for path params, e.g. /orders/{id}.
    path: z.string().regex(/^\/[^?#]*$/),
    params: z.record(slug, httpParamSchema).default({}),
    // Deletes or overwrites: its capability must be destructive. DELETE always is.
    destructive: z.boolean().default(false),
    // Checked in the answer, for calls whose params cannot be limited (e.g. an order by id):
    // while the agent has values of `limit`, the answer's `field` (a dot path; every item of a
    // list) must be one of them, else nothing is returned. A missing field is refused.
    answerLimits: z
      .array(
        z
          .object({ field: z.string().regex(/^[A-Za-z0-9_]+(\.[A-Za-z0-9_]+)*$/), limit: slug })
          .strict(),
      )
      .default([]),
  })
  .strict()
  .refine(
    (t) => {
      const inPath = [...t.path.matchAll(/\{([a-z][a-z0-9_]*)\}/g)].map((m) => m[1]);
      const declared = Object.entries(t.params).filter(([, p]) => p.in === "path");
      return (
        inPath.every((n) => t.params[n]?.in === "path" && t.params[n].required) &&
        declared.every(([n]) => inPath.includes(n))
      );
    },
    { message: "Each {param} in a tool's path is a required path param, and only those" },
  )
  .refine((t) => !Object.values(t.params).some((p) => p.in === "body") || t.method !== "GET", {
    message: "A GET tool has no body params",
  })
  // Checked after the call, so only for reads: a write would already have happened.
  .refine((t) => !t.answerLimits.length || t.method === "GET", {
    message: "answerLimits are for GET tools only",
  });

const httpSchema = z
  .object({
    // Every call goes to this origin, under this path, and nowhere else: redirects are refused.
    baseUrl: z.url({ protocol: /^https?$/ }).refine((u) => !/[?#]/.test(u), "No query or fragment"),
    auth: httpAuthSchema,
    tools: z.array(httpToolSchema).min(1),
    // Plain http for baseUrl and the token URL, e.g. an app inside the cluster. Off by default.
    allowInsecureHttp: z.boolean().default(false),
  })
  .strict()
  .refine(
    (h) =>
      h.allowInsecureHttp ||
      [h.baseUrl, ...(h.auth.type === "oauth2" ? [h.auth.tokenUrl] : [])].every((u) =>
        u.startsWith("https://"),
      ),
    { message: "baseUrl and tokenUrl must be https unless allowInsecureHttp is true" },
  );
export type HttpApp = z.infer<typeof httpSchema>;

export const connectorSchema = z
  .object({
    id: z.string().regex(/^[a-z0-9-]+$/),
    name: text,
    description: text,
    // A path under the admin app's public folder.
    logo: z.string().regex(/^\/brand\/[a-z0-9/-]+\.svg$/),
    identity: text,
    // "agentBot": calls go out as the agent's own identity (its Teams bot), not the provider's
    // shared sign-in. A grant only gives tools once that agent has a connected bot.
    signsInAs: z.enum(["agentBot"]).optional(),
    // A provider connection (providers/<id>.json) this app signs in through, shared with the
    // provider's other apps: its credentials are entered once, not per app.
    provider: z
      .string()
      .regex(/^[a-z0-9-]+$/)
      .optional(),
    // Position among its provider's apps on the Integrations tab, where they are grouped.
    order: z.number().int().optional(),
    connectionKind: z.enum(["guided"]),
    // The connection config value naming the account reached (site, org, tenant). Agents'
    // grants are cleared when a reconnect reaches a different one.
    accountKey: z.string().regex(/^[A-Za-z]+$/),
    // Vendor scopes every ceiling needs, whatever is ticked (e.g. what adding agents' Teams
    // bots takes), filled into {scope}/{scopes} with the capabilities' own.
    scopes: z.array(text).optional(),
    setup: z.array(step),
    afterConnect: z.array(step),
    fields: z.array(fieldSchema),
    capabilities: z
      .array(
        z
          .object({
            id: slug,
            label: text,
            description: text,
            // Always on: every connection and every agent has it.
            required: z.boolean().default(false),
            // Off unless the admin confirms a warning.
            destructive: z.boolean().default(false),
            // Spends money on the account (e.g. buys shipping labels): shown as that instead
            // of "Deletes data". Destructive too, so it still asks first.
            costs: z.boolean().optional(),
            // Vendor scopes or roles, filled into {scope}/{scopes} in setup steps.
            scopes: z.array(text),
            // The permissions a provider's access token must carry for this capability, when
            // they are not `scopes` (e.g. mail scopes are Exchange roles, not token roles).
            roles: z.array(text).optional(),
            // Permissions granted inside the app, listed by steps with `permissions`.
            permissions: z.array(text),
            // Values for {name} placeholders in setup steps while this capability is ticked;
            // a later ticked capability overrides an earlier one (e.g. site role read → write).
            fill: z.record(z.string().regex(/^[a-z_]+$/), text).optional(),
            // A caution shown with the checkbox, e.g. that it reaches the whole organization.
            note: text.optional(),
            // Added to the agent's error when the app answers 403 to one of its tools: what
            // Microsoft (or the vendor) still needs, with {scopes} filled.
            forbidden: text.optional(),
            // Its answers use the organization's Claude API key, a setup step while ticked.
            claudeKey: z.boolean().optional(),
            tools: z.array(slug).min(1),
          })
          .strict()
          .refine((c) => !(c.required && c.destructive), {
            message: "A required capability cannot be destructive",
          })
          .refine((c) => !c.costs || c.destructive, {
            message: "A capability that spends money must be destructive, so it asks first",
          }),
      )
      .min(1),
    limits: z.array(limitSchema),
    // Sites the admin adds one at a time after connecting (config.sites), each given to the
    // app outside the service. `grant`: steps showing how, with {site}, {siteHost} and {sitePath}
    // filled per site; `request`: a note to send IT asking for it.
    // `write`: the capabilities that need write access to each site, so Check access also
    // proves the app can change files there while one of them is in the ceiling.
    siteAccess: z
      .object({
        grant: z.array(step).min(1),
        request: text,
        write: z.array(slug).default([]),
      })
      .strict()
      .optional(),
    // Each agent's own address in this app (e.g. its mailbox), held in the text limit `limit`.
    // `path` (with {value}) is read with the app's sign-in to show whether the app can reach
    // it yet; `label` names that check; the messages answer 2xx, 401/403, 404.
    agentAccess: z
      .object({
        limit: slug,
        label: text,
        path: z.string().startsWith("/").includes("{value}"),
        granted: text,
        notGranted: text,
        notFound: text,
        // Per capability: a call that proves the app may also do that there, changing
        // nothing (Exchange refuses a missing role with 401/403 before reading the body).
        // `allowed`: the statuses that mean it may. `missing` names those it may not yet,
        // {missing} filled with their labels and roles.
        probes: z
          .record(
            slug,
            z
              .object({
                method: z.enum(["GET", "POST", "PATCH"]),
                path: z.string().startsWith("/").includes("{value}"),
                body: z.unknown().optional(),
                allowed: z.array(z.number().int()).min(1),
              })
              .strict(),
          )
          .default({}),
        missing: text.optional(),
      })
      .strict()
      .optional(),
    verify: probeSchema,
    discover: probeSchema.optional(),
    revoke: text,
    http: httpSchema.optional(),
  })
  .strict()
  .refine(
    (c) => {
      if (!c.http) return true;
      const field = (id: string, secret: boolean) =>
        c.fields.some((f) => f.id === id && f.secret === secret && !f.options);
      const a = c.http.auth;
      const secretsOk =
        a.type === "header"
          ? field(a.value, true)
          : a.type === "bearer"
            ? field(a.token, true)
            : a.type === "basic"
              ? field(a.username, false) && field(a.password, true)
              : field(a.clientId, false) && field(a.clientSecret, true);
      const lists = new Set(c.limits.filter((l) => l.kind === "list").map((l) => l.id));
      return (
        secretsOk &&
        !c.provider &&
        (c.accountKey === "host" || c.fields.some((f) => f.id === c.accountKey && !f.secret)) &&
        c.http.tools.every(
          (t) =>
            Object.values(t.params).every((p) => !p.limit || lists.has(p.limit)) &&
            t.answerLimits.every((a) => lists.has(a.limit)),
        )
      );
    },
    {
      message:
        "An http app's auth names its fields (secrets as secret fields), its accountKey is host or a plain field, a param's or answer's limit names a list limit, and it has no provider",
    },
  )
  .refine(
    (c) =>
      !c.agentAccess || c.limits.some((l) => l.id === c.agentAccess!.limit && l.kind === "text"),
    { message: "agentAccess.limit must be one of the connector's text limits" },
  )
  .refine(
    (c) =>
      [...Object.keys(c.agentAccess?.probes ?? {}), ...(c.siteAccess?.write ?? [])].every((id) =>
        c.capabilities.some((x) => x.id === id),
      ) &&
      (!Object.keys(c.agentAccess?.probes ?? {}).length || c.agentAccess?.missing !== undefined),
    { message: "agentAccess probes and siteAccess.write name capabilities; probes need missing" },
  )
  .refine((c) => new Set(c.fields.map((f) => f.id)).size === c.fields.length, {
    message: "Field ids must be unique",
  })
  .refine((c) => new Set(c.capabilities.map((l) => l.id)).size === c.capabilities.length, {
    message: "Capability ids must be unique",
  })
  .refine(
    (c) => {
      const tools = c.capabilities.flatMap((l) => l.tools);
      return new Set(tools).size === tools.length;
    },
    { message: "A tool belongs to one capability only" },
  )
  .refine(
    (c) => c.provider !== undefined || ![...c.setup, ...c.afterConnect].some((s) => s.settings),
    { message: "Step settings are kept on a provider connection: only a provider's app has them" },
  )
  .refine(
    (c) => {
      const isOption = (o: { field: string; value: string }, fields: typeof c.fields) =>
        fields.some((f) => f.id === o.field && f.options?.some((x) => x.value === o.value));
      return (
        [...c.setup, ...c.afterConnect].every((s) => !s.option || isOption(s.option, c.fields)) &&
        c.fields.every(
          (f, i) => !f.requires || (f.options && isOption(f.requires, c.fields.slice(0, i))),
        )
      );
    },
    { message: "A step's option and a field's requires name an option of an earlier choice" },
  )
  .refine((c) => c.provider !== undefined || c.fields.length > 0, {
    message: "A guided connector needs connect form fields",
  })
  .refine(
    (c) =>
      c.limits.every(
        (l, i) => !l.keys || c.limits.slice(0, i).some((k) => k.id === l.keys && k.kind === "list"),
      ),
    { message: "A map limit's keys must be a list limit listed before it" },
  );

export type Connector = z.infer<typeof connectorSchema>;
export type Capability = Connector["capabilities"][number];

// One sign-in shared by several apps, e.g. a customer's Microsoft Entra app used by mail and
// files. Its setup steps are shown with the first of its apps the admin connects.
export const providerSchema = z
  .object({
    id: z.string().regex(/^[a-z0-9-]+$/),
    name: text,
    logo: z.string().regex(/^\/brand\/[a-z0-9/-]+\.svg$/),
    identity: text,
    // `identity` while the connection uses the deployment's shared app instead of the customer's own.
    sharedIdentity: text.optional(),
    accountKey: z.string().regex(/^[A-Za-z]+$/),
    // Config values naming who signs in (e.g. the app and its service principal). A change in
    // any, like a change of account, disconnects the provider's apps: what the admin granted
    // outside the service (mailbox roles, site grants) belongs to the old one.
    principalKeys: z.array(z.string().regex(/^[A-Za-z]+$/)).default([]),
    // The service can make a key pair whose certificate the admin uploads to their app.
    certificate: z.boolean(),
    setup: z.array(step),
    fields: z.array(fieldSchema.extend({ optional: z.boolean().default(false) }).strict()).min(1),
    // What one more agent needs outside the service (e.g. its mailbox), shown when its API key is
    // made with one of the apps each step names.
    addAgent: z.array(agentStep).default([]),
    revoke: text,
  })
  .strict();

export type Provider = z.infer<typeof providerSchema>;

/** The code behind one connector folder: its vendor client and its tools. */
export type ConnectorModule = {
  id: string;
  create(ctx: { microsoft: MicrosoftClient; connector: Connector }): {
    vendor: Vendor;
    toolset: Toolset;
  };
};

/** The repo's connectors/ folder, beside packages/ (and so beside dist/packages/ once built). */
export const CONNECTORS_DIR = new URL("../../../connectors/", import.meta.url);

// Each subfolder of dir holding `file` is one definition, whose id is the folder's name.
// Subfolders may be links: Kubernetes mounts ConfigMap and Secret folders that way.
function loadDir<T>(dir: URL, file: string, schema: z.ZodType<T & { id: string }>): T[] {
  return readdirSync(dir)
    .filter((name) => existsSync(new URL(`${name}/${file}`, dir)))
    .sort()
    .map((name) => {
      const where = `${name}/${file}`;
      const raw = JSON.parse(readFileSync(new URL(where, dir), "utf8"));
      let filled: unknown;
      try {
        filled = fillBrandDeep(raw);
      } catch (error) {
        throw new Error(`${where}: ${(error as Error).message}`);
      }
      const parsed = schema.safeParse(filled);
      if (!parsed.success) throw new Error(`${where}: ${z.prettifyError(parsed.error)}`);
      if (parsed.data.id !== name) throw new Error(`${where}: id must match the folder name`);
      return parsed.data;
    });
}

export function loadConnectors(dir = CONNECTORS_DIR): Connector[] {
  return loadDir(dir, "connector.json", connectorSchema);
}

export function loadProviders(dir = CONNECTORS_DIR): Provider[] {
  return loadDir(dir, "provider.json", providerSchema);
}

/** The ids of an app's choice fields: what its connection keeps from the admin's form. */
export const choiceFields = (connector: Connector) =>
  connector.fields.filter((f) => f.options).map((f) => f.id);

/** The option chosen for each choice field: the stored one, else the first, which also holds
 *  while the earlier choice it `requires` is not made. */
export function chosenOptions(connector: Connector, config: Record<string, string> = {}) {
  const chosen: Record<string, string> = {};
  return connector.fields.flatMap((f) => {
    if (!f.options) return [];
    const asked = !f.requires || chosen[f.requires.field] === f.requires.value;
    const option = (asked && f.options.find((o) => o.value === config[f.id])) || f.options[0];
    chosen[f.id] = option.value;
    return [option];
  });
}

/** The permissions a provider's token needs for these capabilities and the chosen options. */
export function tokenRolesFor(
  connector: Connector,
  capabilities: readonly string[],
  config: Record<string, string> = {},
): string[] {
  return [
    ...new Set([
      ...(connector.scopes ?? []),
      ...connector.capabilities
        .filter((c) => capabilities.includes(c.id))
        .flatMap((c) => c.roles ?? c.scopes),
      ...chosenOptions(connector, config).flatMap((o) => o.scopes ?? []),
    ]),
  ];
}

// What an agent may actually do: its capabilities that the connection also allows.
export function effectiveCapabilities(
  connector: Connector,
  grant: readonly string[],
  ceiling: readonly string[],
): Capability[] {
  return connector.capabilities.filter((c) => grant.includes(c.id) && ceiling.includes(c.id));
}

// Keys of a provider connection's own config, which a setting saved beside them must not replace.
const RESERVED_SETTINGS = new Set(["tenantId", "clientId", "auth", "certificate", "sites"]);

/** The settings a provider's apps ask for in their steps, by id. */
export function providerSettings(connectors: Connector[], provider: string): Map<string, Setting> {
  const out = new Map<string, Setting>();
  for (const c of connectors)
    if (c.provider === provider)
      for (const s of [...c.setup, ...c.afterConnect])
        for (const setting of s.settings ?? []) {
          if (RESERVED_SETTINGS.has(setting.id))
            throw new Error(`${c.id}: setting id ${setting.id} is the connection's own`);
          out.set(setting.id, setting);
        }
  return out;
}

export function expandCopy(lines: string[], scopes: string[]): string {
  return lines
    .flatMap((line) =>
      line.includes("{scope}") ? scopes.map((s) => line.replaceAll("{scope}", s)) : [line],
    )
    .map((line) => line.replaceAll("{scopes}", scopes.join(" ")))
    .join("\n");
}

// Reads `a.b` / `values[].name` paths out of a JSON response for display.
export function pickDisplay(
  data: unknown,
  path: string,
  match: Record<string, unknown> = {},
): string[] {
  const matches = (v: unknown) =>
    Object.entries(match).every(
      ([k, want]) => v && typeof v === "object" && (v as Record<string, unknown>)[k] === want,
    );
  let values: unknown[] = [data];
  for (const part of path.split(".")) {
    const many = part.endsWith("[]");
    const key = many ? part.slice(0, -2) : part;
    values = values.flatMap((v) => {
      const next = v && typeof v === "object" ? (v as Record<string, unknown>)[key] : undefined;
      return many ? (Array.isArray(next) ? next.filter(matches) : []) : [next];
    });
  }
  return values.filter((v) => v !== undefined && v !== null).map(String);
}
