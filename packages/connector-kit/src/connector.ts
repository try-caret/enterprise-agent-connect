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
export const setupStep = stepSchema.refine(...clickRule);
const step = setupStep;
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
    // Asked only while an earlier choice is this option; otherwise a choice's first option
    // holds and a text field is not asked (e.g. a password only for that sign-in method).
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
const headerName = z.string().regex(/^[A-Za-z0-9-]+$/);
// Printable ASCII only: a header value can never carry a line break into the request.
const headerValue = z.string().regex(/^[\x20-\x7E]*$/);
const keyAuths = [
  z
    .object({
      type: z.literal("header"),
      header: headerName,
      value: slug,
      // Written before the value, e.g. "SSWS " (Okta) or "Token token=" (PagerDuty).
      prefix: headerValue.optional(),
    })
    .strict(),
  z.object({ type: z.literal("bearer"), token: slug }).strict(),
  z.object({ type: z.literal("basic"), username: slug, password: slug }).strict(),
] as const;
const httpAuthSchema = z.discriminatedUnion("type", [
  ...keyAuths,
  z
    .object({
      type: z.literal("oauth2"),
      tokenUrl: z.url({ protocol: /^https?$/ }),
      clientId: slug,
      clientSecret: slug,
      // Sent with the token request, e.g. https://<org>.crm.dynamics.com/.default.
      scope: z
        .string()
        .regex(/^[\x21-\x7e]+( [\x21-\x7e]+)*$/)
        .optional(),
    })
    .strict(),
  // Added to every call's query string, e.g. ?api_key=…
  z
    .object({
      type: z.literal("query"),
      param: z.string().regex(/^[A-Za-z0-9_.-]+$/),
      value: slug,
    })
    .strict(),
]);

export type HttpAuth = z.infer<typeof httpAuthSchema>;

// The connect form fields an auth names, each with whether it must be a secret field.
const authFields = (a: HttpAuth): [string, boolean][] =>
  a.type === "header" || a.type === "query"
    ? [[a.value, true]]
    : a.type === "bearer"
      ? [[a.token, true]]
      : a.type === "basic"
        ? [
            [a.username, false],
            [a.password, true],
          ]
        : [
            [a.clientId, false],
            [a.clientSecret, true],
          ];

const scalarType = z.enum(["string", "number", "integer", "boolean"]);
const scalar = z.union([z.string(), z.number(), z.boolean()]);
const fitsType = (type: string, v: unknown) =>
  type === "integer"
    ? Number.isInteger(v)
    : type === "number" || type === "string" || type === "boolean"
      ? typeof v === type
      : false;

const httpParamSchema = z
  .object({
    type: z.enum(["string", "number", "integer", "boolean", "array"]),
    description: text,
    // Where it goes; a {name} in the path template is always "path".
    in: z.enum(["path", "query", "body"]),
    // The name the app knows it by, when it is not this param's own (e.g. pageToken).
    name: z
      .string()
      .regex(/^[A-Za-z0-9_.\-[\]]+$/)
      .optional(),
    required: z.boolean().default(true),
    // Sent when the agent leaves the param out; such a param is never required.
    default: z.union([scalar, z.array(scalar)]).optional(),
    // The only values the agent may send (each item's, for an array).
    enum: z.array(scalar).min(1).optional(),
    // An array's item type.
    items: scalarType.optional(),
    // An array in the query: one `name=` per item, unless joined with this, e.g. ",".
    join: z.string().min(1).max(3).optional(),
    // The value must be one of the agent's values of this list limit, while it has any.
    limit: slug.optional(),
  })
  .strict()
  .refine((p) => (p.type === "array") === Boolean(p.items), {
    message: "An array param names its item type (items), and only an array param",
  })
  .refine((p) => !p.join || (p.type === "array" && p.in === "query"), {
    message: "join is for array params in the query",
  })
  .refine(
    (p) => {
      const itemType = p.items ?? p.type;
      const fits = (v: unknown) =>
        fitsType(itemType, v) && (!p.enum || p.enum.some((e) => e === v));
      return (
        (p.enum ?? []).every((v) => fitsType(itemType, v)) &&
        (p.default === undefined ||
          (p.type === "array"
            ? Array.isArray(p.default) && p.default.every(fits)
            : fits(p.default)))
      );
    },
    { message: "A param's enum and default match its type, and the default is one of its enum" },
  );

// Picks what the agent sees of a successful answer. `items`: a dot path to the list in the
// answer, e.g. "data.results". `fields`: dot paths kept of each item (of the answer itself,
// without `items`); a path through a list keeps that field of each of its items.
const dotPath = z.string().regex(/^[A-Za-z0-9_$@-]+(\.[A-Za-z0-9_$@-]+)*$/);
const httpResultSchema = z
  .object({ items: dotPath.optional(), fields: z.array(dotPath).min(1).optional() })
  .strict()
  .refine((r) => r.items || r.fields, { message: "A result names items, fields or both" });

// One page per call: `param` takes the cursor in, and the answer gives the next one from a
// dot path in its body, a header, or the rel="next" address of its Link header (the value of
// `param` there). The agent gets {items, next}.
const httpPagingSchema = z.union([
  z.object({ param: slug, body: dotPath }).strict(),
  z.object({ param: slug, header: headerName }).strict(),
  z.object({ param: slug, link: z.literal(true) }).strict(),
]);

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
    answerLimits: z.array(z.object({ field: dotPath, limit: slug }).strict()).default([]),
    result: httpResultSchema.optional(),
    paging: httpPagingSchema.optional(),
  })
  .strict()
  .refine(
    (t) => {
      const inPath = [...t.path.matchAll(/\{([a-z][a-z0-9_]*)\}/g)].map((m) => m[1]);
      const declared = Object.entries(t.params).filter(([, p]) => p.in === "path");
      return (
        inPath.every(
          (n) =>
            t.params[n]?.in === "path" &&
            t.params[n].required &&
            t.params[n].default === undefined &&
            t.params[n].type !== "array",
        ) && declared.every(([n]) => inPath.includes(n))
      );
    },
    {
      message:
        "Each {param} in a tool's path is a required path param with one value and no default, and only those",
    },
  )
  .refine((t) => !Object.values(t.params).some((p) => p.in === "body") || t.method !== "GET", {
    message: "A GET tool has no body params",
  })
  // Checked after the call, so only for reads: a write would already have happened.
  .refine((t) => !t.answerLimits.length || t.method === "GET", {
    message: "answerLimits are for GET tools only",
  })
  .refine(
    (t) => {
      const p = t.paging && t.params[t.paging.param];
      return !t.paging || (p && p.in !== "path" && p.type !== "array" && p.type !== "boolean");
    },
    { message: "paging.param is one of the tool's query or body params, a string or number" },
  )
  .refine(
    (t) => {
      const wire = Object.entries(t.params).map(([k, p]) => `${p.in}:${p.name ?? k}`);
      return new Set(wire).size === wire.length;
    },
    { message: "Two params of a tool are sent under the same name" },
  );

// Headers no `headers` entry may set: credentials go through `auth` (secret fields), and the
// rest the service sets itself.
const RESERVED_HEADERS = new Set([
  "authorization",
  "proxy-authorization",
  "cookie",
  "host",
  "content-type",
  "content-length",
  "transfer-encoding",
  "connection",
]);
const CREDENTIAL_LIKE = /auth|token|secret|key|password|session|cookie/i;

const httpSchema = z
  .object({
    // Every call goes to this origin, under this path, and nowhere else: redirects are refused.
    baseUrl: z.url({ protocol: /^https?$/ }).refine((u) => !/[?#]/.test(u), "No query or fragment"),
    auth: httpAuthSchema,
    // Sent with every call, e.g. an API version or Accept. Not secret: never credentials.
    headers: z.record(headerName, headerValue).default({}),
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
  )
  .refine(
    (h) =>
      Object.keys(h.headers).every(
        (name) =>
          !RESERVED_HEADERS.has(name.toLowerCase()) &&
          !CREDENTIAL_LIKE.test(name) &&
          !(h.auth.type === "header" && h.auth.header.toLowerCase() === name.toLowerCase()),
      ),
    {
      message:
        "headers are not secret, so they cannot carry credentials (use auth) or headers the service sets",
    },
  )
  .refine(
    (h) =>
      h.auth.type !== "query" ||
      h.tools.every((t) =>
        Object.entries(t.params).every(
          ([k, p]) => p.in !== "query" || (p.name ?? k) !== (h.auth as { param: string }).param,
        ),
      ),
    { message: "No tool param may be sent as the query auth param" },
  );
export type HttpApp = z.infer<typeof httpSchema>;

const answerLimitsSchema = z.array(z.object({ field: dotPath, limit: slug }).strict()).default([]);
const secure = (allowInsecure: boolean, ...urls: string[]) =>
  allowInsecure || urls.every((u) => u.startsWith("https://"));

// A SOAP service described here, run by connectors/soap. Each tool is one operation: its Body
// content is a template whose {param} placeholders are filled with XML-escaped values.
const soapParamSchema = z
  .object({
    type: z.enum(["string", "number", "integer", "boolean"]),
    description: text,
    // Left out, an optional param's element (`<x>{param}</x>`) is left out of the body.
    required: z.boolean().default(true),
    limit: slug.optional(),
  })
  .strict();

const PLACEHOLDER = /\{([a-z][a-z0-9_]*)\}/g;
const xmlName = z.string().regex(/^[A-Za-z_][\w.-]*$/);

const soapToolSchema = z
  .object({
    name: slug,
    description: text,
    operation: xmlName,
    // Sent as the SOAPAction header (1.1) or the content type's action (1.2).
    soapAction: z
      .string()
      .regex(/^[^"\r\n]*$/)
      .optional(),
    body: text,
    params: z.record(slug, soapParamSchema).default({}),
    // Dot path in the Body's content to the answer, default `<operation>Response`; ending `[]`
    // makes it a list even when the service answers one item.
    result: z
      .string()
      .regex(/^[A-Za-z_][\w.-]*(\.[A-Za-z_][\w.-]*)*(\[\])?$/)
      .optional(),
    // Only these fields of the answer (of each item, for a list) reach the agent.
    fields: z.array(xmlName).min(1).optional(),
    // Changes nothing (a lookup): only these can have answerLimits.
    readOnly: z.boolean().default(false),
    destructive: z.boolean().default(false),
    answerLimits: answerLimitsSchema,
  })
  .strict()
  .refine(
    (t) => {
      const used = [...t.body.matchAll(PLACEHOLDER)].map((m) => m[1]);
      return (
        new Set(used).size === used.length &&
        used.every((n) => n in t.params) &&
        Object.keys(t.params).every((n) => used.includes(n))
      );
    },
    { message: "Each param is a {param} placeholder in the tool's body, once, and only those" },
  )
  .refine(
    (t) =>
      Object.entries(t.params).every(
        ([n, p]) =>
          p.required || new RegExp(`<([\\w.:-]+)(\\s[^>]*)?>\\{${n}\\}</\\1>`).test(t.body),
      ),
    { message: "An optional param's placeholder is the whole content of an element" },
  )
  .refine((t) => !t.answerLimits.length || t.readOnly, {
    message: "answerLimits are for readOnly tools only",
  })
  .refine((t) => !(t.readOnly && t.destructive), {
    message: "A readOnly tool cannot be destructive",
  });

const soapSchema = z
  .object({
    endpoint: z.url({ protocol: /^https?$/ }).refine((u) => !u.includes("#"), "No fragment"),
    version: z.enum(["1.1", "1.2"]).default("1.1"),
    // Prefixes declared on the Envelope for the tools' bodies, e.g. {"hr": "urn:example:hr"}.
    namespaces: z
      .record(
        xmlName.refine((p) => !["soap", "wsse", "wsu", "xml", "xmlns"].includes(p), "Reserved"),
        z.string().regex(/^[^"<>&\s]+$/),
      )
      .default({}),
    auth: z.discriminatedUnion("type", [...keyAuths]).optional(),
    // WS-Security UsernameToken in the SOAP header, with or without other auth.
    wsSecurity: z
      .object({ username: slug, password: slug, digest: z.boolean().default(false) })
      .strict()
      .optional(),
    // The call the connector's `operation` check makes: a tool, its arguments, what to show.
    verify: z
      .object({
        tool: slug,
        args: z.record(slug, z.union([z.string(), z.number(), z.boolean()])).default({}),
        show: text,
      })
      .strict()
      .optional(),
    tools: z.array(soapToolSchema).min(1),
    allowInsecureHttp: z.boolean().default(false),
  })
  .strict()
  .refine((s) => s.auth || s.wsSecurity, { message: "A soap app needs auth or wsSecurity" })
  .refine((s) => secure(s.allowInsecureHttp, s.endpoint), {
    message: "endpoint must be https unless allowInsecureHttp is true",
  })
  .refine((s) => !s.verify || s.tools.some((t) => t.name === s.verify!.tool), {
    message: "verify.tool names one of the soap tools",
  });
export type SoapApp = z.infer<typeof soapSchema>;
export type SoapTool = SoapApp["tools"][number];

// An OData service (v2 or v4: SAP Gateway, Dynamics, SuccessFactors) described here, run by
// connectors/odata. Agents never write OData syntax: values are typed and made into literals.
const odataName = z.string().regex(/^[A-Za-z_][A-Za-z0-9_.]*$/);
const odataValueSchema = z
  .object({
    type: z.enum([
      "string",
      "integer",
      "number",
      "decimal",
      "boolean",
      "guid",
      "datetime",
      "datetimeoffset",
    ]),
    description: text,
    // The OData property (or function parameter) name, default the param's own. A field can
    // bind a lookup, e.g. Dynamics' parentaccountid@odata.bind.
    property: z
      .string()
      .regex(/^[A-Za-z_][A-Za-z0-9_.]*(@odata\.bind)?$/)
      .optional(),
    // Keys are always required; fields default to required on create, never on update.
    required: z.boolean().optional(),
    limit: slug.optional(),
  })
  .strict();
const odataFilterSchema = odataValueSchema
  .extend({
    op: z.enum(["eq", "ne", "gt", "ge", "lt", "le", "contains", "startswith"]).default("eq"),
  })
  .strict()
  .refine((f) => !f.limit || f.op === "eq", { message: "Only an eq filter can have a limit" })
  .refine((f) => !["contains", "startswith"].includes(f.op) || f.type === "string", {
    message: "contains and startswith filter strings",
  });

const ODATA_KINDS = {
  list: { on: "entitySet", parts: ["filters"] },
  get: { on: "entitySet", parts: ["key"] },
  create: { on: "entitySet", parts: ["fields"] },
  update: { on: "entitySet", parts: ["key", "fields"] },
  function: { on: "function", parts: ["params"] },
  action: { on: "function", parts: ["params"] },
} as const;

const odataToolSchema = z
  .object({
    name: slug,
    description: text,
    kind: z.enum(["list", "get", "create", "update", "function", "action"]),
    entitySet: odataName.optional(),
    // A function or action import's name.
    function: odataName.optional(),
    key: z.record(slug, odataValueSchema).default({}),
    filters: z.record(slug, odataFilterSchema).default({}),
    fields: z.record(slug, odataValueSchema).default({}),
    params: z.record(slug, odataValueSchema).default({}),
    // $select, and the only properties passed back.
    select: z.array(odataName).default([]),
    orderBy: z
      .string()
      .regex(/^[A-Za-z_][\w./]*( (asc|desc))?(,[A-Za-z_][\w./]*( (asc|desc))?)*$/)
      .optional(),
    // Most items per page: the agent's max_results is capped at it.
    top: z.number().int().min(1).max(1000).default(50),
    destructive: z.boolean().default(false),
    answerLimits: answerLimitsSchema,
  })
  .strict()
  .refine(
    (t) => {
      const kind = ODATA_KINDS[t.kind];
      const parts = ["key", "filters", "fields", "params"] as const;
      return (
        Boolean(t[kind.on]) &&
        !t[kind.on === "entitySet" ? "function" : "entitySet"] &&
        parts.every(
          (p) => (kind.parts as readonly string[]).includes(p) || !Object.keys(t[p]).length,
        ) &&
        (!(kind.parts as readonly string[]).includes("key") || Object.keys(t.key).length > 0) &&
        (t.kind !== "create" || Object.keys(t.fields).length > 0) &&
        (t.kind === "list" || !t.orderBy) &&
        (kind.on !== "function" || !t.select.length)
      );
    },
    {
      message:
        "A list names an entitySet and filters, get a key, create fields, update a key and fields, function and action a function and params; orderBy is for lists, select not for functions",
    },
  )
  .refine((t) => !["create", "action"].includes(t.kind) || !t.answerLimits.length, {
    message: "answerLimits are for reads and updates (checked before the change)",
  })
  .refine(
    (t) => {
      const names = [t.key, t.filters, t.fields, t.params].flatMap(Object.keys);
      return (
        new Set(names).size === names.length &&
        !names.some((n) => n === "cursor" || n === "max_results")
      );
    },
    { message: "Param names are unique in a tool, and cursor and max_results are taken" },
  )
  .refine(
    (t) =>
      !t.select.length ||
      Object.entries(t.filters).every(([n, f]) => !f.limit || t.select.includes(f.property ?? n)),
    { message: "A limited filter's property must be selected, so answers can be checked" },
  );

const odataSchema = z
  .object({
    // The service root, e.g. https://host/sap/opu/odata/sap/API_BUSINESS_PARTNER.
    serviceRoot: z
      .url({ protocol: /^https?$/ })
      .refine((u) => !/[?#]/.test(u), "No query or fragment"),
    version: z.enum(["2", "4"]),
    auth: httpAuthSchema,
    // Sent on every call, e.g. SAP's sap-client. Not secret: never credentials.
    headers: z
      .record(
        headerName.refine(
          (h) =>
            !RESERVED_HEADERS.has(h.toLowerCase()) &&
            !CREDENTIAL_LIKE.test(h) &&
            !/^(accept|if-match|x-http-method|odata-.*|(max)?dataserviceversion|prefer)$/i.test(h),
          "Set by the service, or a credential (use auth)",
        ),
        headerValue,
      )
      .default({}),
    // Writes first fetch an x-csrf-token (SAP Gateway).
    csrf: z.boolean().default(false),
    // "skip": $top and $skip; "server": the service's own pages (odata.maxpagesize, v4), e.g.
    // Dynamics, which has no $skip.
    paging: z.enum(["skip", "server"]).default("skip"),
    tools: z.array(odataToolSchema).min(1),
    allowInsecureHttp: z.boolean().default(false),
  })
  .strict()
  .refine(
    (o) =>
      secure(
        o.allowInsecureHttp,
        o.serviceRoot,
        ...(o.auth.type === "oauth2" ? [o.auth.tokenUrl] : []),
      ),
    { message: "serviceRoot and tokenUrl must be https unless allowInsecureHttp is true" },
  )
  .refine((o) => o.paging === "skip" || o.version === "4", {
    message: "Server paging is for OData v4",
  });
export type ODataApp = z.infer<typeof odataSchema>;
export type ODataTool = ODataApp["tools"][number];

// An app reached as a database, described entirely here: the generic sql module
// (connectors/sql) runs it. Each tool is one statement the company's developer wrote, with
// :name params bound by the driver, never pasted into the text.
const sqlParamSchema = z
  .object({
    type: z.enum(["string", "number", "integer", "boolean"]),
    description: text,
    // Left out, an optional param is bound as null.
    required: z.boolean().default(true),
    // The value must be one of the agent's values of this list limit, while it has any.
    limit: slug.optional(),
  })
  .strict();

const sqlToolSchema = z
  .object({
    name: slug,
    description: text,
    sql: text,
    params: z.record(slug, sqlParamSchema).default({}),
    // Changes data: run read-write and committed, and never in a required capability. Any
    // other tool runs in a read-only transaction that is rolled back.
    write: z.boolean().default(false),
    // Deletes or overwrites: its capability must be destructive.
    destructive: z.boolean().default(false),
    // While the agent has values of `limit`, every row's `column` must be one of them, else
    // nothing is returned (a write is rolled back). A missing column is refused.
    answerLimits: z.array(z.object({ column: text, limit: slug }).strict()).default([]),
  })
  .strict()
  .refine((t) => !t.destructive || t.write, { message: "A destructive tool is a write" });

const sqlSchema = z
  .object({
    engine: z.enum(["postgres", "sqlserver", "mysql", "oracle"]),
    host: z.string().regex(/^[A-Za-z0-9._-]+$/),
    // The engine's usual port when left out.
    port: z.number().int().min(1).max(65535).optional(),
    // Oracle: the service name.
    database: text,
    // Ids of connect form fields: user a plain one, password a secret one.
    auth: z.object({ user: slug, password: slug }).strict(),
    // "verify": encrypted, the server's certificate checked; "insecure": encrypted, not
    // checked; "off": plain, e.g. a database on the same private network.
    tls: z.enum(["verify", "insecure", "off"]).default("verify"),
    timeoutSeconds: z.number().int().min(1).max(300).default(30),
    // Rows returned per call; more are cut off and the answer says so.
    maxRows: z.number().int().min(1).max(1000).default(200),
    tools: z.array(sqlToolSchema).min(1),
  })
  .strict();
export type SqlApp = z.infer<typeof sqlSchema>;

// Events the app can send (or a poll can read) that start agents (packages/server/src/events,
// docs/triggers.md). A path is `a.b`, or `a[].b` for every item of a list.
const eventPath = z.string().regex(/^[A-Za-z0-9_$@-]+(\[\])?(\.[A-Za-z0-9_$@-]+(\[\])?)*$/);
const eventId = z.string().regex(/^[A-Za-z0-9_.:-]{1,100}$/);
// Where a value is in a delivery: a path in its JSON body, or a request header.
const locator = z.union([
  z.object({ field: eventPath }).strict(),
  z.object({ header: headerName }).strict(),
]);
const eventRequest = z
  .object({
    method: z.enum(["GET", "POST", "PUT"]),
    path: z.string().startsWith("/"),
    body: z.unknown().optional(),
  })
  .strict();
const eventsSchema = z
  .object({
    types: z.array(z.object({ id: eventId, label: text }).strict()).min(1),
    // What an agent is given of each event, and what conditions and limits read: name → path.
    fields: z.record(slug, eventPath),
    // A limit's id → the field holding the event's value for it. An agent with that limit set
    // is started only by events whose value is within it.
    limits: z.record(slug, slug).default({}),
    webhook: z
      .object({
        // How a delivery proves it came from the app, with the subscription's secret: an
        // HMAC-SHA256 of the body (hex, after `prefix`), the secret itself in a header, or
        // nothing at all, only with allowInsecure.
        verify: z.discriminatedUnion("type", [
          z
            .object({
              type: z.literal("hmac-sha256"),
              header: headerName,
              prefix: z.string().max(20).default(""),
            })
            .strict(),
          z.object({ type: z.literal("token"), header: headerName }).strict(),
          z.object({ type: z.literal("none") }).strict(),
        ]),
        allowInsecure: z.boolean().default(false),
        type: locator,
        // Same for every retry of one delivery; without one, the body's hash is used.
        deliveryId: locator.optional(),
        // When the app stamps deliveries: older (or newer) than the tolerance is refused.
        timestamp: z
          .object({
            at: locator,
            unit: z.enum(["s", "ms", "iso"]),
            toleranceSeconds: z
              .number()
              .int()
              .min(60)
              .max(7 * 86400)
              .default(300),
          })
          .strict()
          .optional(),
        // Where the admin pastes the address ({url}) and the secret.
        setup: z.array(step).min(1),
        // Registers the webhook through the app's API with the connection's credentials, filled
        // with {name} {url} {secret} {events} {filter}. `location`: where the answer gives the
        // webhook's own address.
        register: eventRequest.extend({ location: dotPath }).strict().optional(),
      })
      .strict()
      .refine((w) => w.verify.type !== "none" || w.allowInsecure, {
        message: "A webhook verified by nothing needs allowInsecure",
      })
      .optional(),
    // The account the app's connection acts as, so its own changes (an agent's comment) never
    // start agents again: `request` answers its id at `id`; `fields` hold who made the change.
    self: z
      .object({ request: eventRequest, id: dotPath, fields: z.array(slug).min(1) })
      .strict()
      .optional(),
    // For apps without webhooks: `request` read every interval, filled with {minutes} (since
    // the cursor, rounded up, plus one) {cursor} and {filter} (the subscription's filter put
    // into `filter`'s {value}, or nothing); `items` the list in the answer, each kept under
    // `wrap` so the webhook's paths fit; `cursor` the item's change time; `deliveryId` the
    // paths joined into an id; `types` the first whose `same` paths are equal (or that has none).
    poll: z
      .object({
        request: eventRequest,
        filter: z.string().includes("{value}").optional(),
        items: dotPath,
        wrap: slug.optional(),
        cursor: eventPath,
        deliveryId: z.array(eventPath).min(1),
        types: z
          .array(
            z.object({ type: eventId, same: z.tuple([eventPath, eventPath]).optional() }).strict(),
          )
          .min(1),
      })
      .strict()
      .optional(),
  })
  .strict()
  .refine((e) => e.webhook || e.poll, { message: "Events come from a webhook, a poll or both" })
  .refine(
    (e) =>
      Object.values(e.limits).every((f) => f in e.fields) &&
      (e.self?.fields ?? []).every((f) => f in e.fields) &&
      (e.poll?.types ?? []).every((t) => e.types.some((x) => x.id === t.type)),
    { message: "Event limits and self name fields, and a poll's types are declared types" },
  );
export type ConnectorEvents = z.infer<typeof eventsSchema>;

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
    soap: soapSchema.optional(),
    odata: odataSchema.optional(),
    sql: sqlSchema.optional(),
    events: eventsSchema.optional(),
  })
  .strict()
  .refine((c) => [c.http, c.soap, c.odata, c.sql].filter(Boolean).length <= 1, {
    message: "A connector has at most one of http, soap, odata and sql",
  })
  .refine(
    (c) => {
      const app = described(c);
      if (!app) return true;
      const lists = new Set(c.limits.filter((l) => l.kind === "list").map((l) => l.id));
      return (
        app.fields.every(([id, secret]) =>
          c.fields.some((f) => f.id === id && f.secret === secret && !f.options),
        ) &&
        !c.provider &&
        (c.accountKey === "host" || c.fields.some((f) => f.id === c.accountKey && !f.secret)) &&
        app.limits.every((l) => lists.has(l))
      );
    },
    {
      message:
        "An http, soap, odata or sql app's auth names its fields (secrets as secret fields), its accountKey is host or a plain field, a param's or answer's limit names a list limit, and it has no provider",
    },
  )
  .refine(
    (c) =>
      !c.agentAccess || c.limits.some((l) => l.id === c.agentAccess!.limit && l.kind === "text"),
    { message: "agentAccess.limit must be one of the connector's text limits" },
  )
  .refine(
    (c) => {
      if (!c.sql) return true;
      const writes = new Set(c.sql.tools.filter((t) => t.write).map((t) => t.name));
      return (
        [c.verify, ...(c.discover ? [c.discover] : [])].every((p) => "check" in p) &&
        c.capabilities.every((cap) => !cap.required || !cap.tools.some((t) => writes.has(t)))
      );
    },
    { message: "An sql app's probes are checks, and a write tool is in no required capability" },
  )
  .refine(
    (c) =>
      [...Object.keys(c.agentAccess?.probes ?? {}), ...(c.siteAccess?.write ?? [])].every((id) =>
        c.capabilities.some((x) => x.id === id),
      ) &&
      (!Object.keys(c.agentAccess?.probes ?? {}).length || c.agentAccess?.missing !== undefined),
    { message: "agentAccess probes and siteAccess.write name capabilities; probes need missing" },
  )
  .refine(
    (c) =>
      Object.keys(c.events?.limits ?? {}).every((id) =>
        c.limits.some((l) => l.id === id && l.kind !== "map"),
      ),
    { message: "Event limits name the connector's list or text limits" },
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
        c.fields.every((f, i) => !f.requires || isOption(f.requires, c.fields.slice(0, i)))
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

// An app described in its connector.json (http, soap, odata or sql): the connect form fields
// its sign-in names, and the list limits its tools name.
function described(c: {
  http?: HttpApp;
  soap?: SoapApp;
  odata?: ODataApp;
  sql?: SqlApp;
}): { fields: [string, boolean][]; limits: string[] } | null {
  const limitsOf = (...sets: Record<string, { limit?: string }>[]) =>
    sets.flatMap((s) => Object.values(s).flatMap((p) => (p.limit ? [p.limit] : [])));
  if (c.http)
    return {
      fields: authFields(c.http.auth),
      limits: c.http.tools.flatMap((t) => [
        ...limitsOf(t.params),
        ...t.answerLimits.map((a) => a.limit),
      ]),
    };
  if (c.soap) {
    const ws = c.soap.wsSecurity;
    return {
      fields: [
        ...(c.soap.auth ? authFields(c.soap.auth) : []),
        ...(ws
          ? ([
              [ws.username, false],
              [ws.password, true],
            ] as [string, boolean][])
          : []),
      ],
      limits: c.soap.tools.flatMap((t) => [
        ...limitsOf(t.params),
        ...t.answerLimits.map((a) => a.limit),
      ]),
    };
  }
  if (c.odata)
    return {
      fields: authFields(c.odata.auth),
      limits: c.odata.tools.flatMap((t) => [
        ...limitsOf(t.key, t.filters, t.fields, t.params),
        ...t.answerLimits.map((a) => a.limit),
      ]),
    };
  if (c.sql)
    return {
      fields: [
        [c.sql.auth.user, false],
        [c.sql.auth.password, true],
      ],
      limits: c.sql.tools.flatMap((t) => [
        ...limitsOf(t.params),
        ...t.answerLimits.map((a) => a.limit),
      ]),
    };
  return null;
}

export type Connector = z.infer<typeof connectorSchema>;
/** A connector.json as written, before defaults and {brand.*} placeholders are filled. */
export type ConnectorJson = z.input<typeof connectorSchema>;
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

/** The code behind one connector: its vendor client and its tools. */
export type ConnectorModule = {
  id: string;
  // Its connector.json inline, for a module that has no folder of its own.
  connector?: ConnectorJson;
  create(ctx: { microsoft: MicrosoftClient; connector: Connector }): {
    vendor: Vendor;
    toolset: Toolset;
  };
};

/** The repo's connectors/ folder, beside packages/ (and so beside dist/packages/ once built). */
export const CONNECTORS_DIR = new URL("../../../connectors/", import.meta.url);

function parse<T>(raw: unknown, where: string, schema: z.ZodType<T>): T {
  let filled: unknown;
  try {
    filled = fillBrandDeep(raw);
  } catch (error) {
    throw new Error(`${where}: ${(error as Error).message}`);
  }
  const parsed = schema.safeParse(filled);
  if (!parsed.success) throw new Error(`${where}: ${z.prettifyError(parsed.error)}`);
  return parsed.data;
}

/** A definition given inline rather than in a folder, checked as loadConnectors checks one. */
export const parseConnector = (raw: unknown, where: string): Connector =>
  parse(raw, where, connectorSchema);

// Each subfolder of dir holding `file` is one definition, whose id is the folder's name.
// Subfolders may be links: Kubernetes mounts ConfigMap and Secret folders that way.
function loadDir<T>(dir: URL, file: string, schema: z.ZodType<T & { id: string }>): T[] {
  return readdirSync(dir)
    .filter((name) => existsSync(new URL(`${name}/${file}`, dir)))
    .sort()
    .map((name) => {
      const where = `${name}/${file}`;
      const data = parse(JSON.parse(readFileSync(new URL(where, dir), "utf8")), where, schema);
      if (data.id !== name) throw new Error(`${where}: id must match the folder name`);
      return data;
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
