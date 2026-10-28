import { z } from "zod";
import {
  audited,
  Denied,
  json,
  listLimit,
  VendorFailed,
  type Toolset,
} from "../../packages/connector-kit/src/common.js";
import type {
  Connector,
  ConnectorModule,
  ODataApp,
  ODataTool,
} from "../../packages/connector-kit/src/connector.js";
import { clientCredentials, vendorError } from "../../packages/connector-kit/src/oauth.js";
import {
  ConnectError,
  type Fetch,
  type Vendor,
  type VendorConnection,
  type VendorRequest,
} from "../../packages/connector-kit/src/types.js";
import type { Limits } from "../../packages/server/src/store.js";
import {
  answerWithin,
  authHeaders,
  formValues,
  guardedFetch,
  MAX_BODY_BYTES,
  MAX_TARGET_BYTES,
  paramLimits,
} from "../http/index.js";

// The generic module behind every connector.json with an `odata` block: an OData v2 or v4
// service (SAP Gateway and S/4HANA, Dynamics, SuccessFactors) described as data. Agents send
// typed values; this module writes every literal, $filter and key, so no OData syntax of
// theirs reaches the service.

type Version = ODataApp["version"];
type Value = ODataTool["key"][string];
type Params = Record<string, Value>;

const DATETIME = z
  .string()
  .regex(/^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}(:\d{2}(\.\d{1,7})?)?(Z|[+-]\d{2}:\d{2})?)?$/);

export const TYPES = {
  string: z.string(),
  integer: z.number().int(),
  number: z.number(),
  decimal: z.string().regex(/^-?\d{1,29}(\.\d{1,29})?$/),
  boolean: z.boolean(),
  guid: z
    .string()
    .regex(/^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/),
  // A date, or a date and time; without an offset it is UTC.
  datetime: DATETIME,
  // The same; v2 writes it as a datetimeoffset literal (SuccessFactors' change dates).
  datetimeoffset: DATETIME,
};

function instant(value: string): Date {
  const d = new Date(/T/.test(value) && !/(Z|[+-]\d{2}:\d{2})$/.test(value) ? `${value}Z` : value);
  if (Number.isNaN(d.getTime())) throw new Denied(`${value} is not a date`);
  return d;
}

// The URL literal of a value: strings quoted with ' doubled, v2's typed prefixes and suffixes.
export function literal(type: Value["type"], value: unknown, v: Version): string {
  // The input schema checks this too; a literal is never written from an unchecked value.
  if (!TYPES[type].safeParse(value).success) throw new Denied(`${String(value)} is not a ${type}`);
  const s = String(value);
  switch (type) {
    case "string":
      return `'${s.replaceAll("'", "''")}'`;
    case "decimal":
      return v === "2" ? `${s}M` : s;
    case "number":
      return v === "2" ? `${s}d` : s;
    case "guid":
      return v === "2" ? `guid'${s}'` : s;
    case "datetime": {
      const iso = instant(s).toISOString();
      return v === "2" ? `datetime'${iso.slice(0, 19)}'` : iso;
    }
    case "datetimeoffset": {
      const iso = instant(s).toISOString();
      return v === "2" ? `datetimeoffset'${iso}'` : iso;
    }
    default:
      return s;
  }
}

// The JSON value of a field or action parameter.
function jsonValue(type: Value["type"], value: unknown, v: Version): unknown {
  if (type === "decimal") return v === "2" ? value : Number(value);
  if (type === "datetime" || type === "datetimeoffset") {
    const d = instant(String(value));
    if (v === "4") return d.toISOString();
    return `/Date(${d.getTime()}${type === "datetimeoffset" ? "+0000" : ""})/`;
  }
  return value;
}

// A limit's value (always text) as the param's type, refused if it is not one.
function limitValue(type: Value["type"], raw: string): unknown {
  const value =
    type === "integer" || type === "number"
      ? Number(raw)
      : type === "boolean"
        ? raw === "true"
        : raw;
  if (!TYPES[type].safeParse(value).success)
    throw new Error(`Limit value ${raw} is not a valid ${type}`);
  return value;
}

const prop = (name: string, p: { property?: string }) => p.property ?? name;
const isObject = (v: unknown): v is Record<string, unknown> =>
  Boolean(v) && typeof v === "object" && !Array.isArray(v);

// Values the agent sent for limited params must be the agent's (left out is allowed here).
function sentWithin(limits: Limits, params: Params, args: Record<string, unknown>) {
  for (const [name, p] of Object.entries(params)) {
    const allowed = p.limit ? listLimit(limits, p.limit) : null;
    if (allowed && args[name] !== undefined && !allowed.includes(String(args[name])))
      throw new Denied(`${name} ${String(args[name])} is outside this agent's limits`);
  }
}

// The answer without OData's bookkeeping (__metadata, @odata.*, deferred links), and only the
// selected properties of each entity.
export function trim(value: unknown, select: string[] = []): unknown {
  if (Array.isArray(value)) return value.map((item) => trim(item, select));
  if (!isObject(value)) return value;
  return Object.fromEntries(
    Object.entries(value).flatMap(([k, v]) =>
      k.startsWith("@") ||
      k.startsWith("__") ||
      (isObject(v) && "__deferred" in v) ||
      (select.length && !select.includes(k))
        ? []
        : [[k, trim(v)]],
    ),
  );
}

// v2's {d: {results, __next}} and v4's {value, @odata.nextLink}: the entity, list or value.
export function unwrap(data: unknown, v: Version): { value: unknown; next: string | null } {
  const d = v === "2" && isObject(data) && "d" in data ? data.d : data;
  const next = (n: unknown) => (typeof n === "string" ? n : null);
  if (isObject(d) && Array.isArray(d.results)) return { value: d.results, next: next(d.__next) };
  if (
    v === "4" &&
    isObject(d) &&
    "value" in d &&
    Object.keys(d).every((k) => k === "value" || k.startsWith("@"))
  )
    return { value: d.value, next: next(d["@odata.nextLink"]) };
  return { value: d, next: null };
}

// A query option of a next link, read raw: URLSearchParams would turn a + into a space.
function option(link: string, name: string): string | null {
  const query = link.split("?")[1] ?? "";
  for (const part of query.split("&")) {
    const [key, value = ""] = part.split("=");
    if (decodeURIComponent(key) === name) return decodeURIComponent(value);
  }
  return null;
}

export function odataVendor(connector: Connector, fetchImpl: Fetch = fetch): Vendor {
  const od = connector.odata!;
  const root = new URL(od.serviceRoot);
  const prefix = root.pathname.replace(/\/+$/, "");
  const auth = od.auth;
  const safe = guardedFetch(
    [root.origin, ...(auth.type === "oauth2" ? [new URL(auth.tokenUrl).origin] : [])],
    fetchImpl,
  );
  const oauth = auth.type === "oauth2" ? clientCredentials(connector.name, safe) : null;
  const creds = (c: VendorConnection) =>
    auth.type === "oauth2"
      ? {
          tokenUrl: auth.tokenUrl,
          clientId: c.config[auth.clientId],
          clientSecret: c.secret[auth.clientSecret],
          scope: auth.scope,
        }
      : null!;
  const versionHeaders: Record<string, string> =
    od.version === "4"
      ? { "odata-version": "4.0", "odata-maxversion": "4.0" }
      : { dataserviceversion: "2.0", maxdataserviceversion: "2.0" };

  async function send(c: VendorConnection, request: VendorRequest) {
    // The path is the tool's, already encoded; it must stay under the service root.
    const url = new URL(`${root.origin}${prefix}${request.path}`);
    if (url.origin !== root.origin || !`${url.pathname}/`.startsWith(`${prefix}/`))
      throw new ConnectError("The call leaves this app's address");
    // Appended as is: URLSearchParams would rewrite the encoded $filter's spaces as +.
    if (auth.type === "query")
      url.search += `${url.search ? "&" : "?"}${encodeURIComponent(auth.param)}=${encodeURIComponent(c.secret[auth.value])}`;
    const signed = async () =>
      safe(url, {
        method: request.method,
        headers: {
          accept: "application/json",
          ...versionHeaders,
          ...od.headers,
          ...(oauth
            ? { authorization: `Bearer ${await oauth.token(creds(c))}` }
            : authHeaders(auth, c)),
          ...request.headers,
          ...(request.body === undefined ? {} : { "content-type": "application/json" }),
        },
        body: request.body === undefined ? undefined : JSON.stringify(request.body),
      });
    let res = await signed();
    if (res.status === 401 && oauth) {
      oauth.forget(creds(c));
      res = await signed();
    }
    return res;
  }

  // SAP Gateway refuses a write without a token fetched in the same session (its cookies).
  // ponytail: a fresh token per write; cache it per connection if writes get frequent.
  async function csrf(c: VendorConnection): Promise<Record<string, string>> {
    const res = await send(c, { method: "GET", path: "/", headers: { "x-csrf-token": "Fetch" } });
    const token = res.headers.get("x-csrf-token");
    if (!res.ok || !token || /^required$/i.test(token))
      throw new ConnectError(`${connector.name} gave no CSRF token (it answered ${res.status})`);
    const cookie = res.headers
      .getSetCookie()
      .map((s) => s.split(";")[0])
      .join("; ");
    return { "x-csrf-token": token, ...(cookie ? { cookie } : {}) };
  }

  return {
    async connect(fields) {
      const { config, secret } = formValues(connector, fields);
      const conn: VendorConnection = { config: { host: root.host, ...config }, secret };
      if (oauth) await oauth.signIn(creds(conn));
      return { ...conn, accountLabel: conn.config[connector.accountKey] ?? root.host };
    },
    async call(c, request) {
      const extra = request.method !== "GET" && od.csrf ? await csrf(c) : {};
      const res = await send(c, { ...request, headers: { ...request.headers, ...extra } });
      const raw = await res.text();
      let data: unknown = null;
      if (raw)
        try {
          data = JSON.parse(raw);
        } catch {
          data = raw;
        }
      return { status: res.status, data };
    },
  };
}

// v2: {error: {code, message: {value}}}; v4: {error: {code, message}}.
function odataError(data: unknown): string {
  const e = isObject(data) && isObject(data.error) ? data.error : null;
  if (!e) return vendorError(data);
  const message = isObject(e.message) ? e.message.value : e.message;
  return [e.code, message]
    .filter((x) => typeof x === "string" && x)
    .join(": ")
    .slice(0, 500);
}

function filterClause(property: string, op: string, lit: string, v: Version): string {
  if (op === "contains")
    return v === "2" ? `substringof(${lit},${property})` : `contains(${property},${lit})`;
  if (op === "startswith")
    return v === "2" ? `startswith(${property},${lit}) eq true` : `startswith(${property},${lit})`;
  return `${property} ${op} ${lit}`;
}

const query = (options: [string, string | undefined][]) => {
  const parts = options.flatMap(([k, val]) =>
    val === undefined ? [] : [`${k}=${encodeURIComponent(val)}`],
  );
  return parts.length ? `?${parts.join("&")}` : "";
};

export function odataTools(od: ODataApp): Toolset {
  const v = od.version;

  function keyPath(tool: ODataTool, args: Record<string, unknown>) {
    const parts = Object.entries(tool.key).map(
      ([n, p]) => `${prop(n, p)}=${encodeURIComponent(literal(p.type, args[n], v))}`,
    );
    return `/${tool.entitySet}(${parts.join(",")})`;
  }

  function body(params: Params, args: Record<string, unknown>) {
    const out: Record<string, unknown> = {};
    for (const [n, p] of Object.entries(params))
      if (args[n] !== undefined) out[prop(n, p)] = jsonValue(p.type, args[n], v);
    return out;
  }

  // What each kind sends: method, path and query, body and headers; `paging` for lists.
  function request(
    tool: ODataTool,
    args: Record<string, unknown>,
    limits: Limits,
  ): { req: VendorRequest; top?: number; offset?: number } {
    const select = tool.select.length ? tool.select.join(",") : undefined;
    const representation: Record<string, string> =
      v === "4" ? { prefer: "return=representation" } : {};
    switch (tool.kind) {
      case "list": {
        const clauses: string[] = [];
        for (const [n, f] of Object.entries(tool.filters)) {
          const property = prop(n, f);
          const allowed = f.limit ? listLimit(limits, f.limit) : null;
          if (args[n] === undefined) {
            // Left out, a limited filter means all of the agent's values.
            if (allowed)
              clauses.push(
                `(${allowed.map((a) => `${property} eq ${literal(f.type, limitValue(f.type, a), v)}`).join(" or ")})`,
              );
            continue;
          }
          if (allowed && !allowed.includes(String(args[n])))
            throw new Denied(`${n} ${String(args[n])} is outside this agent's limits`);
          clauses.push(filterClause(property, f.op, literal(f.type, args[n], v), v));
        }
        const top = Math.min(Number(args.max_results ?? tool.top), tool.top);
        const cursor = args.cursor as string | undefined;
        if (cursor !== undefined && !/^(s\d{1,9}|t.{1,4000})$/s.test(cursor))
          throw new Denied("cursor is not one this tool gave");
        const server = od.paging === "server";
        const offset = cursor?.startsWith("s") ? Number(cursor.slice(1)) : undefined;
        return {
          req: {
            method: "GET",
            path: `/${tool.entitySet}${query([
              ["$filter", clauses.length ? clauses.join(" and ") : undefined],
              ["$select", select],
              ["$orderby", tool.orderBy],
              ["$top", server ? undefined : String(top)],
              ["$skip", offset === undefined ? undefined : String(offset)],
              ["$skiptoken", cursor?.startsWith("t") ? cursor.slice(1) : undefined],
            ])}`,
            headers: server ? { prefer: `odata.maxpagesize=${top}` } : {},
          },
          top,
          offset: cursor === undefined ? 0 : offset,
        };
      }
      case "get":
        return {
          req: { method: "GET", path: `${keyPath(tool, args)}${query([["$select", select]])}` },
        };
      case "create":
        return {
          req: {
            method: "POST",
            path: `/${tool.entitySet}`,
            body: body(tool.fields, args),
            headers: representation,
          },
        };
      case "update": {
        const changes = body(tool.fields, args);
        if (!Object.keys(changes).length) throw new Denied("Give at least one field to change");
        return {
          req:
            v === "4"
              ? {
                  method: "PATCH",
                  path: keyPath(tool, args),
                  body: changes,
                  headers: { "if-match": "*", ...representation },
                }
              : {
                  method: "POST",
                  path: keyPath(tool, args),
                  body: changes,
                  headers: { "if-match": "*", "x-http-method": "MERGE" },
                },
        };
      }
      default: {
        const fn = tool.function!;
        const sent = Object.entries(tool.params).filter(([n]) => args[n] !== undefined);
        const lits = sent.map(([n, p]) => [prop(n, p), literal(p.type, args[n], v)] as const);
        if (tool.kind === "action" && v === "4")
          return { req: { method: "POST", path: `/${fn}`, body: body(tool.params, args) } };
        const path =
          v === "4"
            ? `/${fn}(${lits.map(([k, l]) => `${k}=${encodeURIComponent(l)}`).join(",")})`
            : `/${fn}${query(lits.map(([k, l]) => [k, l]))}`;
        return { req: { method: tool.kind === "action" ? "POST" : "GET", path } };
      }
    }
  }

  return {
    names: od.tools.map((t) => t.name),
    register(server, ctx, allowed) {
      for (const tool of od.tools) {
        if (!allowed.has(tool.name)) continue;
        const required = (part: "key" | "filters" | "fields" | "params", p: Value) =>
          part === "key" ||
          (part === "fields"
            ? tool.kind === "create" && (p.required ?? true)
            : (p.required ?? part === "params"));
        const shape: Record<string, z.ZodType> = {};
        for (const part of ["key", "filters", "fields", "params"] as const)
          for (const [n, p] of Object.entries(tool[part])) {
            const type = TYPES[p.type].describe(p.description);
            shape[n] = required(part, p) ? type : type.optional();
          }
        if (tool.kind === "list") {
          shape.max_results = z
            .number()
            .int()
            .min(1)
            .max(tool.top)
            .optional()
            .describe(`How many to return, at most ${tool.top} (the default)`);
          shape.cursor = z
            .string()
            .optional()
            .describe("The `next` of the previous page, to read the page after it");
        }
        const reads = tool.kind === "list" || tool.kind === "get" || tool.kind === "function";
        server.registerTool(
          tool.name,
          {
            description: tool.description,
            inputSchema: z.object(shape).strict(),
            annotations: {
              readOnlyHint: reads,
              destructiveHint: tool.destructive,
              openWorldHint: true,
            },
          },
          (args: Record<string, unknown>) =>
            audited(ctx, tool.name, JSON.stringify(args).slice(0, 300), async (mark, retarget) => {
              paramLimits(ctx.limits, tool.key, args);
              paramLimits(ctx.limits, tool.params, args);
              if (tool.kind === "create") paramLimits(ctx.limits, tool.fields, args);
              else sentWithin(ctx.limits, tool.fields, args);
              const { req, top, offset } = request(tool, args, ctx.limits);
              if (Buffer.byteLength(req.path) > MAX_TARGET_BYTES)
                throw new Denied(`The address is longer than ${MAX_TARGET_BYTES / 1024} KB`);
              if (
                req.body !== undefined &&
                Buffer.byteLength(JSON.stringify(req.body)) > MAX_BODY_BYTES
              )
                throw new Denied(`The request is larger than ${MAX_BODY_BYTES / 1024} KB`);
              retarget(`${req.method} ${req.path}`.slice(0, 300));
              const send = async (r: VendorRequest) => {
                let res;
                try {
                  res = await ctx.call(r);
                } catch (error) {
                  if (error instanceof ConnectError) throw new VendorFailed(502, error.message);
                  throw error;
                }
                mark(res.status);
                if (res.status < 200 || res.status >= 300) {
                  const detail = odataError(res.data);
                  throw new VendorFailed(
                    res.status,
                    `The service answered ${res.status}${detail ? `: ${detail}` : ""}`,
                  );
                }
                return res;
              };
              const answerLimited = tool.answerLimits.flatMap((a) => {
                const values = listLimit(ctx.limits, a.limit);
                return values ? [{ field: a.field, values }] : [];
              });
              const outside = (data: unknown) =>
                answerLimited.find((a) => !answerWithin(data, a.field, a.values));
              // An update reads the entity first, so one outside the agent's values is never changed.
              if (tool.kind === "update" && answerLimited.length) {
                const fields = [...new Set(answerLimited.map((a) => a.field.split(".")[0]))];
                const before = await send({
                  method: "GET",
                  path: `${keyPath(tool, args)}${query([["$select", fields.join(",")]])}`,
                });
                const miss = outside(unwrap(before.data, v).value);
                if (miss)
                  throw new Denied(`The entity's ${miss.field} is outside this agent's limits`);
              }
              const res = await send(req);
              if (res.data === null || res.data === "")
                return json(tool.kind === "create" ? { created: true } : { done: true });
              const { value, next } = unwrap(res.data, v);
              if (reads) {
                const miss = outside(value);
                if (miss)
                  throw new Denied(`The answer's ${miss.field} is outside this agent's limits`);
              }
              if (tool.kind === "list") {
                for (const [n, f] of Object.entries(tool.filters)) {
                  const values = f.limit ? listLimit(ctx.limits, f.limit) : null;
                  if (values && !answerWithin(value, prop(n, f), values))
                    throw new Denied(`The answer's ${prop(n, f)} is outside this agent's limits`);
                }
                const items = trim(value, tool.select) as unknown[];
                const token = next && option(next, "$skiptoken");
                const skip = next && option(next, "$skip");
                const cursor = token
                  ? `t${token}`
                  : skip && /^\d+$/.test(skip)
                    ? `s${skip}`
                    : !next && od.paging === "skip" && offset !== undefined && items.length >= top!
                      ? `s${offset + top!}`
                      : null;
                return json({ items, next: cursor });
              }
              return json(trim(value, tool.select));
            }),
        );
      }
    },
  };
}

/** Binds to any connector with an `odata` block; createAgentConnect uses it for those. */
export default {
  id: "odata",
  create: ({ connector }) => ({
    vendor: odataVendor(connector),
    toolset: odataTools(connector.odata!),
  }),
} satisfies ConnectorModule;
