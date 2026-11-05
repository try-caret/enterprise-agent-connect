import { lookup } from "node:dns/promises";
import { isIP } from "node:net";
import { z } from "zod";
import {
  audited,
  binary,
  Denied,
  json,
  listLimit,
  MAX_BINARY_BYTES,
  VendorFailed,
  type Toolset,
} from "../../packages/connector-kit/src/common.js";
import {
  addressFields,
  checkHost,
  fieldRef,
  fillAddress,
  type Connector,
  type ConnectorModule,
  type HttpApp,
  type HttpAuth,
} from "../../packages/connector-kit/src/connector.js";

type HttpTool = HttpApp["tools"][number];
type HttpParam = HttpTool["params"][string];
import { clientCredentials, vendorError } from "../../packages/connector-kit/src/oauth.js";
import {
  ConnectError,
  type Fetch,
  type Limits,
  type Vendor,
  type VendorConnection,
  type VendorRequest,
  type VendorResponse,
} from "../../packages/connector-kit/src/types.js";

// The generic module behind every connector.json with an `http` block: an app described as
// data, e.g. a company's internal REST API. Calls reach only baseUrl's origin under its path.

export const MAX_RESPONSE_BYTES = 1024 * 1024;
export const MAX_BODY_BYTES = 256 * 1024;
export const MAX_TARGET_BYTES = 8 * 1024;
const TIMEOUT_MS = 30_000;

// Cloud metadata services (169.254.169.254, fd00:ec2::254) hand out the server's own credentials,
// and no app is served from a link-local address, so an address that resolves there is refused.
// ponytail: checked at lookup, so a DNS answer that changes before the connect is not caught.
// IPv4-mapped IPv6 comes back from URL as hex: ::ffff:a9fe:a9fe is 169.254.169.254.
const LINK_LOCAL = /^(169\.254\.|::ffff:(169\.254\.|a9fe:)|fe[89ab][0-9a-f]:|fd00:ec2::254$)/i;
async function refuseLinkLocal(hostname: string) {
  const host = hostname.replace(/^\[|\]$/g, "");
  // A name that does not resolve is left to the call itself to fail on.
  const addresses = isIP(host)
    ? [host]
    : await lookup(host, { all: true }).then((all) => all.map((a) => a.address), () => []);
  if (addresses.some((a) => LINK_LOCAL.test(a)))
    throw new ConnectError(`${hostname} is a link-local address, which no app is reached at`);
}

export type AnswerRules = {
  types?: RegExp;
  // What `types` lets through, for the refusal.
  passed?: string;
  maxBytes?: number;
};

// fetch that reaches only `allowed` origins, follows no redirect, gives up after TIMEOUT_MS and
// reads at most `maxBytes` (MAX_RESPONSE_BYTES) of JSON or text (or what `types` matches). Its
// refusals are ConnectError.
export function guardedFetch(
  allowed: string[],
  fetchImpl: Fetch = fetch,
  {
    types = /^(application\/([\w.+-]*\+)?json|text\/)/i,
    passed = "JSON or text",
    maxBytes = MAX_RESPONSE_BYTES,
  }: AnswerRules = {},
): Fetch {
  return (async (input: string | URL | Request, init?: RequestInit) => {
    const url = new URL(input instanceof Request ? input.url : input);
    if (!allowed.includes(url.origin))
      throw new ConnectError(`${url.origin} is not this app's address`);
    await refuseLinkLocal(url.hostname);
    const res = await fetchImpl(url, {
      ...init,
      redirect: "manual",
      signal: AbortSignal.timeout(TIMEOUT_MS),
    });
    if (res.status >= 300 && res.status < 400)
      throw new ConnectError(
        `The app answered with a redirect (${res.status}), which is not followed`,
      );
    const type = res.headers.get("content-type") ?? "";
    if (res.body && !types.test(type) && res.headers.get("content-length") !== "0") {
      await res.body.cancel();
      throw new ConnectError(
        `The app answered ${type || "without a content type"}; only ${passed} is passed on`,
      );
    }
    const chunks: Uint8Array[] = [];
    let size = 0;
    const reader = res.body?.getReader();
    for (let r = await reader?.read(); r && !r.done; r = await reader!.read()) {
      size += r.value.byteLength;
      if (size > maxBytes) {
        await reader!.cancel();
        throw new ConnectError(`The app's answer is larger than ${maxBytes / 1024} KB`);
      }
      chunks.push(r.value);
    }
    // These statuses may not carry a body, even an empty one.
    const body = [204, 205, 304].includes(res.status) ? null : Buffer.concat(chunks);
    return new Response(body, { status: res.status, headers: res.headers });
  }) as Fetch;
}

function parse(raw: string): unknown {
  try {
    return JSON.parse(raw);
  } catch {
    return raw;
  }
}

// The headers signing a call with a key, token or user and password (oauth2 signs its own, and
// query auth goes in the address).
export function authHeaders(auth: HttpAuth, c: VendorConnection): Record<string, string> {
  if (auth.type === "header")
    return { [auth.header.toLowerCase()]: `${auth.prefix ?? ""}${c.secret[auth.value]}` };
  if (auth.type === "bearer") return { authorization: `Bearer ${c.secret[auth.token]}` };
  if (auth.type === "basic")
    return {
      authorization: `Basic ${Buffer.from(`${c.config[auth.username]}:${c.secret[auth.password]}`).toString("base64")}`,
    };
  return {};
}

// The connect form's values: plain fields into config, secret ones into secret; required
// unless `optional`.
export function formValues(connector: Connector, fields: Record<string, string>) {
  const config: Record<string, string> = {};
  const secret: Record<string, string> = {};
  for (const f of connector.fields) {
    const value = fields[f.id]?.trim() ?? "";
    if (!value && f.optional) continue;
    if (!value) throw new ConnectError(`Enter ${f.label}`);
    (f.secret ? secret : config)[f.id] = value;
  }
  return { config, secret };
}

// An app's address and token service filled in for one connection (their {field}s), and the
// guarded fetch reaching only those origins, kept per origin so a token cache survives calls.
export function connectionAddress(
  connector: Connector,
  address: string,
  allowedHosts: string[] | undefined,
  tokenUrl: string | null,
  fetchImpl: Fetch,
  rules?: AnswerRules,
) {
  const label = (id: string) => connector.fields.find((f) => f.id === id)?.label ?? id;
  const fetches = new Map<string, Fetch>();
  return (c: Pick<VendorConnection, "config">) => {
    const url = new URL(fillAddress(address, c.config, label));
    checkHost(url, allowedHosts);
    const token = tokenUrl ? new URL(fillAddress(tokenUrl, c.config, label)) : null;
    if (token && addressFields(tokenUrl!).length) checkHost(token, allowedHosts);
    const key = `${url.origin} ${token?.origin ?? ""}`;
    let safe = fetches.get(key);
    if (!safe) {
      const origins = [url.origin, ...(token ? [token.origin] : [])];
      safe = guardedFetch(origins, fetchImpl, rules);
      fetches.set(key, safe);
    }
    return { url, safe, key, tokenUrl: token?.href ?? null };
  };
}

export function httpVendor(connector: Connector, fetchImpl: Fetch = fetch): Vendor {
  const http = connector.http!;
  const auth = http.auth;
  const addressOf = connectionAddress(
    connector,
    http.baseUrl,
    http.allowedHosts,
    "tokenUrl" in auth ? auth.tokenUrl : null,
    fetchImpl,
  );
  // A binary tool's answer: any type, never parsed, up to the size a tool hands back.
  const bytesAddressOf = connectionAddress(
    connector,
    http.baseUrl,
    http.allowedHosts,
    null,
    fetchImpl,
    { types: /(?:)/, maxBytes: MAX_BINARY_BYTES },
  );
  const clients = new Map<string, ReturnType<typeof clientCredentials>>();
  const oauthFor = (key: string, safe: Fetch) => {
    if (!("tokenUrl" in auth)) return null;
    if (!clients.has(key)) clients.set(key, clientCredentials(connector.name, safe));
    return clients.get(key)!;
  };
  // Lower-case, so a configured Accept replaces the default rather than joining it.
  const fixed = Object.fromEntries(
    Object.entries(http.headers).map(([k, v]) => [k.toLowerCase(), v]),
  );
  const oauthCreds = (c: VendorConnection, tokenUrl: string | null) => {
    if (auth.type === "oauth2")
      return {
        tokenUrl: tokenUrl!,
        clientId: c.config[auth.clientId],
        clientSecret: c.secret[auth.clientSecret],
        scope: auth.scope,
        clientAuth: auth.clientAuth,
        audience: auth.audience,
      };
    if (auth.type === "oauth2_jwt") {
      const audienceField = fieldRef(auth.audience);
      return {
        tokenUrl: tokenUrl!,
        issuer: c.config[auth.issuer],
        subject: c.config[auth.subject],
        audience: audienceField ? c.config[audienceField] : auth.audience,
        scope: auth.scope,
        privateKey: c.secret[auth.privateKey],
      };
    }
    return null!;
  };
  const contentType = (request: VendorRequest) =>
    request.headers?.["content-type"] ?? http.contentType ?? "application/json";

  const fieldHeaders = (c: VendorConnection) =>
    Object.fromEntries(
      Object.entries(http.fieldHeaders).flatMap(([name, id]) => {
        const value = c.secret[id] ?? c.config[id];
        if (!value) return [];
        if (!/^[\x20-\x7E]*$/.test(value))
          throw new ConnectError(`${name} can only hold plain letters, digits and symbols`);
        return [[name.toLowerCase(), value]];
      }),
    );

  const call: Vendor["call"] = async (c, request) => {
    const { url: base, safe, key, tokenUrl } = addressOf(c);
    const prefix = base.pathname.replace(/\/+$/, "");
    const oauth = oauthFor(key, safe);
    // The path is the tool's, already encoded; it must stay under baseUrl's path.
    const url = new URL(`${base.origin}${prefix}${request.path}`);
    if (url.origin !== base.origin || !`${url.pathname}/`.startsWith(`${prefix}/`))
      throw new ConnectError("The call leaves this app's address");
    if (auth.type === "query") url.searchParams.append(auth.param, c.secret[auth.value]);
    const req: VendorRequest = {
      method: request.method,
      path: `${url.pathname}${url.search}`,
      body: request.body,
      ...(request.body === undefined ? {} : { headers: { "content-type": contentType(request) } }),
    };
    const headers = { ...fixed, ...fieldHeaders(c) };
    if (request.responseType === "bytes") {
      const download = bytesAddressOf(c).safe;
      const send = async (signed: Record<string, string>) => {
        const res = await download(`${base.origin}${req.path}`, {
          method: req.method,
          headers: { accept: "*/*", ...headers, ...signed },
        });
        if (!res.ok) {
          const raw = await res.text();
          return { status: res.status, data: raw ? parse(raw) : null };
        }
        const data = new Uint8Array(await res.arrayBuffer());
        return { status: res.status, data, headers: Object.fromEntries(res.headers) };
      };
      if (!oauth) return send(authHeaders(auth, c));
      const creds = oauthCreds(c, tokenUrl);
      const bearer = async () => ({ authorization: `Bearer ${await oauth.token(creds)}` });
      const res = await send(await bearer());
      if (res.status !== 401) return res;
      oauth.forget(creds);
      return send(await bearer());
    }
    if (oauth) return oauth.call(oauthCreds(c, tokenUrl), base.origin, req, headers);
    const res = await safe(`${base.origin}${req.path}`, {
      method: req.method,
      headers: {
        accept: "application/json",
        ...headers,
        ...authHeaders(auth, c),
        ...req.headers,
      },
      body: req.body === undefined ? undefined : JSON.stringify(req.body),
    });
    const raw = await res.text();
    return {
      status: res.status,
      data: raw ? parse(raw) : null,
      headers: Object.fromEntries(res.headers),
    };
  };

  return {
    async connect(fields) {
      const { config, secret } = formValues(connector, fields);
      const { url: base, safe, key, tokenUrl } = addressOf({ config });
      const conn: VendorConnection = { config: { host: base.host, ...config }, secret };
      await oauthFor(key, safe)?.signIn(oauthCreds(conn, tokenUrl));
      return { ...conn, accountLabel: conn.config[connector.accountKey] ?? base.host };
    },
    call,
  };
}

// The values at a dot path; a list on the way gives each of its items' value. Null when any is
// missing.
function valuesAt(data: unknown, keys: string[]): unknown[] | null {
  if (!keys.length) return [data];
  if (Array.isArray(data)) {
    const found = data.map((item) => valuesAt(item, keys));
    return found.some((f) => f === null) ? null : found.flat();
  }
  if (data === null || typeof data !== "object" || !Object.hasOwn(data, keys[0])) return null;
  return valuesAt((data as Record<string, unknown>)[keys[0]], keys.slice(1));
}

// Whether every value of `path` in the answer is among `values`. A missing one fails.
export function answerWithin(data: unknown, path: string, values: string[]): boolean {
  const found = valuesAt(data, path.split("."));
  return (
    found !== null &&
    found.every(
      (v) => (typeof v === "string" || typeof v === "number") && values.includes(String(v)),
    )
  );
}

// The value at a dot path, a list on the way mapped item by item; undefined when missing.
function pick(data: unknown, keys: string[]): unknown {
  if (!keys.length) return data;
  if (Array.isArray(data)) return data.map((item) => pick(item, keys));
  if (data === null || typeof data !== "object" || !Object.hasOwn(data, keys[0])) return undefined;
  return pick((data as Record<string, unknown>)[keys[0]], keys.slice(1));
}

function trim(data: unknown, result: NonNullable<HttpTool["result"]>): unknown {
  // Some apps leave an empty list out of the answer altogether.
  const items = result.items ? (pick(data, result.items.split(".")) ?? []) : data;
  if (!result.fields) return items;
  const keep = (item: unknown) =>
    Object.fromEntries(result.fields!.map((f) => [f, pick(item, f.split("."))]));
  return Array.isArray(items) ? items.map(keep) : keep(items);
}

// The cursor for the next page, or null on the last one.
function nextCursor(tool: HttpTool, res: VendorResponse): string | number | null {
  const paging = tool.paging!;
  const param = tool.params[paging.param];
  let next: unknown;
  if ("body" in paging) next = pick(res.data, paging.body.split("."));
  else if ("header" in paging) next = res.headers?.[paging.header.toLowerCase()];
  else
    for (const [, url, attrs] of (res.headers?.link ?? "").matchAll(/<([^>]*)>([^,]*)/g))
      if (/\brel="?[^"]*\bnext\b/i.test(attrs) && URL.canParse(url, "https://x"))
        next = new URL(url, "https://x").searchParams.get(param.name ?? paging.param);
  if (typeof next === "number") return next;
  if (typeof next !== "string" || next === "") return null;
  return param.type !== "string" && /^-?\d+(\.\d+)?$/.test(next) ? Number(next) : next;
}

// Each param naming a list limit the agent has values in must be sent, and each of its values
// (every item, for an array) must be one of them.
export function paramLimits(
  limits: Limits,
  params: Record<string, { limit?: string }>,
  args: Record<string, unknown>,
) {
  for (const [name, p] of Object.entries(params)) {
    const allowedValues = p.limit ? listLimit(limits, p.limit) : null;
    if (!allowedValues) continue;
    const value = args[name];
    const values = Array.isArray(value) ? value : [value];
    // Left out (or empty), an optional param would mean "all of them".
    if (value === undefined || !values.length)
      throw new Denied(`${name} is required: this agent is limited to some values`);
    for (const v of values)
      if (!allowedValues.includes(String(v)))
        throw new Denied(`${name} ${String(v)} is outside this agent's limits`);
  }
}

// Places a body value at its dot path, building objects on the way (a list for a numeric
// segment). The connector's parse keeps paths from clashing; only own keys are followed, so a
// path never reaches an inherited property such as a prototype's function.
function setPath(body: Record<string, unknown>, keys: string[], value: unknown) {
  let node = body as Record<string, unknown>;
  keys.slice(0, -1).forEach((key, i) => {
    if (!Object.hasOwn(node, key)) node[key] = /^\d+$/.test(keys[i + 1]) ? [] : {};
    node = node[key] as Record<string, unknown>;
  });
  Object.defineProperty(node, keys.at(-1)!, { value, enumerable: true, writable: true, configurable: true });
}

export const zodType = {
  string: z.string(),
  number: z.number(),
  integer: z.number().int(),
  boolean: z.boolean(),
};

function paramType(p: HttpParam) {
  const one = (type: keyof typeof zodType) =>
    p.enum ? z.literal(p.enum) : p.pattern ? z.string().regex(new RegExp(p.pattern)) : zodType[type];
  const type: z.ZodType = (p.type === "array" ? z.array(one(p.items!)) : one(p.type)).describe(
    p.description,
  );
  if (p.default !== undefined) return type.default(p.default);
  return p.required ? type : type.optional();
}

export function httpTools(http: HttpApp): Toolset {
  return {
    names: http.tools.map((t) => t.name),
    register(server, ctx, allowed) {
      for (const tool of http.tools) {
        if (!allowed.has(tool.name)) continue;
        const shape = Object.fromEntries(
          Object.entries(tool.params).flatMap(([name, p]) => (p.field ? [] : [[name, paramType(p)]])),
        );
        const destructive = tool.destructive || tool.method === "DELETE";
        server.registerTool(
          tool.name,
          {
            description: tool.description,
            inputSchema: z.object(shape).strict(),
            annotations: {
              readOnlyHint: tool.method === "GET" || tool.readOnly,
              destructiveHint: destructive,
              openWorldHint: true,
            },
          },
          (args: Record<string, unknown>) =>
            audited(ctx, tool.name, JSON.stringify(args).slice(0, 300), async (mark) => {
              paramLimits(ctx.limits, tool.params, args);
              const path = tool.path.replace(/\{([a-z][a-z0-9_]*)\}/g, (_, name: string) => {
                const value = String(args[name]);
                // URLs resolve . and .. segments even when percent-encoded.
                if (value === "." || value === ".." || value === "")
                  throw new Denied(`${name} is not a valid value`);
                return encodeURIComponent(value);
              });
              const query = new URLSearchParams();
              const body = structuredClone(tool.body ?? {});
              for (const [name, p] of Object.entries(tool.params)) {
                const value = p.field ? ctx.config[p.field] : args[name];
                const wire = p.name ?? name;
                if (value === undefined) continue;
                if (p.in === "body") setPath(body, tool.nestedBody ? wire.split(".") : [wire], value);
                else if (p.in !== "query") continue;
                else if (!Array.isArray(value)) query.set(wire, String(value));
                else if (p.join) query.set(wire, value.join(p.join));
                else for (const v of value) query.append(wire, String(v));
              }
              const target = query.size ? `${path}?${query}` : path;
              if (Buffer.byteLength(target) > MAX_TARGET_BYTES)
                throw new Denied(`The address is longer than ${MAX_TARGET_BYTES / 1024} KB`);
              const hasBody =
                tool.body !== undefined || Object.values(tool.params).some((p) => p.in === "body");
              if (hasBody && Buffer.byteLength(JSON.stringify(body)) > MAX_BODY_BYTES)
                throw new Denied(`The request is larger than ${MAX_BODY_BYTES / 1024} KB`);
              let res;
              try {
                res = await ctx.call({
                  method: tool.method,
                  path: target,
                  ...(tool.binary ? { responseType: "bytes" as const } : {}),
                  ...(hasBody ? { body } : {}),
                  ...(hasBody && tool.contentType
                    ? { headers: { "content-type": tool.contentType } }
                    : {}),
                });
              } catch (error) {
                if (error instanceof ConnectError) throw new VendorFailed(502, error.message);
                throw error;
              }
              mark(res.status);
              if (res.status < 200 || res.status >= 300) {
                const detail = vendorError(res.data);
                throw new VendorFailed(
                  res.status,
                  `The app answered ${res.status}${detail ? `: ${detail}` : ""}`,
                );
              }
              for (const a of tool.answerLimits) {
                const values = listLimit(ctx.limits, a.limit);
                if (values && !answerWithin(res.data, a.field, values))
                  throw new Denied(`The answer's ${a.field} is outside this agent's limits`);
              }
              if (tool.binary)
                return binary(
                  res.data as Uint8Array,
                  res.headers?.["content-type"] || "application/octet-stream",
                  target,
                );
              if (!tool.result && !tool.paging)
                return typeof res.data === "string"
                  ? { content: [{ type: "text", text: res.data }] }
                  : json(res.data);
              if (typeof res.data === "string")
                throw new VendorFailed(res.status, "The app answered with text, not JSON");
              const data = tool.result ? trim(res.data, tool.result) : res.data;
              return json(tool.paging ? { items: data, next: nextCursor(tool, res) } : data);
            }),
        );
      }
    },
  };
}

/** Binds to any connector with an `http` block; createAgentConnect uses it for those. */
export default {
  id: "http",
  create: ({ connector }) => ({
    vendor: httpVendor(connector),
    toolset: httpTools(connector.http!),
  }),
} satisfies ConnectorModule;
