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
  HttpApp,
} from "../../packages/connector-kit/src/connector.js";

type HttpTool = HttpApp["tools"][number];
type HttpParam = HttpTool["params"][string];
import { clientCredentials, vendorError } from "../../packages/connector-kit/src/oauth.js";
import {
  ConnectError,
  type Fetch,
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

// fetch that reaches only `allowed` origins, follows no redirect, gives up after TIMEOUT_MS and
// reads at most MAX_RESPONSE_BYTES of JSON or text. Its refusals are ConnectError.
export function guardedFetch(allowed: string[], fetchImpl: Fetch = fetch): Fetch {
  return (async (input: string | URL | Request, init?: RequestInit) => {
    const url = new URL(input instanceof Request ? input.url : input);
    if (!allowed.includes(url.origin))
      throw new ConnectError(`${url.origin} is not this app's address`);
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
    if (
      res.body &&
      !/^(application\/([\w.+-]*\+)?json|text\/)/i.test(type) &&
      res.headers.get("content-length") !== "0"
    ) {
      await res.body.cancel();
      throw new ConnectError(
        `The app answered ${type || "without a content type"}; only JSON or text is passed on`,
      );
    }
    const chunks: Uint8Array[] = [];
    let size = 0;
    const reader = res.body?.getReader();
    for (let r = await reader?.read(); r && !r.done; r = await reader!.read()) {
      size += r.value.byteLength;
      if (size > MAX_RESPONSE_BYTES) {
        await reader!.cancel();
        throw new ConnectError(`The app's answer is larger than ${MAX_RESPONSE_BYTES / 1024} KB`);
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

export function httpVendor(connector: Connector, fetchImpl: Fetch = fetch): Vendor {
  const http = connector.http!;
  const base = new URL(http.baseUrl);
  const prefix = base.pathname.replace(/\/+$/, "");
  const auth = http.auth;
  const safe = guardedFetch(
    [base.origin, ...(auth.type === "oauth2" ? [new URL(auth.tokenUrl).origin] : [])],
    fetchImpl,
  );
  const oauth = auth.type === "oauth2" ? clientCredentials(connector.name, safe) : null;
  // Lower-case, so a configured Accept replaces the default rather than joining it.
  const fixed = Object.fromEntries(
    Object.entries(http.headers).map(([k, v]) => [k.toLowerCase(), v]),
  );
  const oauthCreds = (c: VendorConnection) =>
    auth.type === "oauth2"
      ? {
          tokenUrl: auth.tokenUrl,
          clientId: c.config[auth.clientId],
          clientSecret: c.secret[auth.clientSecret],
        }
      : null!;

  function headersFor(c: VendorConnection): Record<string, string> {
    if (auth.type === "header")
      return { [auth.header.toLowerCase()]: `${auth.prefix ?? ""}${c.secret[auth.value]}` };
    if (auth.type === "bearer") return { authorization: `Bearer ${c.secret[auth.token]}` };
    if (auth.type === "basic")
      return {
        authorization: `Basic ${Buffer.from(`${c.config[auth.username]}:${c.secret[auth.password]}`).toString("base64")}`,
      };
    return {};
  }

  const call: Vendor["call"] = async (c, request) => {
    // The path is the tool's, already encoded; it must stay under baseUrl's path.
    const url = new URL(`${base.origin}${prefix}${request.path}`);
    if (url.origin !== base.origin || !`${url.pathname}/`.startsWith(`${prefix}/`))
      throw new ConnectError("The call leaves this app's address");
    if (auth.type === "query") url.searchParams.append(auth.param, c.secret[auth.value]);
    const req: VendorRequest = {
      method: request.method,
      path: `${url.pathname}${url.search}`,
      body: request.body,
    };
    if (oauth) return oauth.call(oauthCreds(c), base.origin, req, fixed);
    const res = await safe(`${base.origin}${req.path}`, {
      method: req.method,
      headers: {
        accept: "application/json",
        ...fixed,
        ...headersFor(c),
        ...(req.body === undefined ? {} : { "content-type": "application/json" }),
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
      const config: Record<string, string> = { host: base.host };
      const secret: Record<string, string> = {};
      for (const f of connector.fields) {
        const value = fields[f.id]?.trim() ?? "";
        if (!value) throw new ConnectError(`Enter ${f.label}`);
        (f.secret ? secret : config)[f.id] = value;
      }
      const conn = { config, secret };
      if (oauth) await oauth.signIn(oauthCreds(conn));
      return { ...conn, accountLabel: config[connector.accountKey] ?? base.host };
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
function answerWithin(data: unknown, path: string, values: string[]): boolean {
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

const zodType = {
  string: z.string(),
  number: z.number(),
  integer: z.number().int(),
  boolean: z.boolean(),
};

function paramType(p: HttpParam) {
  const one = (type: keyof typeof zodType) => (p.enum ? z.literal(p.enum) : zodType[type]);
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
          Object.entries(tool.params).map(([name, p]) => [name, paramType(p)]),
        );
        const destructive = tool.destructive || tool.method === "DELETE";
        server.registerTool(
          tool.name,
          {
            description: tool.description,
            inputSchema: z.object(shape).strict(),
            annotations: {
              readOnlyHint: tool.method === "GET",
              destructiveHint: destructive,
              openWorldHint: true,
            },
          },
          (args: Record<string, unknown>) =>
            audited(ctx, tool.name, JSON.stringify(args).slice(0, 300), async (mark) => {
              for (const [name, p] of Object.entries(tool.params)) {
                const allowedValues = p.limit ? listLimit(ctx.limits, p.limit) : null;
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
              const path = tool.path.replace(/\{([a-z][a-z0-9_]*)\}/g, (_, name: string) => {
                const value = String(args[name]);
                // URLs resolve . and .. segments even when percent-encoded.
                if (value === "." || value === ".." || value === "")
                  throw new Denied(`${name} is not a valid value`);
                return encodeURIComponent(value);
              });
              const query = new URLSearchParams();
              const body: Record<string, unknown> = {};
              for (const [name, p] of Object.entries(tool.params)) {
                const value = args[name];
                const wire = p.name ?? name;
                if (value === undefined) continue;
                if (p.in === "body") body[wire] = value;
                else if (p.in !== "query") continue;
                else if (!Array.isArray(value)) query.set(wire, String(value));
                else if (p.join) query.set(wire, value.join(p.join));
                else for (const v of value) query.append(wire, String(v));
              }
              const target = query.size ? `${path}?${query}` : path;
              if (Buffer.byteLength(target) > MAX_TARGET_BYTES)
                throw new Denied(`The address is longer than ${MAX_TARGET_BYTES / 1024} KB`);
              const hasBody = Object.values(tool.params).some((p) => p.in === "body");
              if (hasBody && Buffer.byteLength(JSON.stringify(body)) > MAX_BODY_BYTES)
                throw new Denied(`The request is larger than ${MAX_BODY_BYTES / 1024} KB`);
              let res;
              try {
                res = await ctx.call({
                  method: tool.method,
                  path: target,
                  ...(hasBody ? { body } : {}),
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
