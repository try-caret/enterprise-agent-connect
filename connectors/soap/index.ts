import { createHash, randomBytes } from "node:crypto";
import { z } from "zod";
import {
  audited,
  Denied,
  json,
  listLimit,
  VendorFailed,
  type Toolset,
} from "../../packages/connector-kit/src/common.js";
import {
  pickDisplay,
  type Connector,
  type ConnectorModule,
  type SoapApp,
  type SoapTool,
} from "../../packages/connector-kit/src/connector.js";
import { vendorError } from "../../packages/connector-kit/src/oauth.js";
import {
  ConnectError,
  type Fetch,
  type Vendor,
  type VendorConnection,
} from "../../packages/connector-kit/src/types.js";
import {
  answerWithin,
  authHeaders,
  formValues,
  guardedFetch,
  MAX_BODY_BYTES,
  paramLimits,
  zodType,
} from "../http/index.js";
import { escapeXml, localName, parseXml, toJson, XmlError, type XmlNode } from "./xml.js";

// The generic module behind every connector.json with a `soap` block: a SOAP 1.1 or 1.2
// service described as data. Calls reach only the endpoint's origin, always POSTing to it.

const ENVELOPE = {
  "1.1": "http://schemas.xmlsoap.org/soap/envelope/",
  "1.2": "http://www.w3.org/2003/05/soap-envelope",
};
const WSSE = "http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd";
const WSU = "http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd";
const TOKEN_PROFILE =
  "http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-username-token-profile-1.0";
const XML_TYPES = /^(text\/|application\/([\w.+-]*\+)?(xml|json))/i;
// Characters XML 1.0 cannot carry, even escaped.
const NOT_XML = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/;

// The Body content of one call: the template with each value escaped, and an optional
// param's element dropped when it is left out.
export function renderBody(tool: SoapTool, args: Record<string, unknown>): string {
  let body = tool.body;
  for (const [name, p] of Object.entries(tool.params)) {
    if (args[name] !== undefined) continue;
    if (p.required) throw new Denied(`${name} is required`);
    body = body.replace(new RegExp(`<([\\w.:-]+)(\\s[^>]*)?>\\{${name}\\}</\\1>`), "");
  }
  // One pass, so a value holding "{other}" is never read as a placeholder.
  return body.replace(/\{([a-z][a-z0-9_]*)\}/g, (_, name: string) => {
    const value = String(args[name]);
    if (NOT_XML.test(value)) throw new Denied(`${name} has characters XML cannot carry`);
    return escapeXml(value);
  });
}

function usernameToken(username: string, password: string, digest: boolean): string {
  const nonce = randomBytes(16);
  const created = new Date().toISOString();
  const value = digest
    ? createHash("sha1")
        .update(Buffer.concat([nonce, Buffer.from(created + password)]))
        .digest("base64")
    : password;
  return (
    `<wsse:Security xmlns:wsse="${WSSE}" xmlns:wsu="${WSU}"><wsse:UsernameToken>` +
    `<wsse:Username>${escapeXml(username)}</wsse:Username>` +
    `<wsse:Password Type="${TOKEN_PROFILE}#${digest ? "PasswordDigest" : "PasswordText"}">${escapeXml(value)}</wsse:Password>` +
    `<wsse:Nonce EncodingType="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-soap-message-security-1.0#Base64Binary">${nonce.toString("base64")}</wsse:Nonce>` +
    `<wsu:Created>${created}</wsu:Created></wsse:UsernameToken></wsse:Security>`
  );
}

const child = (node: XmlNode | undefined, name: string) =>
  node?.children.find((c) => localName(c.name) === name);

// A SOAP Fault's code and text, from either version's shape.
function faultText(fault: XmlNode): string {
  const f = toJson(fault) as Record<string, unknown>;
  const first = (v: unknown): unknown => (Array.isArray(v) ? v[0] : v);
  const reason = f.Reason as Record<string, unknown> | undefined;
  const code = f.Code as Record<string, unknown> | undefined;
  const message = String(first(f.faultstring ?? reason?.Text) ?? "no reason given");
  const kind = f.faultcode ?? code?.Value;
  return `${message.slice(0, 300)}${typeof kind === "string" ? ` (${kind})` : ""}`;
}

function failure(status: number, data: unknown): string {
  const fault = (data as { fault?: string } | null)?.fault;
  if (fault) return `The service refused the call: ${fault}`;
  const detail = vendorError(data);
  return `The service answered ${status}${detail ? `: ${detail}` : ""}`;
}

export function soapVendor(connector: Connector, fetchImpl: Fetch = fetch): Vendor {
  const soap = connector.soap!;
  const endpoint = new URL(soap.endpoint);
  const safe = guardedFetch([endpoint.origin], fetchImpl, XML_TYPES, "XML, JSON or text");
  const ns = Object.entries(soap.namespaces)
    .map(([prefix, uri]) => ` xmlns:${prefix}="${escapeXml(uri)}"`)
    .join("");

  // request.body is the Body's content (already escaped), request.headers.soapaction its action.
  const call: Vendor["call"] = async (c, request) => {
    const ws = soap.wsSecurity;
    const header = ws
      ? `<soap:Header>${usernameToken(c.config[ws.username], c.secret[ws.password], ws.digest)}</soap:Header>`
      : "";
    const envelope =
      `<?xml version="1.0" encoding="utf-8"?><soap:Envelope xmlns:soap="${ENVELOPE[soap.version]}"${ns}>` +
      `${header}<soap:Body>${String(request.body)}</soap:Body></soap:Envelope>`;
    const action = request.headers?.soapaction;
    const res = await safe(endpoint, {
      method: "POST",
      headers: {
        accept: soap.version === "1.1" ? "text/xml" : "application/soap+xml",
        ...(soap.auth ? authHeaders(soap.auth, c) : {}),
        ...(soap.version === "1.1"
          ? { "content-type": "text/xml; charset=utf-8", soapaction: `"${action ?? ""}"` }
          : {
              "content-type": `application/soap+xml; charset=utf-8${action ? `; action="${action}"` : ""}`,
            }),
      },
      body: envelope,
    });
    const raw = await res.text();
    let doc: XmlNode | null = null;
    try {
      doc = parseXml(raw);
    } catch (error) {
      if (!(error instanceof XmlError)) throw error;
      if (res.ok)
        throw new ConnectError(`The service's answer is not usable XML: ${error.message}`);
    }
    const body = doc && localName(doc.name) === "Envelope" ? child(doc, "Body") : undefined;
    if (!body) {
      // A gateway's error page, say: its own words, as for any app.
      if (!res.ok) return { status: res.status, data: raw.slice(0, 300) };
      throw new ConnectError("The service answered without a SOAP envelope");
    }
    const fault = child(body, "Fault");
    if (fault) return { status: res.ok ? 500 : res.status, data: { fault: faultText(fault) } };
    return { status: res.status, data: toJson(body) };
  };

  return {
    async connect(fields) {
      const { config, secret } = formValues(connector, fields);
      const conn: VendorConnection = { config: { host: endpoint.host, ...config }, secret };
      return { ...conn, accountLabel: conn.config[connector.accountKey] ?? endpoint.host };
    },
    call,
    checks: soap.verify
      ? {
          // The connector's `verify` call, with the admin's own arguments.
          async operation(c) {
            const verify = soap.verify!;
            const tool = soap.tools.find((t) => t.name === verify.tool)!;
            const res = await call(c, {
              method: "POST",
              path: "",
              body: renderBody(tool, verify.args),
              headers: tool.soapAction ? { soapaction: tool.soapAction } : {},
            });
            if (res.status < 200 || res.status >= 300)
              throw new ConnectError(failure(res.status, res.data));
            const shown = answer(tool, res.data);
            return {
              values: (Array.isArray(shown) ? shown : [shown]).flatMap((item) =>
                pickDisplay(item, verify.show),
              ),
            };
          },
        }
      : undefined,
  };
}

// The tool's part of the Body: its `result` path, a list when it ends in [], and only its
// `fields`.
export function answer(tool: SoapTool, body: unknown): unknown {
  const path = tool.result ?? `${tool.operation}Response`;
  const list = path.endsWith("[]");
  let value: unknown = body;
  for (const key of path.replace(/\[\]$/, "").split(".")) {
    value =
      value && typeof value === "object" && !Array.isArray(value) && Object.hasOwn(value, key)
        ? (value as Record<string, unknown>)[key]
        : undefined;
  }
  if (list) value = value === undefined || value === "" ? [] : [value].flat();
  const keep = tool.fields;
  if (!keep) return value ?? null;
  const trim = (item: unknown) =>
    item && typeof item === "object"
      ? Object.fromEntries(
          keep.flatMap((k) =>
            Object.hasOwn(item, k) ? [[k, (item as Record<string, unknown>)[k]]] : [],
          ),
        )
      : item;
  return Array.isArray(value) ? value.map(trim) : trim(value ?? null);
}

export function soapTools(soap: SoapApp): Toolset {
  return {
    names: soap.tools.map((t) => t.name),
    register(server, ctx, allowed) {
      for (const tool of soap.tools) {
        if (!allowed.has(tool.name)) continue;
        const shape = Object.fromEntries(
          Object.entries(tool.params).map(([name, p]) => {
            const type = zodType[p.type].describe(p.description);
            return [name, p.required ? type : type.optional()];
          }),
        );
        server.registerTool(
          tool.name,
          {
            description: tool.description,
            inputSchema: z.object(shape).strict(),
            annotations: {
              readOnlyHint: tool.readOnly,
              destructiveHint: tool.destructive,
              openWorldHint: true,
            },
          },
          (args: Record<string, unknown>) =>
            audited(ctx, tool.name, JSON.stringify(args).slice(0, 300), async (mark) => {
              paramLimits(ctx.limits, tool.params, args);
              const body = renderBody(tool, args);
              if (Buffer.byteLength(body) > MAX_BODY_BYTES)
                throw new Denied(`The request is larger than ${MAX_BODY_BYTES / 1024} KB`);
              let res;
              try {
                res = await ctx.call({
                  method: "POST",
                  path: "",
                  body,
                  headers: tool.soapAction ? { soapaction: tool.soapAction } : {},
                });
              } catch (error) {
                if (error instanceof ConnectError) throw new VendorFailed(502, error.message);
                throw error;
              }
              mark(res.status);
              if (res.status < 200 || res.status >= 300) {
                throw new VendorFailed(res.status, failure(res.status, res.data));
              }
              const result = answer(tool, res.data);
              for (const a of tool.answerLimits) {
                const values = listLimit(ctx.limits, a.limit);
                if (values && !answerWithin(result, a.field, values))
                  throw new Denied(`The answer's ${a.field} is outside this agent's limits`);
              }
              return json(result);
            }),
        );
      }
    },
  };
}

/** Binds to any connector with a `soap` block; createAgentConnect uses it for those. */
export default {
  id: "soap",
  create: ({ connector }) => {
    const soap = connector.soap!;
    // Each body must be well-formed once filled, so a typo fails at startup, not on a call.
    for (const tool of soap.tools) {
      const sample = tool.body.replace(/\{[a-z][a-z0-9_]*\}/g, "x");
      try {
        parseXml(`<soap:Body xmlns:soap="${ENVELOPE[soap.version]}">${sample}</soap:Body>`);
      } catch (error) {
        throw new Error(
          `${connector.id}: ${tool.name}'s body is not XML: ${(error as Error).message}`,
        );
      }
    }
    return { vendor: soapVendor(connector), toolset: soapTools(soap) };
  },
} satisfies ConnectorModule;
