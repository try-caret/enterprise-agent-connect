import { createHash } from "node:crypto";
import { guardedFetch } from "../http/index.js";
import {
  ConnectError,
  TokenCache,
  type Fetch,
  type Vendor,
  type VendorConnection,
  type VendorRequest,
} from "../../packages/connector-kit/src/types.js";

// SAP S/4HANA's released OData v2 APIs under /sap/opu/odata/sap/<service>/, the same on
// S/4HANA Cloud (a communication user of the services' communication arrangements) and on
// premise through SAP Gateway (a technical user, with the client as sap-client). Basic auth,
// or on Cloud OAuth 2.0 client credentials at /sap/bc/sec/oauth2/token with the
// communication user as the client. Writes carry an x-csrf-token fetched first from the same
// service, with the session cookies it came with.

export const ODATA = "/sap/opu/odata/sap";
const HOST = /^[a-z0-9](?:[a-z0-9-]*[a-z0-9])?(?:\.[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)*(?::\d{1,5})?$/;
const SAP_CLIENT = /^\d{3}$/;
const SERVICE = /^API_[A-Z0-9_]+$/;
// What connect reads to prove the user signs in: 401 is a refused sign-in, anything else
// (even 403 or 404 before the arrangement or service exists) means SAP knew the user.
const SIGN_IN_PROBE = `${ODATA}/API_BUSINESS_PARTNER/`;

export function normalizeHost(input: string): string {
  const raw = input.trim().toLowerCase();
  if (raw.startsWith("http://"))
    throw new ConnectError("SAP must be reached over https. Enter the https address.");
  const host = raw.replace(/^https:\/\//, "").replace(/\/.*$/, "");
  if (!HOST.test(host))
    throw new ConnectError("Enter the API host name, for example my123456-api.s4hana.cloud.sap");
  return host;
}

// The service a path belongs to: the CSRF token is fetched from its root.
export function serviceRoot(path: string): string {
  const m = /^\/sap\/opu\/odata\/sap\/(API_[A-Z0-9_]+)\//.exec(path);
  if (!m) throw new Error(`Not an SAP API path: ${path}`);
  return `${ODATA}/${m[1]}/`;
}

type Sent = { status: number; data: unknown; headers: Headers };

export function sapVendor(fetchImpl: Fetch = fetch): Vendor {
  const tokens = new TokenCache();

  const origin = (c: VendorConnection) => `https://${c.config.host}`;
  const withClient = (c: VendorConnection, path: string) =>
    c.config.sapClient
      ? `${path}${path.includes("?") ? "&" : "?"}sap-client=${c.config.sapClient}`
      : path;
  const basic = (c: VendorConnection) =>
    `Basic ${Buffer.from(`${c.config.user}:${c.secret.password}`).toString("base64")}`;
  const cacheKey = (c: VendorConnection) =>
    createHash("sha256")
      .update(`${c.config.host}\n${c.config.user}\n${c.secret.password}`)
      .digest("hex");

  async function reach(c: VendorConnection, path: string, init: RequestInit): Promise<Response> {
    try {
      return await guardedFetch([origin(c)], fetchImpl)(`${origin(c)}${withClient(c, path)}`, init);
    } catch (error) {
      if (error instanceof ConnectError) throw error;
      throw new ConnectError(`${c.config.host} could not be reached from this service`);
    }
  }

  async function mint(c: VendorConnection) {
    const res = await reach(c, "/sap/bc/sec/oauth2/token", {
      method: "POST",
      headers: {
        authorization: basic(c),
        accept: "application/json",
        "content-type": "application/x-www-form-urlencoded",
      },
      body: new URLSearchParams({ grant_type: "client_credentials" }),
    });
    const body = parse(await res.text()) as {
      access_token?: unknown;
      expires_in?: unknown;
      error_description?: unknown;
      error?: unknown;
    } | null;
    if (!res.ok || typeof body?.access_token !== "string") {
      const why = typeof body?.error_description === "string" ? body.error_description : null;
      throw new ConnectError(
        `SAP refused the sign-in (${res.status})${why ? `: ${why.slice(0, 300)}` : ""}`,
      );
    }
    const expiresIn = Number(body.expires_in);
    if (!Number.isFinite(expiresIn) || expiresIn <= 60)
      throw new ConnectError("SAP answered the sign-in without a usable token lifetime");
    return { token: body.access_token, expiresIn };
  }

  async function authorization(c: VendorConnection) {
    return c.config.auth === "oauth"
      ? `Bearer ${await tokens.get(cacheKey(c), () => mint(c))}`
      : basic(c);
  }

  async function send(
    c: VendorConnection,
    request: VendorRequest,
    extra: Record<string, string> = {},
  ): Promise<Sent> {
    const res = await reach(c, request.path, {
      method: request.method,
      headers: {
        authorization: await authorization(c),
        accept: "application/json",
        "accept-language": "en",
        ...extra,
        ...(request.body === undefined ? {} : { "content-type": "application/json" }),
      },
      body: request.body === undefined ? undefined : JSON.stringify(request.body),
    });
    const raw = await res.text();
    return { status: res.status, data: raw ? parse(raw) : null, headers: res.headers };
  }

  // The token is bound to the session the cookies name, so both go with the write. A refused
  // fetch is answered as is (a 401 is retried with a new OAuth token).
  async function once(c: VendorConnection, request: VendorRequest): Promise<Sent> {
    if (request.method === "GET") return send(c, request);
    const res = await send(
      c,
      { method: "GET", path: serviceRoot(request.path) },
      { "x-csrf-token": "Fetch" },
    );
    if (res.status < 200 || res.status >= 300) return res;
    const token = res.headers.get("x-csrf-token");
    if (!token || token.toLowerCase() === "required")
      throw new ConnectError("SAP did not hand out a CSRF token for the change");
    const cookie = res.headers
      .getSetCookie()
      .map((s) => s.split(";")[0])
      .join("; ");
    return send(c, request, { "x-csrf-token": token, ...(cookie ? { cookie } : {}) });
  }

  return {
    async connect(fields) {
      const deployment = fields.deployment === "on_premise" ? "on_premise" : "cloud";
      const onPremise = deployment === "on_premise";
      const host = normalizeHost((onPremise ? fields.gateway_host : fields.host) ?? "");
      // Not asked on Cloud, which has no client to choose.
      const sapClient = onPremise ? (fields.sap_client?.trim() ?? "") : "";
      if (sapClient && !SAP_CLIENT.test(sapClient))
        throw new ConnectError("The client is three digits, for example 100");
      const user = fields.user?.trim();
      const password = fields.password ?? "";
      if (!user || !password) throw new ConnectError("Enter the user name and password");
      const auth = deployment === "cloud" && fields.auth === "oauth" ? "oauth" : "basic";
      const system = sapClient ? `${host} client ${sapClient}` : host;
      const conn = {
        config: { deployment, auth, host, ...(sapClient ? { sapClient } : {}), user, system },
        secret: { password },
      };
      if (auth === "oauth") tokens.forget(cacheKey(conn));
      const res = await send(conn, { method: "GET", path: SIGN_IN_PROBE });
      if (res.status === 401)
        throw new ConnectError(`SAP did not accept the user name and password for ${user}`);
      return { ...conn, accountLabel: system };
    },

    async call(c, request) {
      let res = await once(c, request);
      if (res.status === 401 && c.config.auth === "oauth") {
        tokens.forget(cacheKey(c));
        res = await once(c, request);
      }
      return { status: res.status, data: res.data };
    },

    checks: {
      // Which of the ticked capabilities' services this user can open (their service roots).
      async services(c, { expected }) {
        const reached: string[] = [];
        const missing: string[] = [];
        for (const service of expected.filter((s) => SERVICE.test(s))) {
          const res = await send(c, { method: "GET", path: `${ODATA}/${service}/` });
          if (res.status >= 200 && res.status < 300) reached.push(service);
          else
            missing.push(
              `${service} (${res.status === 404 ? "not active" : res.status === 403 ? "not allowed for this user" : `answered ${res.status}`})`,
            );
        }
        return {
          values: reached,
          ...(missing.length
            ? {
                warning: `The user cannot use these services yet: ${missing.join(", ")}. Follow the setup steps for them.`,
              }
            : {}),
        };
      },
    },
  };
}

function parse(raw: string): unknown {
  try {
    return JSON.parse(raw);
  } catch {
    return raw;
  }
}
