import { createHash } from "node:crypto";
import {
  ConnectError,
  TokenCache,
  type Fetch,
  type VendorRequest,
  type VendorResponse,
} from "./types.js";

// OAuth 2.0 client credentials (form-encoded, as FedEx and Dell document it), with the token
// kept until a minute before its expires_in and minted again once after a 401. The token is
// never logged or returned: refusals carry only the vendor's own error text.

type Credentials = { tokenUrl: string; clientId: string; clientSecret: string };

// FedEx: { errors: [{ code, message }] }; Dell and most others: { error, error_description }.
export function vendorError(data: unknown): string {
  if (!data || typeof data !== "object") return typeof data === "string" ? data.slice(0, 300) : "";
  const d = data as {
    errors?: { code?: string; message?: string }[];
    error?: unknown;
    error_description?: unknown;
    message?: unknown;
  };
  if (Array.isArray(d.errors))
    return d.errors
      .map((e) => [e.code, e.message].filter(Boolean).join(": "))
      .join("; ")
      .slice(0, 500);
  const text = d.error_description ?? d.message ?? d.error;
  return typeof text === "string" ? text.split(/\r?\n/)[0].slice(0, 300) : "";
}

export function clientCredentials(name: string, fetchImpl: Fetch) {
  const tokens = new TokenCache();
  const key = (c: Credentials) =>
    createHash("sha256").update(`${c.tokenUrl}\n${c.clientId}\n${c.clientSecret}`).digest("hex");

  async function mint(c: Credentials) {
    const res = await fetchImpl(c.tokenUrl, {
      method: "POST",
      headers: { "content-type": "application/x-www-form-urlencoded", accept: "application/json" },
      body: new URLSearchParams({
        grant_type: "client_credentials",
        client_id: c.clientId,
        client_secret: c.clientSecret,
      }),
    });
    const raw = await res.text();
    const body = parse(raw) as { access_token?: unknown; expires_in?: unknown } | string | null;
    if (!res.ok) {
      const why = vendorError(body) || String(res.status);
      throw new ConnectError(`${name} refused the sign-in: ${why}`);
    }
    if (typeof body !== "object" || !body || typeof body.access_token !== "string")
      throw new ConnectError(`${name} answered the sign-in without a token`);
    const expiresIn = Number(body.expires_in);
    if (!Number.isFinite(expiresIn) || expiresIn <= 60)
      throw new ConnectError(`${name} answered the sign-in without a usable token lifetime`);
    return { token: body.access_token, expiresIn };
  }

  const token = (c: Credentials) => tokens.get(key(c), () => mint(c));

  async function send(base: string, access: string, request: VendorRequest, headers = {}) {
    // A string body (e.g. Jamf's Classic API XML) goes as is, under the request's own
    // content-type; anything else is JSON.
    const asIs = typeof request.body === "string";
    const res = await fetchImpl(`${base}${request.path}`, {
      method: request.method,
      headers: {
        authorization: `Bearer ${access}`,
        accept: "application/json",
        ...headers,
        ...(request.body === undefined || asIs ? {} : { "content-type": "application/json" }),
        ...request.headers,
      },
      body:
        request.body === undefined
          ? undefined
          : asIs
            ? (request.body as string)
            : JSON.stringify(request.body),
    });
    const raw = await res.text();
    return {
      status: res.status,
      data: raw ? parse(raw) : null,
      headers: Object.fromEntries(res.headers),
    };
  }

  return {
    // Proves the credentials by minting a token (and keeps it for the calls that follow).
    signIn: (c: Credentials) => token(c).then(() => undefined),
    async call(
      c: Credentials,
      base: string,
      request: VendorRequest,
      headers: Record<string, string> = {},
    ): Promise<VendorResponse> {
      let res = await send(base, await token(c), request, headers);
      if (res.status === 401) {
        tokens.forget(key(c));
        res = await send(base, await token(c), request, headers);
      }
      return res;
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
