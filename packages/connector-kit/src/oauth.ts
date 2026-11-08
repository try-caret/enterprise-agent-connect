import { createHash, createPrivateKey, sign } from "node:crypto";
import { vendorFetch, vendorUrl } from "./fetch.js";
import {
  ConnectError,
  TokenCache,
  type Fetch,
  type VendorRequest,
  type VendorResponse,
} from "./types.js";

// OAuth 2.0 client credentials (form-encoded, as FedEx and Dell document it) or a JWT bearer
// assertion, with the token kept until a minute before its expires_in and minted again once
// after a 401. The token is never logged or returned: refusals carry only the vendor's own
// error text.

// `scope`: for token services that need one, e.g. Microsoft Entra's v2 endpoint. `clientAuth`
// "basic" sends the client id and secret as HTTP Basic (RFC 6749 2.3.1) instead of in the body;
// `audience`: for token services that ask for one, e.g. Auth0's.
type Credentials =
  | {
      tokenUrl: string;
      clientId: string;
      clientSecret: string;
      scope?: string;
      clientAuth?: "body" | "basic";
      audience?: string;
    }
  | JwtBearer;

// RFC 7523 JWT bearer grant (DocuSign's "JWT Grant"): an RS256 assertion signed with the
// app's RSA private key stands in for a client secret.
export type JwtBearer = {
  tokenUrl: string;
  issuer: string;
  subject: string;
  audience: string;
  scope?: string;
  privateKey: string;
};

const ASSERTION_SECONDS = 3600;
const MIN_RSA_BITS = 2048;

// A pasted PEM can lose its line breaks in a form field; they are put back before parsing.
export function rsaKey(pem: string) {
  const m = /-----BEGIN ([A-Z ]+)-----([\s\S]+?)-----END \1-----/.exec(pem.trim());
  if (!m) throw new ConnectError("Paste the whole private key, from -----BEGIN to -----END");
  const body = m[2].replace(/\s+/g, "").match(/.{1,64}/g)?.join("\n") ?? "";
  let key;
  try {
    key = createPrivateKey(`-----BEGIN ${m[1]}-----\n${body}\n-----END ${m[1]}-----\n`);
  } catch {
    throw new ConnectError("The private key could not be read. Paste it as it was downloaded");
  }
  if (key.asymmetricKeyType !== "rsa" || (key.asymmetricKeyDetails?.modulusLength ?? 0) < MIN_RSA_BITS)
    throw new ConnectError(`The private key must be an RSA key of at least ${MIN_RSA_BITS} bits`);
  return key;
}

export function jwtAssertion(j: Omit<JwtBearer, "tokenUrl">, now = Date.now()): string {
  const part = (v: object) => Buffer.from(JSON.stringify(v)).toString("base64url");
  const iat = Math.floor(now / 1000);
  const signed = `${part({ alg: "RS256", typ: "JWT" })}.${part({
    iss: j.issuer,
    sub: j.subject,
    aud: j.audience,
    iat,
    exp: iat + ASSERTION_SECONDS,
    ...(j.scope ? { scope: j.scope } : {}),
  })}`;
  return `${signed}.${sign("sha256", Buffer.from(signed), rsaKey(j.privateKey)).toString("base64url")}`;
}

// The token request's headers and form for each grant.
function grant(c: Credentials): { headers: Record<string, string>; form: Record<string, string> } {
  if ("privateKey" in c)
    return {
      headers: {},
      form: { grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer", assertion: jwtAssertion(c) },
    };
  const basic = c.clientAuth === "basic";
  // RFC 6749 2.3.1: each form-encoded before joining.
  const enc = (v: string) => encodeURIComponent(v).replace(/%20/g, "+");
  const pair = Buffer.from(`${enc(c.clientId)}:${enc(c.clientSecret)}`).toString("base64");
  return {
    headers: basic ? { authorization: `Basic ${pair}` } : {},
    form: {
      grant_type: "client_credentials",
      ...(basic ? {} : { client_id: c.clientId, client_secret: c.clientSecret }),
      ...(c.scope ? { scope: c.scope } : {}),
      ...(c.audience ? { audience: c.audience } : {}),
    },
  };
}

// FedEx: { errors: [{ code, message }] }; JSON:API: { errors: [{ code, detail }] }; Dell and
// most others: { error, error_description }.
export function vendorError(data: unknown): string {
  if (!data || typeof data !== "object") return typeof data === "string" ? data.slice(0, 300) : "";
  const d = data as {
    errors?: { code?: string; message?: string; detail?: string }[];
    error?: unknown;
    error_description?: unknown;
    message?: unknown;
  };
  if (Array.isArray(d.errors))
    return d.errors
      .map((e) => [e.code, e.message ?? e.detail].filter(Boolean).join(": "))
      .join("; ")
      .slice(0, 500);
  const text = d.error_description ?? d.message ?? d.error;
  return typeof text === "string" ? text.split(/\r?\n/)[0].slice(0, 300) : "";
}

export function clientCredentials(name: string, unguarded: Fetch) {
  const fetchImpl = vendorFetch(name, unguarded);
  const tokens = new TokenCache();
  const key = (c: Credentials) =>
    createHash("sha256")
      .update(JSON.stringify(Object.entries(c).sort()))
      .digest("hex");

  async function mint(c: Credentials) {
    const { headers, form } = grant(c);
    const res = await fetchImpl(c.tokenUrl, {
      method: "POST",
      headers: {
        "content-type": "application/x-www-form-urlencoded",
        accept: "application/json",
        ...headers,
      },
      body: new URLSearchParams(form),
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
    const res = await fetchImpl(vendorUrl(base, request.path), {
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
      data: request.responseType === "text" && res.ok ? raw : raw ? parse(raw) : null,
      headers: Object.fromEntries(res.headers),
    };
  }

  return {
    // Proves the credentials by minting a token (and keeps it for the calls that follow).
    signIn: (c: Credentials) => token(c).then(() => undefined),
    // For a caller sending its own requests: the token, and dropping it after a 401.
    token,
    forget: (c: Credentials) => tokens.forget(key(c)),
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
