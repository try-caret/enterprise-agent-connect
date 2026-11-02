import { createHash, createPrivateKey, randomUUID, sign, type KeyObject } from "node:crypto";
import { vendorFetch, vendorUrl } from "../../packages/connector-kit/src/fetch.js";
import { vendorError } from "../../packages/connector-kit/src/oauth.js";
import {
  ConnectError,
  TokenCache,
  type CheckInput,
  type Fetch,
  type Vendor,
  type VendorConnection,
  type VendorRequest,
} from "../../packages/connector-kit/src/types.js";

// An Okta API service app: OAuth 2.0 client credentials from the org authorization server
// (`/oauth2/v1/token`), the app proving itself with a private_key_jwt client assertion (RS256,
// iss = sub = client id, aud = the token URL, 5 minutes). Okta has no catch-all scope, and a
// token request naming a scope the app was not granted is refused whole, so every call gets a
// token for the one scope its path and method need (`scopeFor`). Lists come back as
// `{items, after}`: Okta pages with a Link header, whose `after` is the next page's cursor.
// Sources: developer.okta.com/docs/guides/implement-oauth-for-okta-serviceapp,
// developer.okta.com/docs/api/openapi/okta-oauth/guides/client-auth (private_key_jwt) and
// developer.okta.com/docs/api/#pagination (read October 27, 2026).

// The org's own Okta address; an admin address (acme-admin.okta.com) is its org's too.
const ORG_HOST = /^[a-z0-9][a-z0-9-]*\.(okta|oktapreview|okta-emea|okta-gov)\.com$/;
const CLIENT_ID = /^[A-Za-z0-9_-]{8,64}$/;
const KEY_ID = /^[A-Za-z0-9_.-]{1,128}$/;
const RESOURCES = new Set(["users", "groups", "apps", "logs"]);

export function normalizeOrg(input: string): string {
  let host = input
    .trim()
    .toLowerCase()
    .replace(/^https?:\/\//, "")
    .replace(/\/.*$/, "");
  if (/^[a-z0-9][a-z0-9-]*$/.test(host)) host = `${host}.okta.com`;
  host = host.replace(/-admin\./, ".");
  if (!ORG_HOST.test(host))
    throw new ConnectError("Enter your Okta address, such as your-company.okta.com");
  return host;
}

/** The scope a Management API call needs: okta.<resource>.read for a GET, .manage otherwise. */
export function scopeFor(request: VendorRequest): string {
  const resource = /^\/api\/v1\/([a-z]+)(?:[/?]|$)/.exec(request.path)?.[1];
  if (!resource || !RESOURCES.has(resource)) throw new Error(`No Okta scope for ${request.path}`);
  return `okta.${resource}.${request.method === "GET" ? "read" : "manage"}`;
}

// Okta's console offers the generated key as PEM or as a JWK; either is kept as PKCS#8 PEM.
export function parsePrivateKey(raw: string): string {
  let key: KeyObject;
  try {
    const text = raw.trim();
    key = text.startsWith("{")
      ? createPrivateKey({ key: JSON.parse(text), format: "jwk" })
      : createPrivateKey(text);
  } catch {
    throw new ConnectError("The private key could not be read. Paste the whole key Okta showed.");
  }
  if (key.asymmetricKeyType !== "rsa")
    throw new ConnectError("Okta signs in with an RSA key. Generate an RSA key in the app.");
  return key.export({ type: "pkcs8", format: "pem" }).toString();
}

export function clientAssertion(input: {
  clientId: string;
  // Left out when empty: Okta then tries each of the app's keys.
  keyId: string;
  privateKey: string;
  tokenUrl: string;
}): string {
  const now = Math.floor(Date.now() / 1000);
  const part = (o: object) => Buffer.from(JSON.stringify(o)).toString("base64url");
  const signed = `${part({ alg: "RS256", typ: "JWT", kid: input.keyId || undefined })}.${part({
    iss: input.clientId,
    sub: input.clientId,
    aud: input.tokenUrl,
    iat: now,
    exp: now + 300,
    jti: randomUUID(),
  })}`;
  const signature = sign("sha256", Buffer.from(signed), input.privateKey).toString("base64url");
  return `${signed}.${signature}`;
}

/** The `after` cursor of a Link header's rel="next" page. */
export function nextAfter(link: string | null): string | null {
  for (const part of (link ?? "").split(",")) {
    const m = /<([^>]+)>\s*;\s*rel="next"/.exec(part);
    if (m) return URL.parse(m[1])?.searchParams.get("after") ?? null;
  }
  return null;
}

function safeJson(raw: string): unknown {
  try {
    return JSON.parse(raw);
  } catch {
    return raw;
  }
}

class ScopeRefused extends ConnectError {}

export function oktaVendor(unguarded: Fetch = fetch): Vendor {
  const fetchImpl = vendorFetch("Okta", unguarded);
  const tokens = new TokenCache();

  function host(c: VendorConnection) {
    const org = c.config.org ?? "";
    if (!ORG_HOST.test(org)) throw new Error("Okta connection has no org");
    return `https://${org}`;
  }

  async function mint(c: VendorConnection, scope: string) {
    const tokenUrl = `${host(c)}/oauth2/v1/token`;
    const res = await fetchImpl(tokenUrl, {
      method: "POST",
      headers: { "content-type": "application/x-www-form-urlencoded", accept: "application/json" },
      body: new URLSearchParams({
        grant_type: "client_credentials",
        scope,
        client_assertion_type: "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
        client_assertion: clientAssertion({
          clientId: c.config.clientId,
          keyId: c.config.keyId,
          privateKey: c.secret.private_key,
          tokenUrl,
        }),
      }),
    });
    const body = safeJson(await res.text()) as {
      access_token?: unknown;
      expires_in?: unknown;
      error?: unknown;
    };
    if (!res.ok || typeof body?.access_token !== "string") {
      const why = vendorError(body) || String(res.status);
      if (body?.error === "invalid_scope")
        throw new ScopeRefused(`Okta has not granted the app ${scope}: ${why}`);
      if (/dpop/i.test(why))
        throw new ConnectError(
          `Okta wants a proof-of-possession header (DPoP). In the app's General tab, edit Client Credentials and clear "Require Demonstrating Proof of Possession (DPoP) header in token requests". (${why})`,
        );
      throw new ConnectError(`Okta refused the app's sign-in: ${why}`);
    }
    return { token: body.access_token, expiresIn: Number(body.expires_in) || 3600 };
  }

  // Per credential and scope, so a new key never reuses an old token.
  const cacheKey = (c: VendorConnection, scope: string) =>
    createHash("sha256")
      .update(
        [c.config.org, c.config.clientId, c.config.keyId, c.secret.private_key, scope].join("\n"),
      )
      .digest("hex");

  async function send(c: VendorConnection, token: string, request: VendorRequest) {
    const res = await fetchImpl(vendorUrl(host(c), request.path), {
      method: request.method,
      headers: {
        authorization: `Bearer ${token}`,
        accept: "application/json",
        ...request.headers,
        ...(request.body === undefined ? {} : { "content-type": "application/json" }),
      },
      body: request.body === undefined ? undefined : JSON.stringify(request.body),
    });
    const raw = await res.text();
    const data = raw ? safeJson(raw) : null;
    return {
      status: res.status,
      headers: Object.fromEntries(res.headers),
      data: Array.isArray(data) ? { items: data, after: nextAfter(res.headers.get("link")) } : data,
    };
  }

  return {
    async connect(fields) {
      const org = normalizeOrg(fields.org ?? "");
      const clientId = fields.client_id?.trim() ?? "";
      const keyId = fields.key_id?.trim() ?? "";
      if (!CLIENT_ID.test(clientId))
        throw new ConnectError("Enter the app's Client ID, from its General tab");
      if (keyId && !KEY_ID.test(keyId))
        throw new ConnectError(
          "Enter the key's KID, shown beside the key in the app's General tab",
        );
      const connection = {
        config: { org, clientId, keyId },
        secret: { private_key: parsePrivateKey(fields.private_key ?? "") },
      };
      // okta.users.read is every connection's least: proves the key, the app and the org.
      const scope = "okta.users.read";
      await tokens.get(cacheKey(connection, scope), () => mint(connection, scope));
      return { ...connection, accountLabel: org };
    },

    async call(connection, request) {
      const scope = scopeFor(request);
      const key = cacheKey(connection, scope);
      const token = () => tokens.get(key, () => mint(connection, scope));
      let res = await send(connection, await token(), request);
      if (res.status === 401) {
        tokens.forget(key);
        res = await send(connection, await token(), request);
      }
      return res;
    },

    checks: {
      // One token request per scope the ceiling needs: Okta answers invalid_scope for one
      // not granted. Extra grants are the admin's choice, so only missing ones are named.
      async scopes(connection, input: CheckInput) {
        const granted: string[] = [];
        for (const scope of input.expected)
          try {
            await mint(connection, scope);
            granted.push(scope);
          } catch (error) {
            if (!(error instanceof ScopeRefused)) throw error;
          }
        const missing = input.expected.filter((s) => !granted.includes(s));
        return {
          values: granted,
          ...(missing.length
            ? {
                warning: `Missing: ${missing.join(", ")}. Grant ${missing.length > 1 ? "them" : "it"} on the app's Okta API Scopes tab, or agents cannot do what you ticked.`,
              }
            : {}),
        };
      },
    },
  };
}
