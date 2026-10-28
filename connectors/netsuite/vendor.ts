import {
  constants,
  createHash,
  createHmac,
  createPrivateKey,
  randomBytes,
  randomUUID,
  sign,
} from "node:crypto";
import {
  ConnectError,
  TokenCache,
  type Fetch,
  type Vendor,
  type VendorConnection,
  type VendorRequest,
  type VendorResponse,
} from "../../packages/connector-kit/src/types.js";

// Two sign-ins, both as one NetSuite role: OAuth 2.0 client credentials (a JWT signed with
// the private key of the certificate uploaded to NetSuite's M2M setup, PS256, kid = its
// certificate id) or token-based authentication (OAuth 1.0a, HMAC-SHA256, realm = account).
// REST calls go to <account>.suitetalk.api.netsuite.com, the saved search script to
// <account>.restlets.api.netsuite.com, nowhere else.

export const RESTLET_PATH = "/app/site/hosting/restlet.nl";
export const SUITEQL_PATH = "/services/rest/query/v1/suiteql";
const TOKEN_PATH = "/services/rest/auth/oauth2/v1/token";
const ACCOUNT = /^[A-Z0-9]+(_[A-Z0-9]+)?$/;
const MIN_RSA_BITS = 3072;

export function normalizeAccount(input: string): string {
  const account = input
    .trim()
    .toUpperCase()
    .replace(/^HTTPS?:\/\//, "")
    .replace(/\..*$/, "")
    .replaceAll("-", "_");
  if (!ACCOUNT.test(account))
    throw new ConnectError(
      "Enter the NetSuite account ID, for example 1234567 or 1234567_SB1 for a sandbox",
    );
  return account;
}

// Hosts take the account id in lowercase with a dash: 1234567_SB1 → 1234567-sb1.
const host = (account: string, kind: "suitetalk" | "restlets") => {
  if (!ACCOUNT.test(account)) throw new Error("NetSuite connection has no account id");
  return `https://${account.toLowerCase().replaceAll("_", "-")}.${kind}.api.netsuite.com`;
};

// A pasted PEM can lose its line breaks in a form field; they are put back before parsing.
export function readPrivateKey(pem: string) {
  const m = /-----BEGIN ([A-Z ]+)-----([\s\S]+?)-----END \1-----/.exec(pem.trim());
  if (!m) throw new ConnectError("Paste the whole private key, from -----BEGIN to -----END");
  const body = m[2]
    .replace(/\s+/g, "")
    .match(/.{1,64}/g)!
    .join("\n");
  let key;
  try {
    key = createPrivateKey(`-----BEGIN ${m[1]}-----\n${body}\n-----END ${m[1]}-----\n`);
  } catch {
    throw new ConnectError("The private key could not be read. Paste private.pem as it is");
  }
  const bits = key.asymmetricKeyDetails?.modulusLength ?? 0;
  if (key.asymmetricKeyType !== "rsa" || bits < MIN_RSA_BITS)
    throw new ConnectError(`NetSuite needs an RSA key of at least ${MIN_RSA_BITS} bits`);
  return key;
}

export function clientAssertion(input: {
  clientId: string;
  certificateId: string;
  privateKey: string;
  tokenUrl: string;
  scope: string;
  now?: number;
}): string {
  const b64 = (v: object) => Buffer.from(JSON.stringify(v)).toString("base64url");
  const now = Math.floor((input.now ?? Date.now()) / 1000);
  const header = b64({ alg: "PS256", typ: "JWT", kid: input.certificateId });
  const payload = b64({
    iss: input.clientId,
    scope: [input.scope],
    aud: input.tokenUrl,
    iat: now,
    exp: now + 300,
    jti: randomUUID(),
  });
  const signature = sign("sha256", Buffer.from(`${header}.${payload}`), {
    key: readPrivateKey(input.privateKey),
    padding: constants.RSA_PKCS1_PSS_PADDING,
    saltLength: constants.RSA_PSS_SALTLEN_DIGEST,
  });
  return `${header}.${payload}.${signature.toString("base64url")}`;
}

const enc = (v: string) =>
  encodeURIComponent(v).replace(
    /[!'()*]/g,
    (c) => `%${c.charCodeAt(0).toString(16).toUpperCase()}`,
  );

// OAuth 1.0a (RFC 5849 section 3.4): the query's parameters are signed with the oauth ones.
export function tbaHeader(input: {
  method: string;
  url: string;
  realm: string;
  consumerKey: string;
  consumerSecret: string;
  tokenId: string;
  tokenSecret: string;
  nonce?: string;
  timestamp?: number;
}): string {
  const url = new URL(input.url);
  const oauth: Record<string, string> = {
    oauth_consumer_key: input.consumerKey,
    oauth_nonce: input.nonce ?? randomBytes(16).toString("hex"),
    oauth_signature_method: "HMAC-SHA256",
    oauth_timestamp: String(input.timestamp ?? Math.floor(Date.now() / 1000)),
    oauth_token: input.tokenId,
    oauth_version: "1.0",
  };
  const params = [...url.searchParams, ...Object.entries(oauth)]
    .map(([k, v]) => [enc(k), enc(v)])
    .sort(([a, x], [b, y]) => (a === b ? (x < y ? -1 : 1) : a < b ? -1 : 1))
    .map(([k, v]) => `${k}=${v}`)
    .join("&");
  const base = [input.method, enc(`${url.origin}${url.pathname}`), enc(params)].join("&");
  const signature = createHmac("sha256", `${enc(input.consumerSecret)}&${enc(input.tokenSecret)}`)
    .update(base)
    .digest("base64");
  const fields = { ...oauth, oauth_signature: signature };
  return `OAuth realm="${input.realm}", ${Object.entries(fields)
    .map(([k, v]) => `${k}="${enc(v)}"`)
    .join(", ")}`;
}

// NetSuite REST errors: {title, status, "o:errorDetails": [{detail, "o:errorCode"}]}; token
// errors: {error}.
export function netsuiteError(data: unknown): string {
  if (!data || typeof data !== "object") return typeof data === "string" ? data.slice(0, 300) : "";
  const d = data as {
    title?: unknown;
    error?: unknown;
    error_description?: unknown;
    "o:errorDetails"?: { detail?: string }[];
  };
  // The saved search script's errors: {error: {code, message}}.
  const error =
    d.error && typeof d.error === "object" ? (d.error as { message?: unknown }).message : d.error;
  const details = (d["o:errorDetails"] ?? [])
    .map((e) => e.detail)
    .filter(Boolean)
    .join("; ");
  const text = details || d.error_description || d.title || error;
  return typeof text === "string" ? text.split(/\r?\n/)[0].slice(0, 500) : "";
}

const parse = (raw: string): unknown => {
  try {
    return JSON.parse(raw);
  } catch {
    return raw;
  }
};

export function netsuiteVendor(fetchImpl: Fetch = fetch): Vendor {
  const tokens = new TokenCache();

  async function mint(c: VendorConnection, scope: string) {
    const tokenUrl = `${host(c.config.accountId, "suitetalk")}${TOKEN_PATH}`;
    const res = await fetchImpl(tokenUrl, {
      method: "POST",
      redirect: "error",
      headers: { "content-type": "application/x-www-form-urlencoded", accept: "application/json" },
      body: new URLSearchParams({
        grant_type: "client_credentials",
        client_assertion_type: "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
        client_assertion: clientAssertion({
          clientId: c.secret.client_id,
          certificateId: c.secret.certificate_id,
          privateKey: c.secret.private_key,
          tokenUrl,
          scope,
        }),
      }),
    });
    const body = parse(await res.text()) as { access_token?: unknown; expires_in?: unknown };
    if (!res.ok)
      throw new ConnectError(
        `NetSuite refused the sign-in: ${netsuiteError(body) || res.status}. Check the client ID, certificate ID and that the integration allows this scope`,
      );
    const expiresIn = Number(body?.expires_in);
    if (typeof body?.access_token !== "string" || !(expiresIn > 60))
      throw new ConnectError("NetSuite answered the sign-in without a usable token");
    return { token: body.access_token, expiresIn };
  }

  const cacheKey = (c: VendorConnection, scope: string) =>
    createHash("sha256")
      .update(
        [
          c.config.accountId,
          scope,
          c.secret.client_id,
          c.secret.certificate_id,
          c.secret.private_key,
        ].join("\n"),
      )
      .digest("hex");

  async function send(c: VendorConnection, request: VendorRequest, fresh = false) {
    const restlet = request.path.startsWith(RESTLET_PATH);
    const url = `${host(c.config.accountId, restlet ? "restlets" : "suitetalk")}${request.path}`;
    const scope = restlet ? "restlets" : "rest_webservices";
    let authorization: string;
    if (c.config.auth === "tba")
      authorization = tbaHeader({
        method: request.method,
        url,
        realm: c.config.accountId,
        consumerKey: c.secret.client_id,
        consumerSecret: c.secret.client_secret,
        tokenId: c.secret.token_id,
        tokenSecret: c.secret.token_secret,
      });
    else {
      if (fresh) tokens.forget(cacheKey(c, scope));
      authorization = `Bearer ${await tokens.get(cacheKey(c, scope), () => mint(c, scope))}`;
    }
    const res = await fetchImpl(url, {
      method: request.method,
      redirect: "error",
      headers: {
        authorization,
        accept: "application/json",
        // SuiteQL refuses a query without it.
        ...(request.path.startsWith(SUITEQL_PATH) ? { prefer: "transient" } : {}),
        ...request.headers,
        ...(request.body === undefined ? {} : { "content-type": "application/json" }),
      },
      body: request.body === undefined ? undefined : JSON.stringify(request.body),
    });
    const raw = await res.text();
    // A record made through REST answers 204 with its address in Location.
    const location = res.headers.get("location");
    const data = raw ? parse(raw) : location ? { location } : null;
    return { status: res.status, data } as VendorResponse;
  }

  async function call(c: VendorConnection, request: VendorRequest) {
    const res = await send(c, request);
    return res.status === 401 && c.config.auth !== "tba" ? send(c, request, true) : res;
  }

  return {
    async connect(fields) {
      const accountId = normalizeAccount(fields.account_id ?? "");
      const auth = fields.auth === "tba" ? "tba" : "oauth2";
      const value = (id: string) => fields[id]?.trim() ?? "";
      const needed =
        auth === "tba"
          ? ["client_id", "client_secret", "token_id", "token_secret"]
          : ["client_id", "certificate_id", "private_key"];
      const missing = needed.filter((id) => !value(id));
      if (missing.length)
        throw new ConnectError(
          auth === "tba"
            ? "Enter the consumer key and secret and the token ID and secret"
            : "Enter the client ID, the certificate ID and the private key",
        );
      if (auth === "oauth2") readPrivateKey(value("private_key"));
      const connection = {
        config: { accountId, auth },
        secret: Object.fromEntries(needed.map((id) => [id, value(id)])),
      };
      const res = await call(connection, {
        method: "POST",
        path: `${SUITEQL_PATH}?limit=1`,
        body: { q: "SELECT 1 AS ok FROM DUAL" },
      });
      if (res.status < 200 || res.status >= 300)
        throw new ConnectError(
          `NetSuite account ${accountId} refused the call (${res.status}${netsuiteError(res.data) ? `: ${netsuiteError(res.data)}` : ""}). Check the role's REST Web Services and sign-in permissions`,
        );
      return { ...connection, accountLabel: `NetSuite ${accountId}` };
    },

    call,

    checks: {
      async account(connection) {
        const res = await call(connection, {
          method: "POST",
          path: `${SUITEQL_PATH}?limit=1`,
          body: { q: "SELECT 1 AS ok FROM DUAL" },
        });
        if (res.status < 200 || res.status >= 300)
          return {
            values: [],
            warning: `NetSuite answered ${res.status}: ${netsuiteError(res.data) || "no reason given"}`,
          };
        return { values: [connection.config.accountId] };
      },
    },
  };
}
