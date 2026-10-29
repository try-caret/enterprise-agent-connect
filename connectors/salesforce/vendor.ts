import { createHash } from "node:crypto";
import { vendorFetch, vendorUrl } from "../../packages/connector-kit/src/fetch.js";
import {
  ConnectError,
  TokenCache,
  type Fetch,
  type Vendor,
  type VendorConnection,
  type VendorRequest,
} from "../../packages/connector-kit/src/types.js";

// The customer's External Client App, client credentials flow, running as their integration
// user. Tokens come from their My Domain, and the API is called on the instance_url it returns.

const MY_DOMAIN = /^[a-z0-9][a-z0-9-]*(\.[a-z0-9-]+)?\.my\.salesforce\.com$/;
const USER_ID = /^005[A-Za-z0-9]{12}(?:[A-Za-z0-9]{3})?$/;
// Salesforce does not say when a token expires (it follows the org's session timeout, at
// least 15 minutes), so a token is reused for 10 minutes (the cache drops it a minute before
// expiresIn) and dropped early on a 401.
const TOKEN_SECONDS = 11 * 60;

export function normalizeMyDomain(input: string): string {
  let host = input
    .trim()
    .toLowerCase()
    .replace(/^https?:\/\//, "")
    .replace(/\/.*$/, "");
  if (/^[a-z0-9][a-z0-9-]*$/.test(host)) host = `${host}.my.salesforce.com`;
  if (!MY_DOMAIN.test(host))
    throw new ConnectError("Enter your My Domain, the address ending in .my.salesforce.com");
  return host;
}

function instanceUrl(raw: unknown): string {
  const url = typeof raw === "string" ? URL.parse(raw) : null;
  if (!url || url.protocol !== "https:" || !MY_DOMAIN.test(url.hostname) || url.port)
    throw new ConnectError("Salesforce returned an unexpected address for your org");
  return url.origin;
}

export function salesforceVendor(unguarded: Fetch = fetch): Vendor {
  const fetchImpl = vendorFetch("Salesforce", unguarded);
  const tokens = new TokenCache();

  async function mint(myDomain: string, clientId: string, clientSecret: string) {
    const res = await fetchImpl(`https://${myDomain}/services/oauth2/token`, {
      method: "POST",
      headers: { "content-type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type: "client_credentials",
        client_id: clientId,
        client_secret: clientSecret,
      }),
    });
    const body = (await res.json().catch(() => ({}))) as {
      access_token?: string;
      instance_url?: string;
      id?: string;
      error?: string;
      error_description?: string;
    };
    if (!res.ok || !body.access_token) {
      const why = (body.error_description ?? body.error ?? String(res.status)).split(/\r?\n/)[0];
      throw new ConnectError(`Salesforce refused the sign-in: ${why.slice(0, 300)}`);
    }
    return body as typeof body & { access_token: string };
  }

  const cacheKey = (c: VendorConnection) =>
    createHash("sha256")
      .update(`${c.config.myDomain}\n${c.secret.client_id}\n${c.secret.client_secret}`)
      .digest("hex");

  const token = (c: VendorConnection) =>
    tokens.get(cacheKey(c), async () => ({
      token: (await mint(c.config.myDomain, c.secret.client_id, c.secret.client_secret))
        .access_token,
      expiresIn: TOKEN_SECONDS,
    }));

  async function send(base: string, access: string, request: VendorRequest) {
    const res = await fetchImpl(vendorUrl(base, request.path), {
      method: request.method,
      headers: {
        authorization: `Bearer ${access}`,
        accept: "application/json",
        ...request.headers,
        ...(request.body === undefined ? {} : { "content-type": "application/json" }),
      },
      body: request.body === undefined ? undefined : JSON.stringify(request.body),
    });
    const raw = await res.text();
    return {
      status: res.status,
      data: raw ? safeJson(raw) : null,
      headers: Object.fromEntries(res.headers),
    };
  }

  return {
    async connect(fields) {
      const myDomain = normalizeMyDomain(fields.my_domain ?? "");
      const clientId = fields.client_id?.trim();
      const clientSecret = fields.client_secret?.trim();
      if (!clientId || !clientSecret)
        throw new ConnectError("Enter the consumer key and consumer secret");
      const minted = await mint(myDomain, clientId, clientSecret);
      // The identity URL ends in the Run As user's id; verify reads that user back.
      const userId = minted.id?.split("/").pop() ?? "";
      if (!USER_ID.test(userId))
        throw new ConnectError("Salesforce did not say which user the app runs as");
      return {
        config: { myDomain, instanceUrl: instanceUrl(minted.instance_url), userId },
        secret: { client_id: clientId, client_secret: clientSecret },
        accountLabel: myDomain,
      };
    },

    async call(connection: VendorConnection, request: VendorRequest) {
      const base = instanceUrl(connection.config.instanceUrl);
      let res = await send(base, await token(connection), request);
      if (res.status === 401) {
        tokens.forget(cacheKey(connection));
        res = await send(base, await token(connection), request);
      }
      return res;
    },
  };
}

function safeJson(raw: string): unknown {
  try {
    return JSON.parse(raw);
  } catch {
    return raw;
  }
}
