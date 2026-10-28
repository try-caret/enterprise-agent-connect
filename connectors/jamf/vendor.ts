import { clientCredentials } from "../../packages/connector-kit/src/oauth.js";
import {
  ConnectError,
  type Fetch,
  type Vendor,
  type VendorConnection,
} from "../../packages/connector-kit/src/types.js";

// A Jamf Pro API client (Settings > System > API roles and clients): OAuth client credentials at
// /api/v1/oauth/token (form-encoded, expires_in is the client's access token lifetime), then the
// Jamf Pro API (/api/...) and the Classic API (/JSSResource/...) on the same server with the
// bearer token. Source: developer.jamf.com/jamf-pro/reference/postoauthtoken and the Client
// Credentials guide (read October 27, 2026).

export const TOKEN_PATH = "/api/v1/oauth/token";

/** The server's origin: https only, no path, query or credentials. Cloud or on premises. */
export function normalizeUrl(raw: string): string {
  const text = raw.trim();
  const url = URL.parse(/^[a-z]+:\/\//i.test(text) ? text : `https://${text}`);
  if (
    !url ||
    url.protocol !== "https:" ||
    url.username ||
    url.password ||
    url.search ||
    url.hash ||
    !/^\/*$/.test(url.pathname)
  )
    throw new ConnectError(
      "Enter your Jamf Pro address with https and nothing after the host, such as https://your-company.jamfcloud.com",
    );
  return url.origin;
}

function origin(config: Record<string, string>): string {
  const url = config.url ?? "";
  if (URL.parse(url)?.origin !== url || !url.startsWith("https://"))
    throw new Error("Jamf Pro connection has no server address");
  return url;
}

export function jamfVendor(fetchImpl: Fetch = fetch): Vendor {
  const oauth = clientCredentials("Jamf Pro", fetchImpl);
  const creds = (c: VendorConnection) => ({
    tokenUrl: `${origin(c.config)}${TOKEN_PATH}`,
    clientId: c.secret.client_id,
    clientSecret: c.secret.client_secret,
  });
  const call: Vendor["call"] = async (c, request) =>
    oauth.call(creds(c), origin(c.config), request);

  return {
    async connect(fields) {
      const url = normalizeUrl(fields.url ?? "");
      const clientId = fields.client_id?.trim();
      const clientSecret = fields.client_secret?.trim();
      if (!clientId || !clientSecret) throw new ConnectError("Enter the client ID and secret");
      const host = new URL(url).host;
      const conn: VendorConnection = {
        config: { url, host },
        secret: { client_id: clientId, client_secret: clientSecret },
      };
      await oauth.signIn(creds(conn));
      return { ...conn, accountLabel: host };
    },

    call,

    checks: {
      // The site ids a key's sites limit chooses from.
      async sites(connection) {
        const res = await call(connection, { method: "GET", path: "/api/v1/sites" });
        if (res.status === 401 || res.status === 403)
          throw new ConnectError(
            "Jamf Pro refused reading sites. Add Read Sites to the API role, then check again.",
          );
        if (res.status !== 200 || !Array.isArray(res.data))
          throw new ConnectError(`Jamf Pro answered ${res.status} for its sites`);
        return { values: (res.data as { id: string }[]).map((s) => String(s.id)) };
      },
    },
  };
}
