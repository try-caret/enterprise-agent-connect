import { createHash } from "node:crypto";
import {
  ConnectError,
  TokenCache,
  type Fetch,
  type Vendor,
  type VendorConnection,
  type VendorRequest,
} from "../../packages/connector-kit/src/types.js";

// Atlassian service account, OAuth 2.0 client credentials. Tokens last 60 minutes and
// reach Jira through api.atlassian.com/ex/jira/{cloudId}.

const AUTH_URL = "https://auth.atlassian.com/oauth/token";
const API = "https://api.atlassian.com";
// Jira writes errors in the account's language otherwise; agents and admins read English.
const JSON_EN = { accept: "application/json", "accept-language": "en" };
const CLOUD_ID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/;

export function normalizeSite(input: string): string {
  const host = input
    .trim()
    .toLowerCase()
    .replace(/^https?:\/\//, "")
    .replace(/\/.*$/, "")
    .replace(/\.atlassian\.net$/, "");
  if (!/^[a-z0-9][a-z0-9-]{0,62}$/.test(host))
    throw new ConnectError("Enter the Jira site name, the part before .atlassian.net");
  return host;
}

export function jiraVendor(fetchImpl: Fetch = fetch): Vendor {
  const tokens = new TokenCache();

  const token = (clientId: string, clientSecret: string) =>
    tokens.get(
      createHash("sha256").update(`${clientId}\n${clientSecret}`).digest("hex"),
      async () => {
        const res = await fetchImpl(AUTH_URL, {
          method: "POST",
          headers: { "content-type": "application/json" },
          body: JSON.stringify({
            grant_type: "client_credentials",
            client_id: clientId,
            client_secret: clientSecret,
          }),
        });
        if (!res.ok)
          throw new ConnectError(
            res.status === 401 || res.status === 403
              ? "Atlassian did not accept the client ID and secret"
              : `Atlassian sign-in failed (${res.status})`,
          );
        const body = (await res.json()) as { access_token: string; expires_in: number };
        return { token: body.access_token, expiresIn: body.expires_in };
      },
    );

  return {
    async connect(fields) {
      const site = normalizeSite(fields.site ?? "");
      const clientId = fields.client_id?.trim();
      const clientSecret = fields.client_secret?.trim();
      if (!clientId || !clientSecret) throw new ConnectError("Enter the client ID and secret");
      const access = await token(clientId, clientSecret);
      // accessible-resources is empty for service-account tokens; the site publishes its
      // cloudId, and serverInfo through the gateway proves the account can reach it.
      const tenant = await fetchImpl(`https://${site}.atlassian.net/_edge/tenant_info`);
      if (!tenant.ok) throw new ConnectError(`${site}.atlassian.net is not a Jira site`);
      const { cloudId } = (await tenant.json()) as { cloudId?: string };
      if (!cloudId || !CLOUD_ID.test(cloudId))
        throw new ConnectError(`${site}.atlassian.net did not identify itself`);
      const info = await fetchImpl(`${API}/ex/jira/${cloudId}/rest/api/3/serverInfo`, {
        headers: { authorization: `Bearer ${access}`, ...JSON_EN },
      });
      if (!info.ok)
        throw new ConnectError(
          `The service account cannot reach ${site}.atlassian.net (${info.status}). Check that it belongs to that site and has the scopes listed above.`,
        );
      return {
        config: { site, cloudId },
        secret: { client_id: clientId, client_secret: clientSecret },
        accountLabel: `${site}.atlassian.net`,
      };
    },

    async call(connection: VendorConnection, request: VendorRequest) {
      const { cloudId } = connection.config;
      if (!CLOUD_ID.test(cloudId ?? "")) throw new Error("Jira connection has no cloudId");
      const access = await token(connection.secret.client_id, connection.secret.client_secret);
      const res = await fetchImpl(`${API}/ex/jira/${cloudId}${request.path}`, {
        method: request.method,
        headers: {
          authorization: `Bearer ${access}`,
          ...JSON_EN,
          ...(request.body === undefined ? {} : { "content-type": "application/json" }),
        },
        body: request.body === undefined ? undefined : JSON.stringify(request.body),
      });
      const raw = await res.text();
      return { status: res.status, data: raw ? safeJson(raw) : null };
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
