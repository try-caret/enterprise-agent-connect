import { createHash } from "node:crypto";
import { vendorFetch } from "../../packages/connector-kit/src/fetch.js";
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
// Jira Assets: tools send paths under ASSETS_PATH with the literal {workspaceId}, filled here.
// The site's Assets workspace id comes from GET /rest/servicedeskapi/assets/workspace (Jira
// Service Management's "Get assets workspaces", scope read:servicedesk-request) and is kept per
// cloudId: a site has one workspace and its id does not change. Assets is reached through the
// same gateway as Jira (api.atlassian.com/ex/jira/{cloudId}/jsm/assets/workspace/{id}/v1,
// the server of Atlassian's Assets OpenAPI, developer.atlassian.com/cloud/assets/swagger.v3.json).
export const ASSETS_PATH = "/jsm/assets/workspace/{workspaceId}/v1/";
// Atlassian's own example ids are not hex ("1060ba0e-178b-4e0e-g0h1-jedb02cccb5f"), so only the
// characters, never the UUID shape, are checked: enough to keep it one path segment.
const WORKSPACE_ID = /^[A-Za-z0-9-]{1,64}$/;
// Every answer about the workspace itself says this, so a tool can tell it from an object's 404.
export const WORKSPACE_PROBLEM = "Assets workspace";

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

export function jiraVendor(unguarded: Fetch = fetch): Vendor {
  const fetchImpl = vendorFetch("Jira", unguarded);
  const tokens = new TokenCache();
  const workspaces = new Map<string, string>();

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
      const send = async (method: VendorRequest["method"], path: string, body?: unknown) => {
        const res = await fetchImpl(`${API}/ex/jira/${cloudId}${path}`, {
          method,
          headers: {
            authorization: `Bearer ${access}`,
            ...JSON_EN,
            ...(body === undefined ? {} : { "content-type": "application/json" }),
          },
          body: body === undefined ? undefined : JSON.stringify(body),
        });
        const raw = await res.text();
        return {
          status: res.status,
          data: raw ? safeJson(raw) : null,
          headers: Object.fromEntries(res.headers),
        };
      };
      if (!request.path.startsWith(ASSETS_PATH))
        return send(request.method, request.path, request.body);
      // Per site and credential: another organization's account on the same site may lack the
      // seat, and must see the lookup's own refusal.
      const cacheKey = `${cloudId}:${connection.secret.client_id}`;
      let workspace = workspaces.get(cacheKey);
      if (!workspace) {
        const found = await send("GET", "/rest/servicedeskapi/assets/workspace");
        if (found.status === 401 || found.status === 403)
          return {
            ...found,
            data: {
              errorMessages: [
                `Jira Service Management refused the ${WORKSPACE_PROBLEM} lookup: the service ` +
                  "account needs Jira Service Management access as User (an agent seat) and " +
                  "the read:servicedesk-request scope",
              ],
            },
          };
        if (found.status < 200 || found.status >= 300)
          return {
            ...found,
            data: { errorMessages: [`The ${WORKSPACE_PROBLEM} lookup failed`] },
          };
        const id = (found.data as { values?: { workspaceId?: unknown }[] } | null)?.values?.[0]
          ?.workspaceId;
        if (typeof id !== "string" || !WORKSPACE_ID.test(id))
          return {
            status: 404,
            data: { errorMessages: [`This Jira site has no ${WORKSPACE_PROBLEM}`] },
            headers: {},
          };
        workspaces.set(cacheKey, id);
        workspace = id;
      }
      return send(request.method, request.path.replace("{workspaceId}", workspace), request.body);
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
