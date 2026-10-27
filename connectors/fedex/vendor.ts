import { createHash } from "node:crypto";
import { clientCredentials, vendorError } from "../../packages/connector-kit/src/oauth.js";
import {
  ConnectError,
  type Fetch,
  type Vendor,
  type VendorConnection,
} from "../../packages/connector-kit/src/types.js";

// A FedEx Developer Portal project: API key and secret key, OAuth client credentials at
// /oauth/token (form-encoded, expires_in 3600), then JSON calls with the bearer token. Hosts
// and request shapes are FedEx's published OpenAPI files (developer.fedex.com/api/en-us/catalog/
// {authorization,track,rate,ship}/v1/docs.html, retrieved October 21, 2026).

export const FEDEX_HOSTS: Record<string, string> = {
  sandbox: "https://apis-sandbox.fedex.com",
  production: "https://apis.fedex.com",
};
export const ACCOUNT_NUMBER = /^[0-9]{9,12}$/;
const MAX_ACCOUNTS = 50;
// FedEx documents x-locale; errors and status text come back in English.
const HEADERS = { "x-locale": "en_US" };
// A number no shipment has: the Track API answers it (with a not-found result), which proves
// the token reaches the Track API without reading anyone's shipment.
export const PROBE_TRACKING_NUMBER = "000000000000";

export function parseAccounts(raw: string): string[] {
  const accounts = [...new Set(raw.split(/[\s,;]+/).filter(Boolean))];
  if (!accounts.length) throw new ConnectError("Enter at least one FedEx account number");
  const bad = accounts.find((a) => !ACCOUNT_NUMBER.test(a));
  if (bad) throw new ConnectError(`${bad} is not a FedEx account number (9 to 12 digits)`);
  if (accounts.length > MAX_ACCOUNTS)
    throw new ConnectError(`Enter at most ${MAX_ACCOUNTS} account numbers`);
  return accounts;
}

export function connectedAccounts(config: Record<string, string>): string[] {
  return (config.accounts ?? "").split(",").filter(Boolean);
}

function host(config: Record<string, string>): string {
  const base = Object.hasOwn(FEDEX_HOSTS, config.environment)
    ? FEDEX_HOSTS[config.environment]
    : null;
  if (!base) throw new Error(`Unknown FedEx environment ${config.environment}`);
  return base;
}

export function fedexVendor(fetchImpl: Fetch = fetch): Vendor {
  const oauth = clientCredentials("FedEx", fetchImpl);
  const creds = (c: VendorConnection) => ({
    tokenUrl: `${host(c.config)}/oauth/token`,
    clientId: c.secret.api_key,
    clientSecret: c.secret.secret_key,
  });
  const call: Vendor["call"] = (c, request) =>
    oauth.call(creds(c), host(c.config), request, HEADERS);

  return {
    async connect(fields) {
      const environment = fields.environment ?? "";
      if (!Object.hasOwn(FEDEX_HOSTS, environment))
        throw new ConnectError("Choose Sandbox or Production");
      const apiKey = fields.api_key?.trim();
      const secretKey = fields.secret_key?.trim();
      if (!apiKey || !secretKey)
        throw new ConnectError("Enter the API key and secret key from your FedEx project");
      const accounts = parseAccounts(fields.account_numbers ?? "");
      const conn: VendorConnection = {
        config: {
          environment,
          accounts: accounts.join(","),
          // Names the project reached without keeping the key in plain config.
          project: createHash("sha256").update(apiKey).digest("hex").slice(0, 12),
        },
        secret: { api_key: apiKey, secret_key: secretKey },
      };
      await oauth.signIn(creds(conn));
      return {
        ...conn,
        accountLabel: `${environment === "production" ? "Production" : "Sandbox"} · ${accounts.length} account${accounts.length === 1 ? "" : "s"}`,
      };
    },

    call,

    checks: {
      // A Track API answer about a number no shipment has.
      async track(connection) {
        const res = await call(connection, {
          method: "POST",
          path: "/track/v1/trackingnumbers",
          body: {
            includeDetailedScans: false,
            trackingInfo: [{ trackingNumberInfo: { trackingNumber: PROBE_TRACKING_NUMBER } }],
          },
        });
        if (res.status === 401 || res.status === 403)
          throw new ConnectError(
            `FedEx refused the Track API for this project. Add the Track API to the project, then check again. (${vendorError(res.data) || res.status})`,
          );
        if (res.status < 200 || res.status >= 300)
          throw new ConnectError(
            `FedEx Track API answered ${res.status}: ${vendorError(res.data) || "no reason given"}`,
          );
        return { values: [new URL(host(connection.config)).host] };
      },
      async accounts(connection) {
        return { values: connectedAccounts(connection.config) };
      },
    },
  };
}
