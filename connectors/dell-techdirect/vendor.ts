import { createHash } from "node:crypto";
import { clientCredentials, vendorError } from "../../packages/connector-kit/src/oauth.js";
import {
  ConnectError,
  type Fetch,
  type Vendor,
  type VendorConnection,
} from "../../packages/connector-kit/src/types.js";

// Dell TechDirect Warranty API (v5): a client ID and secret Dell issues once it approves the
// company's API request, OAuth client credentials at apigtwb2c.us.dell.com/auth/oauth/v2/token
// (form-encoded, tokens last 3600 seconds), then asset-entitlements by service tag. Dell's
// reference sits behind a TechDirect sign-in; these addresses are the ones Dell's community
// threads and integrators document (see the connector's notes). Not yet tried with real keys.

export const DELL_API = "https://apigtwb2c.us.dell.com";
export const DELL_TOKEN_URL = `${DELL_API}/auth/oauth/v2/token`;
export const ENTITLEMENTS_PATH = "/PROD/sbil/eapi/v5/asset-entitlements";
// A tag no machine has: Dell answers it as invalid, which proves the Warranty API is granted
// without looking up anyone's asset.
export const PROBE_SERVICE_TAG = "0000000";

const MISSING_KEYS =
  "Enter the client ID and client secret from Dell TechDirect (APIs, then API Key Management). Dell issues them only after it approves your Warranty API request, so this app cannot connect until then.";

export function dellVendor(fetchImpl: Fetch = fetch): Vendor {
  const oauth = clientCredentials("Dell TechDirect", fetchImpl);
  const creds = (c: VendorConnection) => {
    if (!c.secret.client_id || !c.secret.client_secret) throw new ConnectError(MISSING_KEYS);
    return {
      tokenUrl: DELL_TOKEN_URL,
      clientId: c.secret.client_id,
      clientSecret: c.secret.client_secret,
    };
  };
  const call: Vendor["call"] = (c, request) => oauth.call(creds(c), DELL_API, request);

  return {
    async connect(fields) {
      const clientId = fields.client_id?.trim() ?? "";
      const clientSecret = fields.client_secret?.trim() ?? "";
      const conn: VendorConnection = {
        config: {
          keyId: createHash("sha256").update(clientId).digest("hex").slice(0, 12),
        },
        secret: { client_id: clientId, client_secret: clientSecret },
      };
      await oauth.signIn(creds(conn));
      return { ...conn, accountLabel: `Client ID ending ${clientId.slice(-4)}` };
    },

    call,

    checks: {
      async warranty(connection) {
        const res = await call(connection, {
          method: "GET",
          path: `${ENTITLEMENTS_PATH}?servicetags=${PROBE_SERVICE_TAG}`,
        });
        if (res.status === 401 || res.status === 403)
          throw new ConnectError(
            `Dell TechDirect refused the Warranty API for these keys. Check in TechDirect that the key is approved for the Warranty API. (${vendorError(res.data) || res.status})`,
          );
        if (res.status < 200 || res.status >= 300 || !Array.isArray(res.data))
          throw new ConnectError(
            `Dell TechDirect Warranty API answered ${res.status}: ${vendorError(res.data) || "an unexpected answer"}`,
          );
        return { values: ["Warranty lookup"] };
      },
    },
  };
}
