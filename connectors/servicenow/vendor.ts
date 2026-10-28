import { clientCredentials } from "../../packages/connector-kit/src/oauth.js";
import {
  ConnectError,
  type Fetch,
  type Vendor,
  type VendorConnection,
  type VendorRequest,
} from "../../packages/connector-kit/src/types.js";

// A ServiceNow integration user, signing in through an OAuth application registry entry
// (client credentials, token at /oauth_token.do, run as the entry's OAuth Application User) or
// with its user name and password (basic). Every call goes to the instance's own host.

const INSTANCE = /^[a-z0-9][a-z0-9-]{0,62}\.service-now\.com$/;

export function normalizeInstance(input: string): string {
  let host = input
    .trim()
    .toLowerCase()
    .replace(/^https:\/\//, "")
    .replace(/\/.*$/, "");
  if (/^[a-z0-9][a-z0-9-]{0,62}$/.test(host)) host = `${host}.service-now.com`;
  if (!INSTANCE.test(host))
    throw new ConnectError(
      "Enter your instance address, for example https://your-company.service-now.com",
    );
  return host;
}

function base(config: Record<string, string>): string {
  if (!INSTANCE.test(config.instance ?? ""))
    throw new Error("ServiceNow connection has no instance");
  return `https://${config.instance}`;
}

export function servicenowVendor(fetchImpl: Fetch = fetch): Vendor {
  const oauth = clientCredentials("ServiceNow", fetchImpl);
  const creds = (c: VendorConnection) => ({
    tokenUrl: `${base(c.config)}/oauth_token.do`,
    clientId: c.secret.client_id,
    clientSecret: c.secret.client_secret,
  });

  async function basic(c: VendorConnection, request: VendorRequest) {
    const auth = Buffer.from(`${c.secret.username}:${c.secret.password}`).toString("base64");
    const res = await fetchImpl(`${base(c.config)}${request.path}`, {
      method: request.method,
      headers: {
        authorization: `Basic ${auth}`,
        accept: "application/json",
        ...(request.body === undefined ? {} : { "content-type": "application/json" }),
      },
      body: request.body === undefined ? undefined : JSON.stringify(request.body),
    });
    const raw = await res.text();
    return { status: res.status, data: raw ? safeJson(raw) : null };
  }

  const call: Vendor["call"] = (c, request) =>
    c.config.auth === "basic" ? basic(c, request) : oauth.call(creds(c), base(c.config), request);

  return {
    async connect(fields) {
      const instance = normalizeInstance(fields.instance ?? "");
      const auth = fields.auth === "basic" ? "basic" : "oauth";
      const [idField, secretField] =
        auth === "basic" ? ["username", "password"] : ["client_id", "client_secret"];
      const id = fields[idField]?.trim();
      const secret = fields[secretField]?.trim();
      if (!id || !secret)
        throw new ConnectError(
          auth === "basic"
            ? "Enter the integration user's user name and password"
            : "Enter the client ID and client secret of the application registry entry",
        );
      const conn: VendorConnection = {
        config: { instance, auth },
        secret: { [idField]: id, [secretField]: secret },
      };
      if (auth === "oauth") await oauth.signIn(creds(conn));
      else {
        const res = await basic(conn, {
          method: "GET",
          path: "/api/now/table/sys_user?sysparm_limit=1&sysparm_fields=sys_id",
        });
        if (res.status === 401)
          throw new ConnectError("ServiceNow did not accept the user name and password");
        if (res.status < 200 || res.status >= 300)
          throw new ConnectError(`ServiceNow answered ${res.status} to the sign-in check`);
      }
      return { ...conn, accountLabel: instance };
    },

    call,
  };
}

function safeJson(raw: string): unknown {
  try {
    return JSON.parse(raw);
  } catch {
    return raw;
  }
}
