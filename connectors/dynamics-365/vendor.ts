import type { MicrosoftClient } from "../microsoft/client.js";
import {
  ConnectError,
  type CheckInput,
  type Vendor,
  type VendorConnection,
  type VendorRequest,
} from "../../packages/connector-kit/src/types.js";

// Dynamics 365 (Dataverse) through the Microsoft connection: the same app registration signs in
// with client credentials for the environment itself (`https://<org>.crm.dynamics.com/.default`)
// and acts as the environment's application user, so what it reaches is that user's security
// role. No Entra API permission is involved. Every call goes to the one environment connected.
// Sources: learn.microsoft.com/power-apps/developer/data-platform/
// use-single-tenant-server-server-authentication (read October 27, 2026); the Web API's WhoAmI
// and RetrieveUserPrivileges from Microsoft's Dataverse reference.

export const API = "/api/data/v9.2";

// The Dataverse regions all sign in through Microsoft's global cloud: crm, crm2 ... crm21
// (crm9 is US Government GCC).
const GLOBAL_HOST = /^[a-z0-9][a-z0-9-]{0,62}\.crm\d{0,2}\.dynamics\.com$/;
// GCC High, DoD and China environments, whose tenants sign in elsewhere.
const SOVEREIGN_HOST =
  /^[a-z0-9][a-z0-9-]{0,62}\.crm\d{0,2}\.(microsoftdynamics\.us|appsplatform\.us|dynamics\.cn)$/;

/** An environment address as admins paste it (the app's address, or the Web API's), reduced to
 *  the environment's origin. */
export function normalizeEnvironment(input: string): string {
  const raw = input.trim();
  const url = URL.parse(/^[a-z]+:\/\//i.test(raw) ? raw : `https://${raw}`);
  const host = url?.hostname.toLowerCase().replace(/\.api\.(crm\d{0,2}\.)/, ".$1") ?? "";
  if (url && url.protocol === "https:" && !url.port && SOVEREIGN_HOST.test(host))
    throw new ConnectError(
      "US Government (GCC High, DoD) and China environments are not supported: the Microsoft connection signs in to Microsoft's global cloud.",
    );
  if (!url || url.protocol !== "https:" || url.port || !GLOBAL_HOST.test(host))
    throw new ConnectError(
      `${raw || "The address"} is not a Dynamics 365 environment URL, such as https://your-org.crm.dynamics.com`,
    );
  return `https://${host}`;
}

/** Dataverse's error text, `{error: {code, message}}`, as one line. */
export function dataverseError(status: number, data: unknown): string {
  const message = (data as { error?: { message?: unknown } } | null)?.error?.message;
  const detail =
    typeof message === "string" ? message.split(/\r?\n/)[0].slice(0, 300).replace(/\.$/, "") : "";
  return `Dynamics 365 answered ${status}${detail ? `: ${detail}` : ""}`;
}

type WhoAmI = { UserId?: string; BusinessUnitId?: string };
type Privilege = { PrivilegeName?: string };

export function dynamicsVendor(ms: MicrosoftClient): Vendor {
  const call = async (c: VendorConnection, request: VendorRequest) =>
    ms.call(
      c,
      {
        ...request,
        headers: {
          "OData-MaxVersion": "4.0",
          "OData-Version": "4.0",
          ...request.headers,
        },
      },
      normalizeEnvironment(c.config.environment ?? ""),
    );

  async function whoAmI(c: VendorConnection) {
    const res = await call(c, { method: "GET", path: `${API}/WhoAmI` });
    const who = res.data as WhoAmI | null;
    if (res.status === 200 && who?.UserId) return who;
    const host = new URL(c.config.environment).hostname;
    throw new ConnectError(
      res.status === 401 || res.status === 403
        ? `${dataverseError(res.status, res.data)}. Add the app to ${host} as an application user with the agents' security role (the steps above), then try again. A new application user can take a few minutes to apply.`
        : `${dataverseError(res.status, res.data)} when the app signed in to ${host}.`,
    );
  }

  return {
    async connect(fields, provider) {
      if (!provider) throw new Error("Dynamics 365 connects through the Microsoft connection");
      const environment = normalizeEnvironment(fields.environment ?? "");
      await whoAmI({
        config: { ...provider.config, environment },
        secret: provider.secret,
      });
      return {
        config: { environment },
        secret: {},
        accountLabel: new URL(environment).hostname,
      };
    },

    call,

    checks: {
      // The application user's security roles, and the privileges the ticked capabilities need
      // that its roles do not give. Extra privileges are the admin's choice.
      async access(c, input: CheckInput) {
        const who = await whoAmI(c);
        const user = `${API}/systemusers(${who.UserId})`;
        const [roles, privileges] = await Promise.all([
          call(c, {
            method: "GET",
            path: `${user}?$select=fullname&$expand=systemuserroles_association($select=name)`,
          }),
          call(c, {
            method: "GET",
            path: `${user}/Microsoft.Dynamics.CRM.RetrieveUserPrivileges()`,
          }),
        ]);
        if (roles.status !== 200 || privileges.status !== 200) {
          const failed = roles.status !== 200 ? roles : privileges;
          return {
            values: [],
            warning: `${dataverseError(failed.status, failed.data)} when reading the application user's security roles. Give its role prvReadUser (Read on User), then check again.`,
          };
        }
        const names = (
          (roles.data as { systemuserroles_association?: { name?: string }[] })
            .systemuserroles_association ?? []
        )
          .map((r) => r.name)
          .filter((n): n is string => typeof n === "string")
          .sort();
        const held = new Set(
          ((privileges.data as { RolePrivileges?: Privilege[] }).RolePrivileges ?? []).map(
            (p) => p.PrivilegeName,
          ),
        );
        const missing = input.expected.filter((p) => !held.has(p));
        return {
          values: names,
          ...(missing.length
            ? {
                warning: `Its security roles lack ${missing.join(", ")}. Add them to the role, or agents cannot do what you ticked.`,
              }
            : {}),
        };
      },
    },
  };
}
