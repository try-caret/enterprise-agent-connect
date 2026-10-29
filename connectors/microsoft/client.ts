import { createHash } from "node:crypto";
import { brand } from "../../packages/connector-kit/src/branding.js";
import { clientAssertion } from "../../packages/connector-kit/src/certificate.js";
import type { Toolset } from "../../packages/connector-kit/src/common.js";
import {
  choiceFields,
  type ConnectorContext,
  type ConnectorModule,
} from "../../packages/connector-kit/src/connector.js";
import { vendorFetch, vendorUrl } from "../../packages/connector-kit/src/fetch.js";
import {
  ConnectError,
  TokenCache,
  type CheckInput,
  type Fetch,
  type ProviderClient,
  type Vendor,
  type VendorConnection,
  type VendorRequest,
} from "../../packages/connector-kit/src/types.js";

// One Microsoft connection per organization, shared by the Microsoft apps (mail, files).
// Normally the customer's own single-tenant Entra app, signing in with a certificate the
// service made (client assertion) or a client secret. Where the deployment offers it, instead
// the deployment's shared multi-tenant app, approved by the customer's administrator.
// config.auth: "certificate" | "client_secret" | SHARED_APP_AUTH.

/** config.auth of a connection approved for the shared app. A value stored in existing rows
 *  (it predates the service's own name), so it never changes. */
export const SHARED_APP_AUTH = "arbium_app";

// Entra role template ids (the same in every tenant) of the roles that can approve an app:
// Global Administrator, Privileged Role Administrator, Cloud Application Administrator,
// Application Administrator.
const APPROVER_ROLES = new Set([
  "62e90394-69f5-4237-9190-012177145e10",
  "e8611ab8-c189-46e8-94e1-60213ab1f814",
  "158c047a-c907-4556-b7ef-446551a6b5f7",
  "9b895d92-2cd3-44c7-9d02-a6ac2d5ea5c3",
]);

const LOGIN = "https://login.microsoftonline.com";
const GRAPH = "https://graph.microsoft.com";
// Office 365 Exchange Online: permissions granted on it in Entra (full_access_as_app, its
// Mail.*) open mailboxes through EWS and never show in a Graph token.
const EXCHANGE = "https://outlook.office365.com";
export const EXCHANGE_SUFFIX = " (Exchange Online)";
export const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/;
// Right after approval Microsoft can take a few seconds to make the app usable in the tenant.
const CONNECT_RETRY_MS = [3000, 6000, 10000];

export type SharedApp = { clientId: string; clientSecret: string };

const tokenUrl = (tenant: string) => `${LOGIN}/${tenant}/oauth2/v2.0/token`;

// Claims of a token Microsoft just handed us over TLS; no signature check needed.
function tokenClaims(jwt: string): Record<string, unknown> {
  try {
    return JSON.parse(Buffer.from(jwt.split(".")[1], "base64url").toString("utf8"));
  } catch {
    throw new ConnectError("Microsoft returned an unreadable token");
  }
}

export function tokenClaim(jwt: string, claim: string): string | null {
  try {
    const value = tokenClaims(jwt)[claim];
    return typeof value === "string" ? value : null;
  } catch {
    return null;
  }
}

/** The application permissions in an access token (its `roles` claim). */
export function tokenRoles(jwt: string): string[] {
  const roles = tokenClaims(jwt).roles;
  return (Array.isArray(roles) ? roles : [])
    .filter((r): r is string => typeof r === "string")
    .sort();
}

/** The permissions the ticked capabilities need that the app lacks. Extra ones are the
 *  admin's choice, so they are not called out. */
export function compareRoles(granted: string[], input: CheckInput): string | undefined {
  const missing = input.expected.filter((r) => !granted.includes(r));
  if (!missing.length) return undefined;
  const them = missing.length > 1 ? "them" : "it";
  return `Missing: ${missing.join(", ")}. Add ${them} to the app and grant admin consent, or agents cannot do what you ticked.`;
}

function safeJson(raw: string): unknown {
  try {
    return JSON.parse(raw);
  } catch {
    return raw;
  }
}

type TokenAnswer = {
  access_token?: string;
  expires_in?: number;
  error?: string;
  error_description?: string;
  id_token?: string;
};

export function microsoftClient(
  options: { sharedApp: SharedApp | null },
  unguarded: Fetch = fetch,
  sleep = (ms: number) => new Promise((r) => setTimeout(r, ms)),
) {
  const fetchImpl = vendorFetch("Microsoft", unguarded, (s) => sleep(s * 1000));
  const { sharedApp } = options;
  const tokens = new TokenCache();

  async function requestToken(tenant: string, form: Record<string, string>) {
    const res = await fetchImpl(tokenUrl(tenant), {
      method: "POST",
      headers: { "content-type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams(form),
    });
    const body = (await res.json().catch(() => ({}))) as TokenAnswer;
    return { ok: res.ok, status: res.status, body };
  }

  // The token request's form for this connection's way of signing in.
  function credentials(c: VendorConnection): Record<string, string> {
    const { auth, tenantId, clientId, certificate } = c.config;
    if (auth === SHARED_APP_AUTH) {
      if (!sharedApp)
        throw new ConnectError(
          `This Microsoft connection was approved for ${brand().productName}'s shared app, which this deployment does not offer. Disconnect Microsoft and connect it with your own app registration.`,
        );
      return { client_id: sharedApp.clientId, client_secret: sharedApp.clientSecret };
    }
    if (!GUID.test(tenantId ?? "") || !GUID.test(clientId ?? ""))
      throw new Error("Microsoft connection has no tenantId or clientId");
    if (auth === "client_secret") {
      if (!c.secret.client_secret) throw new Error("Microsoft connection has no client secret");
      return { client_id: clientId, client_secret: c.secret.client_secret };
    }
    if (auth === "certificate") {
      if (!certificate || !c.secret.private_key)
        throw new Error("Microsoft connection has no certificate");
      return {
        client_id: clientId,
        client_assertion_type: "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
        client_assertion: clientAssertion({
          clientId,
          tokenUrl: tokenUrl(tenantId),
          certificate,
          privateKey: c.secret.private_key,
        }),
      };
    }
    throw new Error(`Microsoft connection has unknown sign-in ${auth}`);
  }

  // `resource`: the API the token is for, Microsoft Graph unless an app names another (e.g.
  // Exchange Online).
  async function mint(c: VendorConnection, resource = GRAPH) {
    const tenant = c.config.tenantId ?? "";
    if (!GUID.test(tenant)) throw new Error("Microsoft connection has no tenantId");
    const { ok, status, body } = await requestToken(tenant, {
      grant_type: "client_credentials",
      scope: `${resource}/.default`,
      ...credentials(c),
    });
    if (!ok || !body.access_token) {
      const why = (body.error_description ?? body.error ?? String(status)).split(/\r?\n/)[0];
      throw new ConnectError(`Microsoft refused the app's sign-in: ${why.slice(0, 300)}`);
    }
    return { token: body.access_token, expiresIn: body.expires_in ?? 3600 };
  }

  // Cached per credential, so a new certificate or secret never reuses an old token.
  const cacheKey = (c: VendorConnection, resource: string) =>
    createHash("sha256")
      .update(
        [
          resource,
          c.config.auth,
          c.config.tenantId,
          c.config.clientId,
          c.config.certificate,
          c.secret.client_secret,
        ].join("\n"),
      )
      .digest("hex");

  async function send(token: string, request: VendorRequest, origin = GRAPH) {
    // Uploads send the file's bytes; everything else is JSON.
    const raw = request.body instanceof Uint8Array;
    // A file's content answers 302 to a pre-authenticated SharePoint download address, fetched
    // without the token; any other redirect is refused.
    const download = request.responseType === "text" ? unguarded : fetchImpl;
    let res = await download(vendorUrl(origin, request.path), {
      redirect: "manual",
      method: request.method,
      headers: {
        authorization: `Bearer ${token}`,
        accept: "application/json",
        ...request.headers,
        ...(request.body === undefined
          ? {}
          : { "content-type": raw ? "application/octet-stream" : "application/json" }),
      },
      body:
        request.body === undefined
          ? undefined
          : raw
            ? (request.body as Uint8Array<ArrayBuffer>)
            : JSON.stringify(request.body),
    });
    if (download === unguarded && res.status >= 300 && res.status < 400) {
      await res.body?.cancel();
      res = await fetchImpl(sharePointUrl(res.headers.get("location") ?? ""));
    }
    const text = await res.text();
    const asText = request.responseType === "text" && res.ok;
    return {
      status: res.status,
      data: asText ? text : text ? safeJson(text) : null,
      headers: Object.fromEntries(res.headers),
    };
  }

  /** A request to Graph, or to `origin` with a token for it, re-minting the token once on a
   *  401. */
  async function call(connection: VendorConnection, request: VendorRequest, origin = GRAPH) {
    const key = cacheKey(connection, origin);
    const token = () => tokens.get(key, () => mint(connection, origin));
    let res = await send(await token(), request, origin);
    if (res.status === 401) {
      tokens.forget(key);
      res = await send(await token(), request, origin);
    }
    return res;
  }

  /** Compares the app's permissions, read from fresh Graph and Exchange Online tokens, with
   *  what the app's capabilities need. Exchange Online ones are shown with EXCHANGE_SUFFIX.
   *  Microsoft calls its tokens opaque to clients, so an unreadable one only means this
   *  readout is unavailable; the connection itself still works. */
  async function permissions(connection: VendorConnection, input: CheckInput) {
    const [graph, exchange] = await Promise.all([mint(connection), mint(connection, EXCHANGE)]);
    let roles: string[];
    try {
      roles = [
        ...tokenRoles(graph.token),
        ...tokenRoles(exchange.token).map((r) => `${r}${EXCHANGE_SUFFIX}`),
      ];
    } catch (error) {
      if (!(error instanceof ConnectError)) throw error;
      return {
        values: [],
        warning: "Could not read the app's permissions from Microsoft's token.",
      };
    }
    return { values: roles, warning: compareRoles(roles, input) };
  }

  const provider: ProviderClient = {
    async connect(fields, existing) {
      const tenantId = fields.tenant_id?.trim().toLowerCase() ?? "";
      const clientId = fields.client_id?.trim().toLowerCase() ?? "";
      const clientSecret = fields.client_secret?.trim() ?? "";
      if (!GUID.test(tenantId))
        throw new ConnectError("Enter the Directory (tenant) ID from the app's Overview page");
      if (!GUID.test(clientId))
        throw new ConnectError("Enter the Application (client) ID from the app's Overview page");
      let connection: VendorConnection;
      if (clientSecret)
        connection = {
          config: { auth: "client_secret", tenantId, clientId },
          secret: { client_secret: clientSecret },
        };
      else {
        const certificate = existing?.config.certificate;
        const privateKey = existing?.secret.private_key;
        if (!certificate || !privateKey)
          throw new ConnectError(
            "Create the certificate and upload it to the app, or enter a client secret",
          );
        connection = {
          config: { auth: "certificate", tenantId, clientId, certificate },
          secret: { private_key: privateKey },
        };
      }
      const { token } = await mint(connection);
      const servicePrincipalId = tokenClaim(token, "oid");
      if (!servicePrincipalId) throw new Error("Microsoft token has no oid claim");
      if (tokenClaim(token, "tid") !== tenantId)
        throw new ConnectError("Microsoft signed the app in to another organization");
      // A certificate made for rotating stays waiting with the connection.
      const next = existing?.config.nextCertificate;
      return {
        config: {
          ...connection.config,
          servicePrincipalId,
          ...(next ? { nextCertificate: next } : {}),
        },
        secret: {
          ...connection.secret,
          ...(next && existing?.secret.next_private_key
            ? { next_private_key: existing.secret.next_private_key }
            : {}),
        },
        accountLabel: tenantId,
      };
    },

    async verify(connection) {
      await mint(connection);
    },

    ...(sharedApp ? { approval: sharedApproval(sharedApp) } : {}),
  };

  function sharedApproval(app: SharedApp): NonNullable<ProviderClient["approval"]> {
    return {
      clientId: app.clientId,
      consentUrl({ state, redirectUri }) {
        const q = new URLSearchParams({
          client_id: app.clientId,
          scope: `${GRAPH}/.default`,
          redirect_uri: redirectUri,
          state,
        });
        return `${LOGIN}/organizations/v2.0/adminconsent?${q}`;
      },

      signInUrl({ tenant, state, redirectUri, nonce }) {
        if (!GUID.test(tenant))
          throw new ConnectError(
            `Microsoft did not say which organization approved ${brand().productName}`,
          );
        const q = new URLSearchParams({
          client_id: app.clientId,
          response_type: "code",
          response_mode: "query",
          scope: "openid",
          redirect_uri: redirectUri,
          state,
          nonce,
        });
        return `${LOGIN}/${tenant}/oauth2/v2.0/authorize?${q}`;
      },

      async proveTenant({ tenant, code, redirectUri, nonce }) {
        const { ok, status, body } = await requestToken(tenant, {
          grant_type: "authorization_code",
          client_id: app.clientId,
          client_secret: app.clientSecret,
          code,
          redirect_uri: redirectUri,
          scope: "openid",
        });
        if (!ok || !body.id_token)
          throw new ConnectError(`Microsoft did not confirm the sign-in (${body.error ?? status})`);
        // Straight from Microsoft's token endpoint over TLS with the shared app's secret, so the token's
        // signature need not be checked again (OpenID Connect Core 3.1.3.7).
        const claims = tokenClaims(body.id_token);
        const issuer = `${LOGIN}/${tenant}/v2.0`;
        const fresh = typeof claims.exp === "number" && claims.exp * 1000 > Date.now();
        if (
          claims.aud !== app.clientId ||
          claims.iss !== issuer ||
          claims.nonce !== nonce ||
          !fresh
        )
          throw new ConnectError("Microsoft's sign-in did not match this approval. Start again.");
        // A guest signs in to the tenant but belongs to another one (idp names its home).
        const guest = typeof claims.idp === "string" && !claims.idp.includes(tenant);
        if (claims.tid !== tenant || guest)
          throw new ConnectError(
            `Sign in with an account of the organization that approved ${brand().productName}, not a guest or another organization's account.`,
          );
        // The approval's admin_consent=true came through the browser, so the account signing in
        // must itself hold a role that can approve apps. `wids` is in the id token only when
        // the shared app's manifest sets groupMembershipClaims to DirectoryRole (or All).
        const roles = Array.isArray(claims.wids) ? (claims.wids as unknown[]) : null;
        if (!roles)
          throw new ConnectError(
            `Microsoft did not say whether this account is an administrator, so the approval cannot be confirmed. Contact ${brand().productName}.`,
          );
        if (!roles.some((r) => typeof r === "string" && APPROVER_ROLES.has(r.toLowerCase())))
          throw new ConnectError(
            "Sign in with an administrator who can approve apps for your organization (Global Administrator, Privileged Role Administrator, Cloud Application Administrator or Application Administrator).",
          );
      },

      async connect(rawTenant) {
        const tenant = rawTenant.trim().toLowerCase();
        if (!GUID.test(tenant))
          throw new ConnectError(
            `Microsoft did not say which organization approved ${brand().productName}`,
          );
        const connection = {
          config: { auth: SHARED_APP_AUTH, tenantId: tenant, clientId: app.clientId },
          secret: {},
        };
        let failure: ConnectError | null = null;
        for (const wait of [0, ...CONNECT_RETRY_MS]) {
          if (wait) await sleep(wait);
          try {
            // A fresh token each try: one minted before the approval applied lacks its role.
            const { token } = await mint(connection);
            const org = await send(token, {
              method: "GET",
              path: "/v1.0/organization?$select=displayName,verifiedDomains",
            });
            if (org.status !== 200) {
              failure = new ConnectError(
                `Microsoft answered ${org.status} when ${brand().productName} checked the connection`,
              );
              continue;
            }
            const servicePrincipalId = tokenClaim(token, "oid");
            if (!servicePrincipalId) throw new Error("Microsoft token has no oid claim");
            const info = (org.data as { value?: OrgInfo[] }).value?.[0];
            const domain = info?.verifiedDomains?.find((d) => d.isDefault)?.name;
            const name = info?.displayName ?? tenant;
            return {
              config: { ...connection.config, servicePrincipalId, organization: name },
              secret: {},
              accountLabel: domain ? `${name} (${domain})` : name,
            };
          } catch (error) {
            if (!(error instanceof ConnectError)) throw error;
            failure = error;
          }
        }
        throw failure!;
      },
    };
  }

  // Cancels an upload session. Its address is pre-authenticated, so no token is sent, and
  // only a SharePoint address is called.
  async function cancelUpload(uploadUrl: string): Promise<number> {
    return (await fetchImpl(sharePointUrl(uploadUrl), { method: "DELETE" })).status;
  }

  // An administrator signing in to the connection's own app, for what Microsoft allows only on
  // a person's behalf (publishing to the Teams app catalog). The token is used at once and
  // never kept.
  const delegated = {
    signInUrl(
      c: VendorConnection,
      input: { state: string; redirectUri: string; nonce: string; scope: string },
    ) {
      const { tenantId, clientId } = c.config;
      if (!GUID.test(tenantId ?? "") || !GUID.test(clientId ?? ""))
        throw new Error("Microsoft connection has no tenantId or clientId");
      const q = new URLSearchParams({
        client_id: clientId,
        response_type: "code",
        response_mode: "query",
        scope: `openid ${input.scope}`,
        redirect_uri: input.redirectUri,
        state: input.state,
        nonce: input.nonce,
        prompt: "select_account",
      });
      return `${LOGIN}/${tenantId}/oauth2/v2.0/authorize?${q}`;
    },

    /** The code redeemed with the connection's own credential. Returns the access token and
     *  the id token's claims (straight from Microsoft over TLS, so not signature-checked). */
    async redeem(c: VendorConnection, input: { code: string; redirectUri: string; scope: string }) {
      const { ok, status, body } = await requestToken(c.config.tenantId ?? "", {
        grant_type: "authorization_code",
        code: input.code,
        redirect_uri: input.redirectUri,
        scope: `openid ${input.scope}`,
        ...credentials(c),
      });
      if (!ok || !body.access_token || !body.id_token) {
        const why = (body.error_description ?? body.error ?? String(status)).split(/\r?\n/)[0];
        throw new ConnectError(`Microsoft did not confirm the sign-in: ${why.slice(0, 300)}`);
      }
      return { accessToken: body.access_token, claims: tokenClaims(body.id_token) };
    },
  };

  return { provider, call, permissions, cancelUpload, delegated };
}

export type MicrosoftClient = ReturnType<typeof microsoftClient>;

/** What a built-in app's module is given: also the deployment's Microsoft client. */
export type BuiltinContext = ConnectorContext & { microsoft: MicrosoftClient };
export type BuiltinModule = ConnectorModule<BuiltinContext>;

// Pre-authenticated addresses Graph hands out (downloads, upload sessions) are on SharePoint.
function sharePointUrl(raw: string): string {
  const url = URL.parse(raw);
  if (url?.protocol !== "https:" || !/\.sharepoint\.com$/i.test(url.hostname))
    throw new ConnectError(
      `Microsoft pointed to an address that is not SharePoint: ${url?.hostname ?? "none"}`,
    );
  return url.toString();
}

type OrgInfo = { displayName?: string; verifiedDomains?: { name: string; isDefault: boolean }[] };

/** Apps that need only the Microsoft connection: Outlook mail and calendar (mailboxes are
 *  reached through the customer's Exchange role assignments to the app, scoped to the
 *  branding's mailbox group, never through Graph Mail or Calendars permissions), Teams, Entra ID
 *  and Intune. `keep`: the app's own choice fields kept in its config; nothing else the form
 *  sends is, so no field can shadow the provider's config. */
export function mailboxVendor(ms: MicrosoftClient, keep: readonly string[] = []): Vendor {
  return {
    async connect(fields, provider) {
      if (!provider) throw new Error("Microsoft apps connect through the Microsoft connection");
      const config = Object.fromEntries(keep.filter((k) => fields[k]).map((k) => [k, fields[k]]));
      return { config, secret: {}, accountLabel: provider.accountLabel };
    },
    call: ms.call,
    checks: { permissions: ms.permissions },
  };
}

// A Microsoft app keeps its own choice fields (its JSON) and nothing else.
export const microsoftApp = (id: string, toolset: Toolset): BuiltinModule => ({
  id,
  create: ({ microsoft, connector }) => ({
    vendor: mailboxVendor(microsoft, choiceFields(connector)),
    toolset,
  }),
});
