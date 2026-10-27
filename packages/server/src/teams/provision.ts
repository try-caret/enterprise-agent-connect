import { X509Certificate } from "node:crypto";
import { brand } from "../../../connector-kit/src/branding.js";
import { describeCertificate } from "../../../connector-kit/src/certificate.js";
import {
  ConnectError,
  type Fetch,
  type VendorRequest,
  type VendorResponse,
} from "../../../connector-kit/src/types.js";
import { RSC_PERMISSIONS } from "./package.js";

// "Add to Teams": what the Microsoft connection's own app (branding.microsoft.appName) does in the
// customer's tenant to make an agent's bot, app-only: the bot's app registration and service
// principal (Application.ReadWrite.OwnedBy), its Azure Bot with the Teams channel (a custom
// role on one resource group), and installing it where the admin picks
// (TeamsAppInstallation.ReadWriteAndConsentFor{Team,Chat,User}.All). Publishing to the org's
// Teams app catalog has no app-only permission, so an administrator signs in for it
// (delegated AppCatalog.ReadWrite.All). docs/teams.md "Add to Teams".

export const ARM = "https://management.azure.com";
const BOT_API = "2022-09-15";
export const CATALOG_SCOPE = "https://graph.microsoft.com/AppCatalog.ReadWrite.All";
const GRAPH = "https://graph.microsoft.com";
const LOGIN = "https://login.microsoftonline.com";

/** The live checks the Teams page's set-up-once steps name (`live` in connectors/teams.json). */
export const LIVE_CHECKS = ["azure-target", "azure-role"] as const;

export type Call = (request: VendorRequest, origin?: string) => Promise<VendorResponse>;

export type StepState = {
  id: "app" | "principal" | "azureBot" | "channel" | "signIn";
  label: string;
  state: "done" | "failed" | "waiting" | "todo";
  message?: string;
};

export type AzureTarget = { subscriptionId: string; resourceGroup: string };

export function azureTarget(config: Record<string, string>): AzureTarget | null {
  const { azureSubscriptionId: subscriptionId, azureResourceGroup: resourceGroup } = config;
  return subscriptionId && resourceGroup ? { subscriptionId, resourceGroup } : null;
}

// The same name the setup script gives it, so either way's removal finds it.
export const azureBotName = (botAppId: string) =>
  `${brand().resourcePrefix}-${botAppId.slice(0, 8)}`;

const botPath = (t: AzureTarget, name: string) =>
  `/subscriptions/${t.subscriptionId}/resourceGroups/${encodeURIComponent(t.resourceGroup)}/providers/Microsoft.BotService/botServices/${name}`;

function errorOf(data: unknown): { code: string; message: string } {
  const e = (data as { error?: { code?: unknown; message?: unknown } } | null)?.error;
  return {
    code: typeof e?.code === "string" ? e.code : "",
    message: (typeof e?.message === "string" ? e.message : "").split(/\r?\n/)[0].slice(0, 300),
  };
}

/** A Graph refusal in plain words. A missing permission names it. */
export function graphFailure(res: VendorResponse, what: string, permission: string) {
  const { code, message } = errorOf(res.data);
  if (res.status === 403 && (code === "Authorization_RequestDenied" || code === "Forbidden"))
    return new ConnectError(
      `Microsoft refused ${what}: ${brand().microsoft.appName} needs the ${permission} permission with admin consent (Add the Microsoft Graph permissions for adding bots, then Grant admin consent, on the Microsoft Teams page).`,
    );
  return new ConnectError(
    `Microsoft refused ${what} (${res.status}${code ? `, ${code}` : ""})${message ? `: ${message}` : ""}`,
  );
}

/** An Azure Resource Manager refusal in plain words. */
export function armFailure(
  res: VendorResponse,
  target: AzureTarget,
  what: string,
  // Shown on the role step itself, which then points at "this step".
  onRoleStep = false,
) {
  const { code, message } = errorOf(res.data);
  if (code === "AuthorizationFailed" || code === "LinkedAuthorizationFailed")
    return new ConnectError(
      `${brand().microsoft.appName} has no role on the resource group ${target.resourceGroup} yet (${onRoleStep ? "this step" : `Give ${brand().microsoft.appName} a role on that resource group, on the Microsoft Teams page`}). A new role assignment can take a few minutes to apply.`,
    );
  if (code === "MissingSubscriptionRegistration")
    return new ConnectError(
      `The subscription ${target.subscriptionId} is not registered for Azure Bot Service yet. Register Microsoft.BotService (Register the Bot Service in the subscription, on the Microsoft Teams page), then try again.`,
    );
  if (code === "ResourceGroupNotFound")
    return new ConnectError(
      `There is no resource group ${target.resourceGroup} in the subscription ${target.subscriptionId}.`,
    );
  if (code === "SubscriptionNotFound" || code === "InvalidSubscriptionId")
    return new ConnectError(
      `Azure has no subscription ${target.subscriptionId} that ${brand().microsoft.appName} can see.`,
    );
  return new ConnectError(
    `Azure refused ${what} (${res.status}${code ? `, ${code}` : ""})${message ? `: ${message}` : ""}`,
  );
}

const ok = (res: VendorResponse) => res.status >= 200 && res.status < 300;

// Graph returns a certificate's customKeyIdentifier as its SHA-1 thumbprint, base64.
function hasCertificate(keys: unknown, certificate: string): boolean {
  const thumbprint = describeCertificate(certificate).thumbprint.toUpperCase();
  return (Array.isArray(keys) ? keys : []).some((k) => {
    const id = (k as { customKeyIdentifier?: unknown }).customKeyIdentifier;
    if (typeof id !== "string") return false;
    const bytes = Buffer.from(id, "base64");
    return (
      bytes.toString("hex").toUpperCase() === thumbprint ||
      bytes.toString("utf8").toUpperCase() === thumbprint
    );
  });
}

export type ProvisionInput = {
  call: Call;
  bot: { id: string; displayName: string; tenantId: string; certificate: string };
  // What earlier tries saved (appObjectId, appId), and where to save more as it is made.
  config: Record<string, string>;
  save: (config: Record<string, string>) => Promise<void>;
  target: AzureTarget;
  // This service's public address (…/agent-connect), where Teams sends the bot's messages.
  publicUrl: string;
  // Signs in as the bot; Microsoft can take a while to know a new app.
  prove: (botAppId: string) => Promise<void>;
  sleep: (ms: number) => Promise<void>;
  waits?: number[];
};

const LABELS: Record<StepState["id"], string> = {
  app: "App registration with the bot's certificate",
  principal: "Its service principal",
  azureBot: "Azure Bot on the free tier",
  channel: "Teams channel on the Azure Bot",
  signIn: "The bot signs in to Microsoft",
};

// Retried while Microsoft replicates a new app (a few seconds, sometimes a minute).
const REPLICATION_MS = [2000, 4000, 8000];

/** Makes the bot in the tenant, each part only if missing, so running it again finishes what
 *  an earlier try left. Steps after a failed one are left to do. */
export async function provisionBot(input: ProvisionInput): Promise<{
  steps: StepState[];
  botAppId: string | null;
}> {
  const { call, bot, target } = input;
  const waits = input.waits ?? REPLICATION_MS;
  const config = { ...input.config };
  const steps: StepState[] = [];
  let failed = false;

  async function step(id: StepState["id"], run: () => Promise<string | null>) {
    if (failed) return void steps.push({ id, label: LABELS[id], state: "todo" });
    try {
      // A message back means it is not done yet, only waiting on Microsoft.
      const waiting = await run();
      if (waiting) failed = true;
      steps.push({
        id,
        label: LABELS[id],
        ...(waiting ? { state: "waiting" as const, message: waiting } : { state: "done" as const }),
      });
    } catch (error) {
      if (!(error instanceof ConnectError)) throw error;
      failed = true;
      steps.push({ id, label: LABELS[id], state: "failed", message: error.message });
    }
  }

  async function retried(
    attempt: () => Promise<VendorResponse>,
    again: (r: VendorResponse) => boolean,
  ) {
    let res = await attempt();
    for (const wait of waits) {
      if (!again(res)) break;
      await input.sleep(wait);
      res = await attempt();
    }
    return res;
  }

  const uniqueName = `${brand().resourcePrefix}-bot-${bot.id}`;
  await step("app", async () => {
    const upsert = await call({
      method: "PATCH",
      path: `/v1.0/applications(uniqueName='${uniqueName}')`,
      headers: { prefer: "create-if-missing" },
      body: {
        displayName: `${bot.displayName} (${brand().productName})`,
        signInAudience: "AzureADMyOrg",
      },
    });
    if (!ok(upsert))
      throw graphFailure(
        upsert,
        "creating the bot's app registration",
        "Application.ReadWrite.OwnedBy",
      );
    const read = await retried(
      () =>
        call({
          method: "GET",
          path: `/v1.0/applications(uniqueName='${uniqueName}')?$select=id,appId,keyCredentials`,
        }),
      (r) => r.status === 404,
    );
    if (!ok(read))
      throw graphFailure(
        read,
        "reading the bot's app registration",
        "Application.ReadWrite.OwnedBy",
      );
    const app = read.data as { id: string; appId: string; keyCredentials?: unknown };
    config.appObjectId = app.id;
    config.appId = app.appId;
    await input.save(config);
    // The service owns this app's credentials outright (it is the bot's own, found by its unique
    // name), so replacing them all with the current certificate is intended.
    if (!hasCertificate(app.keyCredentials, bot.certificate)) {
      const set = await call({
        method: "PATCH",
        path: `/v1.0/applications/${app.id}`,
        body: {
          keyCredentials: [
            {
              type: "AsymmetricX509Cert",
              usage: "Verify",
              key: new X509Certificate(bot.certificate).raw.toString("base64"),
              displayName: brand().productName,
            },
          ],
        },
      });
      if (!ok(set))
        throw graphFailure(set, "adding the bot's certificate", "Application.ReadWrite.OwnedBy");
    }
    return null;
  });

  const appId = config.appId ?? null;
  await step("principal", async () => {
    const found = await call({
      method: "GET",
      path: `/v1.0/servicePrincipals(appId='${appId}')?$select=id`,
    });
    if (ok(found)) return null;
    if (found.status !== 404)
      throw graphFailure(
        found,
        "reading the bot's service principal",
        "Application.ReadWrite.OwnedBy",
      );
    const made = await retried(
      () => call({ method: "POST", path: "/v1.0/servicePrincipals", body: { appId } }),
      // A new app is not known everywhere at once.
      (r) => r.status === 400 || r.status === 404,
    );
    if (!ok(made) && made.status !== 409)
      throw graphFailure(
        made,
        "creating the bot's service principal",
        "Application.ReadWrite.OwnedBy",
      );
    return null;
  });

  const path = appId ? botPath(target, azureBotName(appId)) : "";
  await step("azureBot", async () => {
    // The name holds only 8 hex of the app id, so one made for another app may hold it.
    const found = await call({ method: "GET", path: `${path}?api-version=${BOT_API}` }, ARM);
    if (!ok(found) && found.status !== 404)
      throw armFailure(found, target, "reading the Azure Bot");
    const owner = ok(found)
      ? (found.data as { properties?: { msaAppId?: unknown } } | null)?.properties?.msaAppId
      : appId;
    if (owner !== appId)
      throw new ConnectError(
        `An Azure Bot named ${azureBotName(appId!)} in ${target.resourceGroup} belongs to another app, so it is left as it is. Rename or remove it, then add the bot again.`,
      );
    const res = await call(
      {
        method: "PUT",
        path: `${path}?api-version=${BOT_API}`,
        body: {
          location: "global",
          kind: "azurebot",
          sku: { name: "F0" },
          properties: {
            displayName: bot.displayName,
            endpoint: `${input.publicUrl}/teams/bots/${appId}/messages`,
            msaAppId: appId,
            msaAppType: "SingleTenant",
            msaAppTenantId: bot.tenantId,
          },
        },
      },
      ARM,
    );
    if (!ok(res)) throw armFailure(res, target, "creating the Azure Bot");
    config.azureBot = path;
    await input.save(config);
    return null;
  });

  await step("channel", async () => {
    const res = await call(
      {
        method: "PUT",
        path: `${path}/channels/MsTeamsChannel?api-version=${BOT_API}`,
        body: {
          location: "global",
          properties: { channelName: "MsTeamsChannel", properties: { isEnabled: true } },
        },
      },
      ARM,
    );
    if (!ok(res)) throw armFailure(res, target, "adding the Teams channel");
    return null;
  });

  let signedIn = false;
  await step("signIn", async () => {
    let last: ConnectError | null = null;
    for (const wait of [0, ...waits]) {
      if (wait) await input.sleep(wait);
      try {
        await input.prove(appId!);
        signedIn = true;
        return null;
      } catch (error) {
        if (!(error instanceof ConnectError)) throw error;
        last = error;
      }
    }
    return `Microsoft is still setting up the new app (${last!.message}). Select Add to Teams again in a minute: nothing is made twice.`;
  });

  return { steps, botAppId: signedIn ? appId : null };
}

/** Removes what provisioning made: the Azure Bot and the app registration (which also ends
 *  its service principal). Each failure is returned in plain words, never dropped. The app
 *  registration is kept while its Azure Bot is: the removal script deletes a bot only when its
 *  app still carries the bot's unique name. */
export async function deprovisionBot(call: Call, config: Record<string, string>) {
  const failed: string[] = [];
  if (config.azureBot) {
    const res = await call(
      { method: "DELETE", path: `${config.azureBot}?api-version=${BOT_API}` },
      ARM,
    );
    if (!ok(res) && res.status !== 404) {
      const { code, message } = errorOf(res.data);
      return [
        `Azure did not delete the bot (${res.status}${code ? `, ${code}` : ""})${message ? `: ${message}` : ""}`,
      ];
    }
  }
  if (config.appObjectId) {
    const res = await call({ method: "DELETE", path: `/v1.0/applications/${config.appObjectId}` });
    if (!ok(res) && res.status !== 404)
      failed.push(
        graphFailure(res, "deleting the bot's app registration", "Application.ReadWrite.OwnedBy")
          .message,
      );
  }
  return failed;
}

// ---- the set-up-once steps, checked live ----

// What Add to Teams writes in the resource group.
const BOT_WRITES = [
  "Microsoft.BotService/botServices/write",
  "Microsoft.BotService/botServices/channels/write",
];

// Azure's permissions answer: one entry per role, each allowing its `actions` (with *
// wildcards, any case) except its `notActions`.
function allows(data: unknown, action: string): boolean {
  const matches = (pattern: unknown) =>
    typeof pattern === "string" &&
    new RegExp(
      `^${pattern.replace(/[.+?^${}()|[\]\\]/g, "\\$&").replaceAll("*", ".*")}$`,
      "i",
    ).test(action);
  const entries = (data as { value?: unknown } | null)?.value;
  return (Array.isArray(entries) ? entries : []).some(
    (e: { actions?: unknown; notActions?: unknown }) =>
      Array.isArray(e.actions) &&
      e.actions.some(matches) &&
      !(Array.isArray(e.notActions) && e.notActions.some(matches)),
  );
}

export type LiveCheck = { state: "done" | "failing" | "todo"; message: string };

export async function setupChecks(
  call: Call,
  config: Record<string, string>,
): Promise<Record<(typeof LIVE_CHECKS)[number], LiveCheck>> {
  const target = azureTarget(config);
  if (!target) {
    const todo = {
      state: "todo" as const,
      message: "Enter the subscription and resource group first.",
    };
    return { "azure-target": { state: "todo", message: "Not entered yet." }, "azure-role": todo };
  }
  const saved: LiveCheck = {
    state: "done",
    message: `Bots go in the resource group ${target.resourceGroup} of the subscription ${target.subscriptionId}.`,
  };
  let role: LiveCheck;
  try {
    // What the Microsoft app may do there (all its role assignments, inherited ones too): seeing
    // the group is not enough, since Add to Teams writes the bot and its channel.
    const res = await call(
      {
        method: "GET",
        path: `/subscriptions/${target.subscriptionId}/resourceGroups/${encodeURIComponent(target.resourceGroup)}/providers/Microsoft.Authorization/permissions?api-version=2022-04-01`,
      },
      ARM,
    );
    const missing = ok(res) ? BOT_WRITES.filter((a) => !allows(res.data, a)) : [];
    role = !ok(res)
      ? {
          state: "failing",
          message: armFailure(res, target, "reading its permissions", true).message,
        }
      : missing.length === BOT_WRITES.length
        ? {
            state: "failing",
            message: `${brand().microsoft.appName} has no role on the resource group ${target.resourceGroup} that can create Azure Bots yet (this step). A new role assignment can take a few minutes to apply.`,
          }
        : missing.length
          ? {
              state: "failing",
              message: `${brand().microsoft.appName}'s role on ${target.resourceGroup} lacks ${missing.join(" and ")} (this step).`,
            }
          : {
              state: "done",
              message: `Checked live: ${brand().microsoft.appName} can create and change Azure Bots in ${target.resourceGroup}.`,
            };
  } catch (error) {
    if (!(error instanceof ConnectError)) throw error;
    role = { state: "failing", message: error.message };
  }
  const missingGroup =
    role.state === "failing" && role.message.startsWith("There is no resource group");
  return {
    "azure-target": missingGroup ? { state: "failing", message: role.message } : saved,
    "azure-role": role,
  };
}

// ---- publishing to the org's Teams app catalog, as the administrator who signed in ----

/** The id token's claims must be this connection's app, tenant and sign-in, by someone of the
 *  organization (a guest names another home in `idp`). */
export function checkSignIn(
  claims: Record<string, unknown>,
  expect: { clientId: string; tenantId: string; nonce: string },
  now = Date.now(),
) {
  const fresh = typeof claims.exp === "number" && claims.exp * 1000 > now;
  if (
    claims.aud !== expect.clientId ||
    claims.iss !== `${LOGIN}/${expect.tenantId}/v2.0` ||
    claims.nonce !== expect.nonce ||
    !fresh
  )
    throw new ConnectError("Microsoft's sign-in did not match this request. Start again.");
  const guest = typeof claims.idp === "string" && !claims.idp.includes(expect.tenantId);
  if (claims.tid !== expect.tenantId || guest)
    throw new ConnectError(
      "Sign in with an administrator of your own organization, not a guest or another organization's account.",
    );
}

function catalogFailure(res: VendorResponse): ConnectError {
  const error = (
    res.data as { error?: { innerError?: { code?: string; details?: unknown } } } | null
  )?.error;
  const inner = error?.innerError;
  if (inner?.code === "UnableToParseTeamsAppManifest") {
    const details = (Array.isArray(inner.details) ? inner.details : [])
      .map(
        (d: { target?: string; message?: string }) =>
          `${d.target ? `${d.target}: ` : ""}${d.message ?? ""}`,
      )
      .join("; ");
    return new ConnectError(
      `Microsoft refused the Teams app package: ${details || errorOf(res.data).message}`,
    );
  }
  if (res.status === 403 || res.status === 401)
    return new ConnectError(
      `Microsoft refused to publish: sign in as a Teams Administrator or Global Administrator, and check that ${brand().microsoft.appName} has AppCatalog.ReadWrite.All with admin consent.`,
    );
  if (res.status === 409)
    return new ConnectError(
      "Your Teams apps already have an app for this bot under another ID. Delete it in the Teams admin center under Teams apps, then Manage apps, and publish again.",
    );
  const { code, message } = errorOf(res.data);
  return new ConnectError(
    `Microsoft refused to publish (${res.status}${code ? `, ${code}` : ""})${message ? `: ${message}` : ""}`,
  );
}

/** Publishes the package, or a new version of it, under the manifest id `externalId`. Finds the
 *  catalog app first, so publishing twice changes nothing. */
export async function publishPackage(
  input: { token: string; externalId: string; version: string; zip: Buffer },
  fetchImpl: Fetch = fetch,
): Promise<{ catalogAppId: string; changed: boolean }> {
  const as = async (path: string, init: { method: string; zip?: Buffer }) => {
    const res = await fetchImpl(`${GRAPH}${path}`, {
      method: init.method,
      redirect: "error",
      headers: {
        authorization: `Bearer ${input.token}`,
        accept: "application/json",
        ...(init.zip ? { "content-type": "application/zip" } : {}),
      },
      body: init.zip ? new Uint8Array(init.zip) : undefined,
    });
    const text = await res.text();
    let data: unknown = null;
    try {
      data = text ? JSON.parse(text) : null;
    } catch {
      data = text;
    }
    return { status: res.status, data };
  };
  const q = new URLSearchParams({
    $filter: `externalId eq '${input.externalId}'`,
    $expand: "appDefinitions($select=version)",
  });
  const found = await as(`/v1.0/appCatalogs/teamsApps?${q}`, { method: "GET" });
  if (!ok(found) || errorOf(found.data).code) throw catalogFailure(found);
  const existing = (
    found.data as { value?: { id: string; appDefinitions?: { version?: string }[] }[] }
  ).value?.[0];
  if (existing?.appDefinitions?.some((d) => d.version === input.version))
    return { catalogAppId: existing.id, changed: false };
  const res = existing
    ? await as(`/v1.0/appCatalogs/teamsApps/${existing.id}/appDefinitions?requiresReview=false`, {
        method: "POST",
        zip: input.zip,
      })
    : await as("/v1.0/appCatalogs/teamsApps?requiresReview=false", {
        method: "POST",
        zip: input.zip,
      });
  // Microsoft can answer a refused package with 200 and an error body.
  if (!ok(res) || errorOf(res.data).code) throw catalogFailure(res);
  const id = existing?.id ?? (res.data as { id?: string } | null)?.id;
  if (!id) throw new ConnectError("Microsoft did not say which app it published");
  return { catalogAppId: id, changed: true };
}

// ---- installing where the admin picks ----

export type InstallKind = "team" | "chat" | "person";

const INSTALL: Record<
  InstallKind,
  { path: (id: string) => string; permission: string; what: string }
> = {
  team: {
    path: (id) => `/v1.0/teams/${id}/installedApps`,
    permission: "TeamsAppInstallation.ReadWriteAndConsentForTeam.All",
    what: "team",
  },
  chat: {
    path: (id) => `/v1.0/chats/${encodeURIComponent(id)}/installedApps`,
    permission: "TeamsAppInstallation.ReadWriteAndConsentForChat.All",
    what: "chat",
  },
  person: {
    path: (id) => `/v1.0/users/${id}/teamwork/installedApps`,
    permission: "TeamsAppInstallation.ReadWriteAndConsentForUser.All",
    what: "person",
  },
};

/** Adds the published app there, consenting to the package's resource-specific permissions
 *  (Microsoft requires exactly the package's list). Already there counts as added. */
export async function installBot(
  call: Call,
  catalogAppId: string,
  kind: InstallKind,
  id: string,
): Promise<"added" | "already"> {
  const spec = INSTALL[kind];
  const res = await call({
    method: "POST",
    path: spec.path(id),
    body: {
      "teamsApp@odata.bind": `${GRAPH}/v1.0/appCatalogs/teamsApps/${catalogAppId}`,
      consentedPermissionSet: {
        resourceSpecificPermissions: RSC_PERMISSIONS.map((permissionValue) => ({
          permissionValue,
          permissionType: "application",
        })),
      },
    },
  });
  if (ok(res)) return "added";
  if (res.status === 409) return "already";
  const { code, message } = errorOf(res.data);
  if (res.status === 403 && code !== "Authorization_RequestDenied")
    throw new ConnectError(
      `Microsoft would not add the bot to this ${spec.what}. Your organization's Teams app settings may block it (the app's status, who it is available to, or a permission policy)${message ? `: ${message}` : "."}`,
    );
  if (res.status === 404)
    throw new ConnectError(
      `Microsoft found no such ${spec.what}, or the bot is not in your Teams apps yet. Publish it first.`,
    );
  throw graphFailure(res, `adding the bot to this ${spec.what}`, spec.permission);
}

/** Chat ids from a pasted Teams link or the id itself. */
export function chatIdFrom(raw: string): string | null {
  let text = raw.trim();
  try {
    text = decodeURIComponent(text);
  } catch {
    return null;
  }
  return /19:[A-Za-z0-9_.-]+@thread\.(v2|skype)/.exec(text)?.[0] ?? null;
}

/** Teams or people matching what the admin typed, for the picker. */
export async function searchTargets(
  call: Call,
  kind: "team" | "person",
  rawQuery: string,
): Promise<{ id: string; name: string; detail: string | null }[]> {
  const query = rawQuery.replace(/["\\]/g, "").trim().slice(0, 100);
  if (kind === "team") {
    const q = new URLSearchParams({ $select: "id,displayName,description", $top: "20" });
    if (query) q.set("$filter", `startswith(displayName,'${query.replaceAll("'", "''")}')`);
    const res = await call({ method: "GET", path: `/v1.0/teams?${q}` });
    if (!ok(res)) throw graphFailure(res, "listing teams", "Team.ReadBasic.All");
    return (
      (res.data as { value?: { id: string; displayName?: string; description?: string }[] })
        .value ?? []
    ).map((t) => ({ id: t.id, name: t.displayName ?? t.id, detail: t.description || null }));
  }
  if (!query) return [];
  const q = new URLSearchParams({
    $search: `"displayName:${query}" OR "mail:${query}"`,
    $select: "id,displayName,mail",
    $top: "20",
    $count: "true",
  });
  const res = await call({
    method: "GET",
    path: `/v1.0/users?${q}`,
    headers: { consistencylevel: "eventual" },
  });
  if (!ok(res)) throw graphFailure(res, "looking up people", "User.ReadBasic.All");
  return (
    (res.data as { value?: { id: string; displayName?: string; mail?: string }[] }).value ?? []
  ).map((u) => ({ id: u.id, name: u.displayName ?? u.id, detail: u.mail ?? null }));
}
