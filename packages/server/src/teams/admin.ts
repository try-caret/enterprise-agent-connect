import { randomUUID } from "node:crypto";
import { z } from "zod";
import { brand } from "../../../connector-kit/src/branding.js";
import { describeCertificate, newCertificate } from "../../../connector-kit/src/certificate.js";
import { signInNonce, signState, stateApp, verifyState } from "../consent.js";
import { credentialAad, type CredentialKey } from "../secrets.js";
import { InvalidInput } from "../service.js";
import type { Store } from "../store.js";
import { SHARED_APP_AUTH, type MicrosoftClient } from "../../../../connectors/microsoft/client.js";
import { ConnectError, type Fetch } from "../../../connector-kit/src/types.js";
import type { BotClient } from "./bot.js";
import { agentMcpUrl, type ClaudeFactory } from "./claude.js";
import { packageZip, removalScript, setupScript } from "./package.js";
import {
  CATALOG_SCOPE,
  LIVE_CHECKS,
  azureTarget,
  chatIdFrom,
  checkSignIn,
  deprovisionBot,
  installBot,
  provisionBot,
  publishPackage,
  searchTargets,
  setupChecks,
  type Call,
  type InstallKind,
  type LiveCheck,
} from "./provision.js";
import { botAad, botCredential } from "./runtime.js";
import { CLAUDE_RUNNER, openTriggerSecret, triggerAad, type TriggerRunner } from "./runner.js";
import type { BotRow, RunRow, TeamsStore, TriggerRow } from "./store.js";

// The admin side of Teams: each agent's bot (made, proven, rotated, removed), its package and
// setup script, its real-time settings and recent responses, and the organization's Claude
// API key. docs/teams.md 9.

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/;
const CLAUDE_SCOPE = "provider:anthropic";

/** The setup script passes the bot's name to az, which runs through cmd.exe on Windows, where
 *  quoting cannot make & " | ^ % < > inert; so only characters inert there are allowed. */
export const BOT_NAME = /^[A-Za-z0-9 ._-]{1,30}$/;

export const triggerSchema = z
  .object({
    enabled: z.boolean(),
    // What answers (runner.ts); Claude Managed Agents unless named.
    runner: z
      .string()
      .regex(/^[a-z][a-z0-9-]{1,40}$/)
      .default(CLAUDE_RUNNER),
    // Claude Managed Agents' settings, kept where they were before there were other runners.
    managedAgentId: z
      .string()
      .regex(/^[A-Za-z0-9_-]{1,128}$/)
      .optional(),
    environmentId: z
      .string()
      .regex(/^[A-Za-z0-9_-]{1,128}$/)
      .optional(),
    vaultIds: z
      .array(z.string().regex(/^[A-Za-z0-9_-]{1,128}$/))
      .max(20)
      .default([]),
    budgetCents: z.number().int().min(1).max(2000),
    maxRunsPerHour: z.number().int().min(1).max(100).default(20),
    maxCentsPerDay: z.number().int().min(1).max(100_000).default(500),
    maxConcurrent: z.number().int().min(1).max(10).default(2),
    replyMode: z.enum(["mentions", "all"]).default("mentions"),
    allowedTeams: z.array(z.string().regex(GUID)).max(500).default([]),
    allowedPeople: z.array(z.string().regex(GUID)).max(500).default([]),
    // Another runner's settings (e.g. the signed webhook's url), which that runner checks.
    settings: z.record(z.string(), z.unknown()).default({}),
    // The runner's signing secret (the signed webhook's): never shown again, kept until a
    // new one is given.
    signingSecret: z.string().min(32).max(256).optional(),
    // Whoever may ask gets everything the agent may do, so "anyone" is a choice, never a default.
    anyoneInOrganization: z.boolean().default(false),
    by: z.string().trim().min(1).max(200),
  })
  .refine((t) => !t.enabled || t.anyoneInOrganization || t.allowedPeople.length > 0, {
    message: "Name the people who may ask it, or allow anyone in your organization",
    path: ["allowedPeople"],
  })
  .refine((t) => t.maxCentsPerDay >= t.budgetCents, {
    message: "The daily limit must be at least the limit per answer",
    path: ["maxCentsPerDay"],
  });

const certificateInfo = (pem: string | undefined) => (pem ? describeCertificate(pem) : null);

/** A trigger as the admin sees it: whether it has a signing secret, never the secret. */
function publicTrigger(row: TriggerRow) {
  const { secret, secretKeyId: _k, ...rest } = row;
  return { ...rest, hasSigningSecret: secret !== null };
}

const packageVersion = (bot: BotRow) => `1.0.${bot.manifestVersion}`;

function publicBot(bot: BotRow) {
  const { credential: _c, credentialKeyId: _k, config, ...rest } = bot;
  return {
    ...rest,
    certificate: certificateInfo(config.certificate),
    nextCertificate: certificateInfo(config.nextCertificate),
    // Made by Add to Teams (so removed by the service too), not by the setup script.
    madeByService: Boolean(config.appObjectId),
    // The package version in the Teams app catalog, once published from here.
    publishedVersion: config.publishedVersion ?? null,
    packageVersion: packageVersion(bot),
  };
}

const PUBLISH_APP =
  /^teams-publish:([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})$/;

export function createTeamsAdmin(deps: {
  store: Store;
  teams: TeamsStore;
  key: CredentialKey;
  bots: BotClient;
  claude: ClaudeFactory;
  // By id: what answers each trigger (runner.ts).
  runners: Map<string, TriggerRunner>;
  // This service's public address, e.g. https://agents.example.com/agent-connect. Null when the
  // deployment has none, which Teams needs to reach the bots.
  publicUrl: string | null;
  // The Microsoft connection's own app: Add to Teams and the administrator's sign-in to publish.
  microsoft: Pick<MicrosoftClient, "call" | "delegated">;
  fetch?: Fetch;
  sleep?: (ms: number) => Promise<void>;
  waits?: number[];
}) {
  const { store, teams, key } = deps;
  const stateKey = key.derive("teams-publish-state");
  const sleep = deps.sleep ?? ((ms: number) => new Promise((r) => setTimeout(r, ms)));

  // The Microsoft connection, signing in as the customer's own app (branding.microsoft.appName).
  async function microsoftFor(org: string) {
    const row = await store.getProvider(org, "microsoft");
    if (!row || row.status === "pending") throw new ConnectError("Connect Microsoft 365 first");
    if (row.config.auth === SHARED_APP_AUTH)
      throw new ConnectError(
        "Add to Teams needs your own app registration for Microsoft 365. Use the setup script instead.",
      );
    const conn = {
      config: row.config,
      secret: key.open(
        { ciphertext: row.credential, keyId: row.credentialKeyId },
        credentialAad(org, row.credentialScope),
      ),
    };
    const call: Call = (request, origin) => deps.microsoft.call(conn, request, origin);
    return { config: row.config, conn, call };
  }

  // Bot config changes (what Add to Teams made, what was published) over the credential read.
  async function saveConfig(b: BotRow, config: Record<string, string>) {
    if (
      !(await teams.updateBotCredential(
        b.id,
        config,
        b.credential,
        b.credentialKeyId,
        b.credential,
      ))
    )
      throw new ConnectError("This bot changed meanwhile. Reload the page.");
  }

  // The Teams page's opt-in to adding bots for one person (its own permission).
  const personInstalls = async (org: string) =>
    (await store.getConnection(org, "teams"))?.config.person_installs === "on";
  // Its choice of making and adding bots here (the permissions for that), or only scripts.
  const addsBots = async (org: string) =>
    (await store.getConnection(org, "teams"))?.config.add_to_teams !== "off";
  const addsOff = () =>
    `Making and adding bots is off. Choose "${brand().productName} makes and adds them" in step 1 of the Microsoft Teams page and add the permissions it shows, or use the setup script.`;

  const publicUrl = () => {
    if (!deps.publicUrl)
      throw new ConnectError(
        "This deployment has no public address for Teams to reach. Set global.publicBaseUrl in its settings.",
      );
    return deps.publicUrl;
  };

  async function log(org: string, by: string, action: string, target: string) {
    await store.audit({
      organizationId: org,
      agentId: null,
      agentName: by,
      connectorId: "teams",
      tool: `admin:${action}`,
      target,
      decision: "allowed",
      reason: null,
      vendorStatus: null,
    });
  }

  async function bot(org: string, agentId: string) {
    const found = await teams.botForAgent(org, agentId);
    if (!found) throw new InvalidInput("This agent has no Teams bot yet");
    return found;
  }

  const open = (b: BotRow) =>
    key.open(
      { ciphertext: b.credential, keyId: b.credentialKeyId },
      botAad(b.organizationId, b.id),
    );

  async function claudeFor(org: string) {
    const row = await teams.claudeKey(org);
    if (!row) throw new ConnectError("Set your organization's Claude API key first");
    const { api_key } = key.open(
      { ciphertext: row.credential, keyId: row.credentialKeyId },
      credentialAad(org, CLAUDE_SCOPE),
    );
    return deps.claude(api_key);
  }

  // Live sessions are stopped before what they use goes; `runner`: only that runner's.
  async function stopSessions(org: string, which: (run: RunRow) => boolean, runner?: string) {
    const bots = new Set((await teams.listBots(org)).map((b) => b.id));
    const live = (await teams.runningRuns()).filter(
      (r) => bots.has(r.botId) && which(r) && (!runner || r.runner === runner),
    );
    for (const run of live) {
      const answers = deps.runners.get(run.runner);
      if (answers) await answers.stop(run.sessionId!, org);
      await teams.updateRun(run.id, { state: "failed", error: "Stopped by an administrator" });
    }
  }

  return {
    async listBots(org: string) {
      const rows = await teams.listBots(org);
      const places = await teams.conversationsOf(rows.map((b) => b.id));
      // Where Teams sends this bot's messages, and the MCP address its Claude vault holds a key
      // for: shown to the admin for setting up by hand.
      const base = deps.publicUrl;
      return rows.map((b) => ({
        ...publicBot(b),
        messagingEndpoint: base
          ? `${base}/teams/bots/${b.botAppId ?? "<application ID>"}/messages`
          : null,
        mcpUrl: base ? agentMcpUrl(base, b.agentId) : null,
        installedIn: places
          .filter((p) => p.botId === b.id)
          .map((p) => ({
            kind: p.kind,
            name: p.name,
            id: p.teamGroupId ?? p.conversationId,
            conversationId: p.conversationId,
          })),
      }));
    },

    // A key pair for the agent's bot, kept until the admin runs the setup script and pastes
    // the app id it printed. The tenant is the Microsoft connection's.
    async createBot(org: string, agentId: string, displayName: string, by: string) {
      const name = displayName.trim();
      if (!BOT_NAME.test(name))
        throw new InvalidInput(
          "A bot name is 1 to 30 letters, digits, spaces, dots, dashes or underscores",
        );
      const provider = await store.getProvider(org, "microsoft");
      const tenantId = provider?.status === "connected" ? provider.config.tenantId : undefined;
      if (!tenantId) throw new ConnectError("Connect Microsoft 365 first");
      // Two bots with one name would look the same in Teams and in the Entra app list.
      const taken = (await teams.listBots(org)).find(
        (b) => b.agentId !== agentId && b.displayName.toLowerCase() === name.toLowerCase(),
      );
      if (taken) throw new InvalidInput(`Another agent's Teams bot is already named ${name}`);
      const id = randomUUID();
      const made = newCertificate(`${brand().productName} agent ${name}`);
      const sealed = key.seal({ private_key: made.privateKey }, botAad(org, id));
      const created = await teams.createBot({
        id,
        organizationId: org,
        agentId,
        tenantId,
        displayName: name,
        config: { certificate: made.certificate },
        credential: sealed.ciphertext,
        credentialKeyId: sealed.keyId,
        createdBy: by,
      });
      if (!created) {
        if (await teams.botForAgent(org, agentId))
          throw new InvalidInput("This agent already has a Teams bot");
        return null;
      }
      await log(org, by, "create_bot", `agent ${agentId} bot ${id} "${name}"`);
      return publicBot((await teams.botById(id))!);
    },

    async setupScript(org: string, agentId: string) {
      const b = await bot(org, agentId);
      return setupScript({
        botId: b.id,
        displayName: b.displayName,
        teamsAppId: b.teamsAppId,
        tenantId: b.tenantId,
        certificate: b.config.certificate,
        publicUrl: publicUrl(),
      });
    },

    // The app id the script printed, proven by signing in as it with the bot's key.
    async connectBot(org: string, agentId: string, rawAppId: string, by: string) {
      const botAppId = rawAppId.trim().toLowerCase();
      if (!GUID.test(botAppId))
        throw new InvalidInput("Paste the application ID the script printed");
      const b = await bot(org, agentId);
      if (b.botAppId && b.botAppId !== botAppId)
        throw new InvalidInput(
          "This bot is connected to another app. Remove it and set it up again.",
        );
      const provider = await store.getProvider(org, "microsoft");
      if (provider?.config.clientId === botAppId)
        throw new InvalidInput("That is the Microsoft connection's app, not the agent's bot");
      await deps.bots.prove({
        botAppId,
        tenantId: b.tenantId,
        certificate: b.config.certificate,
        privateKey: open(b).private_key,
      });
      if (!(await teams.connectBot(b.id, botAppId, b.credential)))
        throw new ConnectError(
          "Another bot already uses this app, or this one changed meanwhile. Reload the page.",
        );
      await log(org, by, "connect_bot", `bot ${b.id} app ${botAppId}`);
      return publicBot((await teams.botById(b.id))!);
    },

    async packageZip(org: string, agentId: string) {
      const b = await bot(org, agentId);
      if (!b.botAppId) throw new InvalidInput("Finish setting up the bot first");
      return packageZip({
        displayName: b.displayName,
        teamsAppId: b.teamsAppId,
        botAppId: b.botAppId,
        manifestVersion: b.manifestVersion,
        publicUrl: publicUrl(),
      });
    },

    // A certificate to rotate to: the admin appends it to the bot's app, then switches.
    async newCertificate(org: string, agentId: string, by: string) {
      const b = await bot(org, agentId);
      const made = newCertificate(`${brand().productName} agent ${b.displayName}`);
      const sealed = key.seal({ ...open(b), next_private_key: made.privateKey }, botAad(org, b.id));
      const saved = await teams.updateBotCredential(
        b.id,
        { ...b.config, nextCertificate: made.certificate },
        sealed.ciphertext,
        sealed.keyId,
        b.credential,
      );
      if (!saved) throw new ConnectError("This bot changed meanwhile. Reload the page.");
      await log(
        org,
        by,
        "create_certificate",
        `bot ${b.id} next ${describeCertificate(made.certificate).thumbprint}`,
      );
      return {
        bot: publicBot((await teams.botById(b.id))!),
        // The line that adds it to the bot's app, keeping the current one.
        command: b.botAppId
          ? `az ad app credential reset --id ${b.botAppId} --cert "@next.pem" --append`
          : null,
        certificate: made.certificate,
      };
    },

    // Signs in with the new certificate and, only if that works, switches to it.
    async activateCertificate(org: string, agentId: string, by: string) {
      const b = await bot(org, agentId);
      const { nextCertificate, ...config } = b.config;
      const { next_private_key, ...secret } = open(b);
      if (!nextCertificate || !next_private_key)
        throw new InvalidInput("There is no new certificate to switch to");
      if (b.botAppId)
        await deps.bots.prove({
          botAppId: b.botAppId,
          tenantId: b.tenantId,
          certificate: nextCertificate,
          privateKey: next_private_key,
        });
      const sealed = key.seal({ ...secret, private_key: next_private_key }, botAad(org, b.id));
      const saved = await teams.updateBotCredential(
        b.id,
        { ...config, certificate: nextCertificate },
        sealed.ciphertext,
        sealed.keyId,
        b.credential,
      );
      if (!saved) throw new ConnectError("This bot changed meanwhile. Reload the page.");
      await log(
        org,
        by,
        "switch_certificate",
        `bot ${b.id} ${describeCertificate(b.config.certificate).thumbprint} -> ${describeCertificate(nextCertificate).thumbprint}`,
      );
      return publicBot((await teams.botById(b.id))!);
    },

    // Signs in again, as the integrations page does for other connections.
    async checkBot(org: string, agentId: string) {
      const b = await bot(org, agentId);
      if (!b.botAppId) return { ok: false as const, error: "Not set up yet" };
      try {
        await deps.bots.prove(botCredential(key, b));
        await teams.markBot(b.id, true);
        return { ok: true as const };
      } catch (error) {
        if (!(error instanceof ConnectError)) throw error;
        await teams.markBot(b.id, false);
        return { ok: false as const, error: error.message };
      }
    },

    // The certificate's public half, for uploading to the app registration by hand.
    async certificate(org: string, agentId: string) {
      return (await bot(org, agentId)).config.certificate;
    },

    // A short message from the bot in one place it was added to, proving the whole path.
    async sendTest(org: string, agentId: string, conversationId: string, by: string) {
      const b = await bot(org, agentId);
      if (b.status !== "connected") throw new InvalidInput("Finish setting up the bot first");
      const place = (await teams.conversations(b.id)).find(
        (p) => p.conversationId === conversationId,
      );
      if (!place) throw new InvalidInput("The bot is not in that conversation any more");
      const res = await deps.bots.connector(botCredential(key, b), place.serviceUrl, {
        method: "POST",
        path: `/v3/conversations/${encodeURIComponent(place.conversationId)}/activities`,
        body: {
          type: "message",
          text: `Test message from ${b.displayName}, sent from ${brand().productName} by ${by}.`,
        },
      });
      if (res.status >= 300)
        throw new ConnectError(`Teams did not accept the message (status ${res.status})`);
      await log(org, by, "send_test", `bot ${b.id} conversation ${place.conversationId}`);
      return { ok: true as const };
    },

    // Forgets the bot and its key here. What Add to Teams made is deleted with API calls; each
    // one that fails is returned, with the commands that remove the rest from the tenant.
    async removeBot(org: string, agentId: string, by: string) {
      const b = await teams.botForAgent(org, agentId);
      if (!b) return null;
      await stopSessions(org, (run) => run.botId === b.id);
      const made = Boolean(b.config.appObjectId || b.config.azureBot);
      const failed: string[] = [];
      if (made)
        try {
          failed.push(...(await deprovisionBot((await microsoftFor(org)).call, b.config)));
        } catch (error) {
          if (!(error instanceof ConnectError)) throw error;
          failed.push(error.message);
        }
      if (!(await teams.deleteBot(org, agentId))) return null;
      await log(
        org,
        by,
        "remove_bot",
        `bot ${b.id} app ${b.botAppId ?? "-"}${made ? (failed.length ? `; not removed: ${failed.join("; ")}` : "; removed from the tenant") : ""}`,
      );
      return {
        removal: removalScript({
          botId: b.id,
          botAppId: b.botAppId ?? b.config.appId ?? null,
          teamsAppId: b.teamsAppId,
          catalogOnly: made && failed.length === 0,
        }),
        failed,
      };
    },

    // ---- Add to Teams ----

    // The Teams page's set-up-once steps that the service checks live (`live` in teams.json).
    async setupState(org: string): Promise<Record<string, LiveCheck>> {
      const row = await store.getProvider(org, "microsoft");
      if (!row || row.status === "pending") {
        const todo = { state: "todo" as const, message: "Connect Microsoft 365 first." };
        return Object.fromEntries(LIVE_CHECKS.map((id) => [id, todo]));
      }
      let microsoft;
      try {
        microsoft = await microsoftFor(org);
      } catch (error) {
        // E.g. the shared app: reloading never helps, so each step says why.
        if (!(error instanceof ConnectError)) throw error;
        const failing = { state: "failing" as const, message: error.message };
        return Object.fromEntries(LIVE_CHECKS.map((id) => [id, failing]));
      }
      return setupChecks(microsoft.call, microsoft.config);
    },

    // Makes the agent's bot in the tenant: its key pair here if missing, then the app
    // registration, service principal, Azure Bot and Teams channel, then a sign-in as it.
    // Safe to run again: each part is made only if missing.
    async addToTeams(org: string, agentId: string, displayName: string, by: string) {
      if (!(await store.getConnection(org, "teams")))
        throw new ConnectError("Turn on Microsoft Teams first");
      if (!(await addsBots(org))) throw new ConnectError(addsOff());
      const { call, config } = await microsoftFor(org);
      const target = azureTarget(config);
      if (!target)
        throw new ConnectError(
          "Enter the Azure subscription and resource group on the Microsoft Teams page first",
        );
      if (!(await teams.botForAgent(org, agentId)))
        if (!(await this.createBot(org, agentId, displayName, by)))
          throw new InvalidInput("No such agent");
      const b = (await teams.botForAgent(org, agentId))!;
      if (b.botAppId) throw new InvalidInput("This agent's bot is already set up");
      if (b.tenantId !== config.tenantId)
        throw new ConnectError(
          "This bot was made for another Microsoft organization. Remove it and add it again.",
        );
      // Logged before any check below: by then the run may have made things in the tenant.
      const result = await provisionBot({
        call,
        bot: {
          id: b.id,
          displayName: b.displayName,
          tenantId: b.tenantId,
          certificate: b.config.certificate,
        },
        config: b.config,
        save: async (next) => {
          try {
            await saveConfig(b, next);
          } catch (error) {
            if (!(error instanceof ConnectError) || (await teams.botById(b.id))) throw error;
            // Removed while this ran: nothing would record what this run made, so undo it.
            const left = await deprovisionBot(call, next);
            throw new ConnectError(
              left.length
                ? `This bot was removed while it was being added. Remove what is left by hand: ${left.join("; ")}`
                : "This bot was removed while it was being added, so what was made for it is deleted again.",
            );
          }
        },
        target,
        publicUrl: publicUrl(),
        prove: (botAppId) =>
          deps.bots.prove({
            botAppId,
            tenantId: b.tenantId,
            certificate: b.config.certificate,
            privateKey: open(b).private_key,
          }),
        sleep,
        waits: deps.waits,
      }).catch(async (error: unknown) => {
        const why = error instanceof Error ? error.message : String(error);
        await log(org, by, "add_to_teams", `bot ${b.id} stopped: ${why.slice(0, 200)}`);
        throw error;
      });
      await log(
        org,
        by,
        "add_to_teams",
        `bot ${b.id} ${result.steps.map((s) => `${s.id}=${s.state}`).join(" ")}`,
      );
      if (!(await teams.botById(b.id)))
        throw new ConnectError(
          result.steps.find((x) => x.state === "failed")?.message ??
            "This bot was removed while it was being added.",
        );
      if (result.botAppId && !(await teams.connectBot(b.id, result.botAppId, b.credential)))
        throw new ConnectError(
          "Another bot already uses this app, or this one changed meanwhile. Reload the page.",
        );
      return { steps: result.steps, bot: publicBot((await teams.botById(b.id))!) };
    },

    // Where the administrator signs in to publish the bot. The nonce goes back to the admin
    // app, which keeps it in this browser until the sign-in returns.
    async startPublish(org: string, agentId: string, redirectUri: string) {
      const b = await bot(org, agentId);
      if (!b.botAppId || b.status !== "connected")
        throw new InvalidInput("Finish adding the bot first");
      const { conn } = await microsoftFor(org);
      const { state, nonce } = signState(stateKey, {
        org,
        app: `teams-publish:${agentId}`,
        ceiling: [],
        redirectUri,
      });
      return {
        url: deps.microsoft.delegated.signInUrl(conn, {
          state,
          redirectUri,
          nonce: signInNonce(nonce),
          scope: CATALOG_SCOPE,
        }),
        nonce,
      };
    },

    // The sign-in came back: the signed state names the agent, the nonce must be this
    // browser's, the id token this tenant's administrator. Then the package is published (a
    // new version of it if one is there already) with that sign-in, never kept.
    async finishPublish(
      org: string,
      input: { state: string; nonce: string; params: Record<string, string>; by: string },
    ) {
      const app = stateApp(input.state) ?? "";
      const agentId = PUBLISH_APP.exec(app)?.[1];
      if (!agentId)
        throw new ConnectError("This sign-in did not start here. Start again from API keys.");
      const state = verifyState(stateKey, input.state, { org, app, nonce: input.nonce });
      const { error, error_description: description, code } = input.params;
      if (error)
        throw new ConnectError(
          `Microsoft did not sign you in: ${(description ?? error).split(/\r?\n/)[0].slice(0, 300)}`,
        );
      if (!code) throw new ConnectError("Microsoft did not confirm the sign-in. Start again.");
      const b = await bot(org, agentId);
      // Removed or set up again during the sign-in.
      if (b.status !== "connected")
        throw new ConnectError("This bot is no longer set up. Start again from API keys.");
      const { conn, config } = await microsoftFor(org);
      const { accessToken, claims } = await deps.microsoft.delegated.redeem(conn, {
        code,
        redirectUri: state.redirectUri,
        scope: CATALOG_SCOPE,
      });
      checkSignIn(claims, {
        clientId: config.clientId,
        tenantId: config.tenantId,
        nonce: signInNonce(input.nonce),
      });
      if (!b.botAppId || b.tenantId !== config.tenantId)
        throw new ConnectError("This bot is not set up in this Microsoft organization");
      const version = packageVersion(b);
      const published = await publishPackage(
        {
          token: accessToken,
          externalId: b.teamsAppId,
          version,
          zip: packageZip({
            displayName: b.displayName,
            teamsAppId: b.teamsAppId,
            botAppId: b.botAppId,
            manifestVersion: b.manifestVersion,
            publicUrl: publicUrl(),
          }),
        },
        deps.fetch,
      );
      // Logged before saving: a save refused by a change meanwhile must not hide the publish.
      const who = typeof claims.preferred_username === "string" ? claims.preferred_username : "-";
      await log(
        org,
        input.by,
        "publish_bot",
        `bot ${b.id} catalog ${published.catalogAppId} version ${version}${published.changed ? "" : " (already there)"} signed in as ${who}`,
      );
      await saveConfig(b, {
        ...b.config,
        catalogAppId: published.catalogAppId,
        publishedVersion: version,
      });
      return { agentId, version, changed: published.changed };
    },

    // Teams or people to add the bot to, matching what the admin typed.
    async installTargets(org: string, agentId: string, kind: "team" | "person", query: string) {
      await bot(org, agentId);
      return searchTargets((await microsoftFor(org)).call, kind, query);
    },

    // Adds the published bot to a team, a group chat (its id or a pasted link) or a person.
    async install(org: string, agentId: string, kind: InstallKind, rawId: string, by: string) {
      const b = await bot(org, agentId);
      const catalogAppId = b.config.catalogAppId;
      if (b.status !== "connected" || !catalogAppId)
        throw new InvalidInput("Publish the bot to your Teams apps first");
      const id = kind === "chat" ? chatIdFrom(rawId) : rawId.trim().toLowerCase();
      if (!id || (kind !== "chat" && !GUID.test(id)))
        throw new InvalidInput(
          kind === "chat"
            ? "Paste the group chat's link from Teams (Copy link) or its id"
            : `Pick a ${kind} from the list`,
        );
      if (!(await addsBots(org))) throw new ConnectError(addsOff());
      if (kind === "person" && (await personInstalls(org)) !== true)
        throw new ConnectError(
          'Adding a bot for one person is off. Choose "Also for one person" in step 1 of the Microsoft Teams page and add the permission it shows, or let the person add the bot in Teams.',
        );
      const { call, config } = await microsoftFor(org);
      if (b.tenantId !== config.tenantId)
        throw new ConnectError(
          "This bot was made for another Microsoft organization. Remove it and add it again.",
        );
      const result = await installBot(call, catalogAppId, kind, id);
      await log(org, by, "install_bot", `bot ${b.id} ${kind} ${id} ${result}`);
      return { result };
    },

    async trigger(org: string, agentId: string) {
      const b = await teams.botForAgent(org, agentId);
      const row = await teams.trigger(agentId);
      // Only an agent of this organization (its bot proves that, or nothing is set).
      return b && row ? publicTrigger(row) : null;
    },

    // Proven by its runner first (for Claude: the agent and environment exist under the
    // organization's key, see verifyTrigger).
    async setTrigger(org: string, agentId: string, raw: unknown) {
      const parsed = triggerSchema.safeParse(raw);
      if (!parsed.success) throw new InvalidInput(z.prettifyError(parsed.error));
      const { by, signingSecret, settings: rawSettings, ...input } = parsed.data;
      await bot(org, agentId);
      const runner = deps.runners.get(input.runner);
      if (!runner) throw new InvalidInput(`This deployment offers no runner ${input.runner}`);
      const settings = runner.settings.safeParse(rawSettings);
      if (!settings.success) throw new InvalidInput(z.prettifyError(settings.error));
      publicUrl(); // Every runner answers through this service's MCP address.
      const existing = await teams.trigger(agentId);
      const sealed = signingSecret
        ? key.seal({ secret: signingSecret }, triggerAad(org, agentId))
        : existing?.secret
          ? { ciphertext: existing.secret, keyId: existing.secretKeyId! }
          : null;
      const row = {
        agentId,
        ...input,
        managedAgentId: input.managedAgentId ?? null,
        environmentId: input.environmentId ?? null,
        settings: settings.data,
        secret: sealed?.ciphertext ?? null,
        secretKeyId: sealed?.keyId ?? null,
        allowedTeams: input.allowedTeams.map((t) => t.toLowerCase()),
        allowedPeople: input.allowedPeople.map((p) => p.toLowerCase()),
        updatedBy: by,
      };
      const secret = signingSecret ?? (existing ? openTriggerSecret(key, org, existing) : null);
      await runner.verify({ ...row, updatedAt: "" }, org, secret, "save");
      // Another runner from now on: the old one's live sessions are stopped while it can.
      if (existing && existing.runner !== input.runner)
        await stopSessions(org, (r) => r.agentId === agentId, existing.runner);
      const saved = await teams.setTrigger(org, row);
      if (!saved) return null;
      await log(
        org,
        by,
        "set_trigger",
        `agent ${agentId} ${input.enabled ? "on" : "off"} ${input.runner === CLAUDE_RUNNER ? `agent ${input.managedAgentId}` : `runner ${input.runner}`} budget ${input.budgetCents}c`,
      );
      return this.trigger(org, agentId);
    },

    async runs(org: string, agentId: string, limit: number) {
      return teams.recentRuns(org, agentId, limit);
    },

    // ---- the organization's Claude API key ----

    async claudeKey(org: string) {
      const row = await teams.claudeKey(org);
      return row
        ? ({
            set: true,
            ending: row.config.ending,
            by: row.connectedBy,
            at: row.connectedAt,
          } as const)
        : ({ set: false } as const);
    },

    // Proven before saving: listing the Managed Agents it can use. That list is also what the
    // admin chooses from.
    async setClaudeKey(org: string, rawKey: string, by: string) {
      const apiKey = rawKey.trim();
      if (!/^sk-ant-[A-Za-z0-9_-]{20,}$/.test(apiKey))
        throw new InvalidInput("Paste a Claude API key (it starts with sk-ant-)");
      const agents = await deps.claude(apiKey).listAgents();
      const sealed = key.seal({ api_key: apiKey }, credentialAad(org, CLAUDE_SCOPE));
      await teams.setClaudeKey({
        organizationId: org,
        config: { ending: apiKey.slice(-4) },
        credential: sealed.ciphertext,
        credentialKeyId: sealed.keyId,
        credentialScope: CLAUDE_SCOPE,
        connectedBy: by,
      });
      await log(org, by, "set_claude_key", `ending ${apiKey.slice(-4)}`);
      const saved = await this.claudeKey(org);
      if (!saved.set) throw new Error("The Claude key vanished after saving");
      return { ...saved, agents };
    },

    async claudeAgents(org: string) {
      return (await claudeFor(org)).listAgents();
    },

    async removeClaudeKey(org: string, by: string) {
      await stopSessions(org, () => true, CLAUDE_RUNNER);
      const removed = await teams.removeClaudeKey(org);
      if (removed) await log(org, by, "remove_claude_key", "removed");
      return removed;
    },
  };
}

export type TeamsAdmin = ReturnType<typeof createTeamsAdmin>;

/** During a key rotation, moves every bot credential, Claude key and trigger secret still sealed with the
 *  previous key onto the current one (only over what was read), so the previous key can be
 *  dropped. Connections and the Microsoft sign-in are re-sealed by their own checks. */
export async function resealTeamsCredentials(
  teams: Pick<
    TeamsStore,
    "staleSealed" | "updateBotCredential" | "resealClaudeKey" | "resealTriggerSecret"
  >,
  key: CredentialKey,
): Promise<{ bots: number; claudeKeys: number; triggerSecrets: number }> {
  if (!key.previousIds.length) return { bots: 0, claudeKeys: 0, triggerSecrets: 0 };
  const stale = await teams.staleSealed(key.previousIds);
  let bots = 0;
  for (const b of stale.bots) {
    const aad = botAad(b.organizationId, b.id);
    const secret = key.open({ ciphertext: b.credential, keyId: b.credentialKeyId }, aad);
    const sealed = key.seal(secret, aad);
    if (
      await teams.updateBotCredential(b.id, b.config, sealed.ciphertext, sealed.keyId, b.credential)
    )
      bots++;
  }
  let claudeKeys = 0;
  for (const c of stale.claudeKeys) {
    const aad = credentialAad(c.organizationId, CLAUDE_SCOPE);
    const secret = key.open({ ciphertext: c.credential, keyId: c.credentialKeyId }, aad);
    const sealed = key.seal(secret, aad);
    if (await teams.resealClaudeKey(c.id, sealed.ciphertext, sealed.keyId, c.credential))
      claudeKeys++;
  }
  let triggerSecrets = 0;
  for (const t of stale.triggers) {
    const aad = triggerAad(t.organizationId, t.agentId);
    const sealed = key.seal(key.open({ ciphertext: t.secret, keyId: t.secretKeyId }, aad), aad);
    if (await teams.resealTriggerSecret(t.agentId, sealed.ciphertext, sealed.keyId, t.secret))
      triggerSecrets++;
  }
  return { bots, claudeKeys, triggerSecrets };
}
