import { createHmac } from "node:crypto";
import { z } from "zod";
import { brand } from "../../../connector-kit/src/branding.js";
import { credentialAad, type CredentialKey } from "../secrets.js";
import { InvalidInput } from "../service.js";
import { ConnectError, type Fetch } from "../../../connector-kit/src/types.js";
import {
  agentMcpUrl,
  claudeClient,
  verifyTrigger,
  type ClaudeFactory,
  type SessionState,
} from "./claude.js";
import type { RunRow, TeamsStore, TriggerRow } from "./store.js";

// What answers a Teams message once the core decided it should be answered (who may ask, reply
// mode, the hourly, daily and concurrent caps, the reply token and teams_reply stay in
// runtime.ts whatever the runner).

export const CLAUDE_RUNNER = "claude-managed-agents";

export type RunnerContext = {
  teams: Pick<TeamsStore, "claudeKey">;
  key: CredentialKey;
  // This service's public address (see createTeamsAdmin), null when it has none.
  publicUrl: string | null;
};

export type RunInput = {
  org: string;
  agentId: string;
  runId: string;
  botId: string;
  requestedBy: string;
  // The prompt, the thread quoted as untrusted text, with the reply token in it.
  text: string;
  replyToken: string;
  budgetCents: number;
  trigger: TriggerRow;
  // The trigger's signing secret, opened; null when it has none.
  secret: string | null;
};

export interface TriggerRunner {
  id: string;
  // The trigger body's `settings` for this runner, stored as they are (so never a secret).
  settings: z.ZodType<Record<string, unknown>>;
  // Why the organization cannot run it now (recorded as the message's reason), or null.
  unready(org: string): Promise<string | null>;
  // Proves the trigger: when saved, and again before each answer (`when`), since what it
  // points at can change elsewhere. Refuses with ConnectError or InvalidInput.
  verify(
    trigger: TriggerRow,
    org: string,
    secret: string | null,
    when: "save" | "answer",
  ): Promise<void>;
  // Starts answering; the id later calls name the session by.
  start(input: RunInput): Promise<string>;
  // The next message in the same session (the same person in the same thread).
  continue(sessionId: string, input: RunInput): Promise<void>;
  // Where the session stands, or why it is gone.
  state(run: RunRow, org: string): Promise<SessionState | { gone: string }>;
  stop(sessionId: string, org: string): Promise<void>;
}

export type RunnerFactory = (ctx: RunnerContext) => TriggerRunner;

export const triggerAad = (org: string, agentId: string) =>
  credentialAad(org, `teams-trigger:${agentId}`);

/** The trigger's signing secret, opened, or null when it has none. */
export function openTriggerSecret(key: CredentialKey, org: string, trigger: TriggerRow) {
  if (!trigger.secret || !trigger.secretKeyId) return null;
  return String(
    key.open(
      { ciphertext: trigger.secret, keyId: trigger.secretKeyId },
      triggerAad(org, trigger.agentId),
    ).secret,
  );
}

/** Claude Managed Agents with the organization's own Claude key (providers row anthropic). Its
 *  settings are the trigger's managedAgentId, environmentId and vaultIds. */
export const claudeManagedAgents =
  (claude: ClaudeFactory = claudeClient): RunnerFactory =>
  (ctx) => {
    async function claudeFor(org: string) {
      const row = await ctx.teams.claudeKey(org);
      if (!row) return null;
      const { api_key } = ctx.key.open(
        { ciphertext: row.credential, keyId: row.credentialKeyId },
        credentialAad(org, "provider:anthropic"),
      );
      return claude(api_key);
    }
    async function need(org: string) {
      const c = await claudeFor(org);
      if (!c) throw new ConnectError("Set your organization's Claude API key first");
      return c;
    }
    return {
      id: CLAUDE_RUNNER,
      settings: z.strictObject({}),
      async unready(org) {
        return (await ctx.teams.claudeKey(org)) ? null : "organization has no Claude API key";
      },
      async verify(trigger, org, _secret, when) {
        if (!trigger.managedAgentId || !trigger.environmentId)
          throw new InvalidInput("Choose the Claude agent and environment");
        if (!ctx.publicUrl) throw new ConnectError("This deployment has no public address");
        const c = await need(org);
        await verifyTrigger(c, ctx.publicUrl, trigger.agentId, {
          managedAgentId: trigger.managedAgentId,
          vaultIds: trigger.vaultIds,
        });
        if (when === "save" && !(await c.environmentExists(trigger.environmentId)))
          throw new ConnectError(
            "Claude has no environment with that id under your organization's key",
          );
      },
      async start(input) {
        return (await need(input.org)).createSession({
          agentId: input.trigger.managedAgentId!,
          environmentId: input.trigger.environmentId!,
          vaultIds: input.trigger.vaultIds,
          budgetCents: input.budgetCents,
          metadata: {
            source: `${brand().resourcePrefix}-teams`,
            agent: input.agentId,
            bot: input.botId,
            run: input.runId,
            requested_by: input.requestedBy,
          },
          text: input.text,
        });
      },
      async continue(sessionId, input) {
        await (await need(input.org)).sendMessage(sessionId, input.text, input.budgetCents);
      },
      async state(run, org) {
        const c = await claudeFor(org);
        if (!c) return { gone: "Claude API key removed" };
        return (await c.session(run.sessionId!)) ?? { gone: "The Claude session is gone" };
      },
      // Without the key there is nothing left to stop it with.
      async stop(sessionId, org) {
        await (await claudeFor(org))?.interrupt(sessionId);
      },
    };
  };

export const SIGNATURE_HEADER = "x-agent-connect-signature";

/** `t=<unix seconds>,v1=<hex HMAC-SHA256 of "<t>.<body>">`; receivers reject an old `t`. */
export function signature(secret: string, body: string, t = Math.floor(Date.now() / 1000)) {
  return `t=${t},v1=${createHmac("sha256", secret).update(`${t}.${body}`).digest("hex")}`;
}

/** A signed POST of the message and its context to the trigger's `url`, for any agent
 *  framework. The receiver answers through MCP teams_reply with the reply token and this
 *  agent's own key, as a Claude agent does. It reports no cost; the trigger's caps still
 *  apply. A run is answered once a reply arrives, and stopped after the turn timeout. */
export const signedWebhook =
  (fetchImpl: Fetch = fetch): RunnerFactory =>
  (ctx) => {
    async function post(trigger: TriggerRow, secret: string | null, payload: unknown) {
      if (!secret) throw new InvalidInput("Give the webhook a signing secret");
      const body = JSON.stringify(payload);
      const res = await fetchImpl(String(trigger.settings.url), {
        method: "POST",
        headers: {
          "content-type": "application/json",
          [SIGNATURE_HEADER]: signature(secret, body),
        },
        body,
        redirect: "error",
        signal: AbortSignal.timeout(10_000),
      });
      if (!res.ok) throw new ConnectError(`The webhook answered ${res.status}`);
    }
    const message = (input: RunInput) => ({
      type: "message",
      run: input.runId,
      agent: input.agentId,
      bot: input.botId,
      requestedBy: input.requestedBy,
      text: input.text,
      replyToken: input.replyToken,
      mcpUrl: ctx.publicUrl ? agentMcpUrl(ctx.publicUrl, input.agentId) : null,
    });
    return {
      id: "webhook",
      settings: z.strictObject({ url: z.url({ protocol: /^https$/ }) }),
      unready: async () => null,
      async verify(trigger, _org, secret, when) {
        if (when === "save") await post(trigger, secret, { type: "ping", agent: trigger.agentId });
      },
      async start(input) {
        await post(input.trigger, input.secret, message(input));
        return `webhook:${input.runId}`;
      },
      async continue(_sessionId, input) {
        await post(input.trigger, input.secret, message(input));
      },
      async state(run) {
        return { status: run.replyCount > 0 ? "idle" : "running", costCents: 0, atBudget: false };
      },
      async stop() {},
    };
  };
