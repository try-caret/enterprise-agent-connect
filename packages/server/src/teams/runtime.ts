import { effectiveCapabilities, type Connector } from "../../../connector-kit/src/connector.js";
import { credentialAad, type CredentialKey } from "../secrets.js";
import type { Store } from "../store.js";
import { ConnectError, type VendorResponse } from "../../../connector-kit/src/types.js";
import {
  decide,
  guid,
  installChange,
  placeOf,
  plainText,
  promptText,
  newReplyToken,
  hashReplyToken,
  type Activity,
  type Place,
  type ThreadMessage,
} from "./activity.js";
import {
  approvalCard,
  cardAttachment,
  invokeCard,
  invokeError,
  invokeMessage,
  MAX_COMMENT,
  type ApprovalFact,
  type ApprovalRow,
} from "./approval.js";
import { validServiceUrl, type BotClient, type BotCredential } from "./bot.js";
import { drainer } from "../queue.js";
import { openTriggerSecret, type TriggerRunner } from "./runner.js";
import { Unauthorized, type BotAuth } from "./jwt.js";
import type { BotRow, ConversationRow, EventRow, RunRow, TeamsStore, TriggerRow } from "./store.js";

// Teams' requests to the agents' bots, the queue they feed, and the sessions that answer them
// (Claude Managed Agents or another runner, runner.ts). docs/teams.md 6.

const SESSION_MAX_AGE_MS = 12 * 3600_000;
const SESSION_MAX_TURNS = 20;
const THREAD_CONTEXT = 20;
const FAILED_NOTICE = "Sorry, I couldn't finish this. Your admin can see what happened.";
const BUDGET_NOTICE = "I stopped because this answer reached its spending limit.";
const NO_REPLY_NOTICE = "I could not answer this one. Ask your admin to check my Teams settings.";
// A turn still running after this is stopped: it holds a slot and the thread's later messages.
const STUCK_MS = 30 * 60_000;
// How long a passed trigger check is trusted before the next answer checks again.
const VERIFY_TTL_MS = 60_000;
// Teams activities are small; anything bigger is not Teams.
const MAX_BODY = 256 * 1024;

// What the admin sees when Microsoft's request for a bot was refused.
const REFUSED: Record<string, string> = {
  "wrong recipient": "A message for another bot arrived at this bot's address.",
  "wrong tenant": "Microsoft sent a message from another organization.",
  "not a Teams service address": "A message came from an address that is not Microsoft Teams.",
  "serviceUrl does not match the token": "A message did not match Microsoft's signature.",
  "key not endorsed for this channel": "A message came from a channel other than Teams.",
};

/** The body as text, or null once it passes `max` bytes (read no further). */
export async function readCapped(req: Request, max = MAX_BODY): Promise<string | null> {
  return (await readCappedBytes(req, max))?.toString("utf8") ?? null;
}

/** The body as sent (a signature covers these bytes), or null once it passes `max`. */
export async function readCappedBytes(req: Request, max: number): Promise<Buffer | null> {
  if (Number(req.headers.get("content-length") ?? 0) > max) return null;
  if (!req.body) return Buffer.alloc(0);
  const chunks: Uint8Array[] = [];
  let size = 0;
  for await (const chunk of req.body as unknown as AsyncIterable<Uint8Array>) {
    size += chunk.length;
    if (size > max) return null;
    chunks.push(chunk);
  }
  return Buffer.concat(chunks);
}

export const botAad = (org: string, botId: string) => credentialAad(org, `teams-bot:${botId}`);
const eventAad = (botId: string, activityId: string) =>
  `agent-connect:teams-event:${botId}:${activityId}`;

/** The bot's sign-in: its app id, tenant, certificate and private key. */
export function botCredential(key: CredentialKey, bot: BotRow): BotCredential {
  if (!bot.botAppId) throw new Error("The bot has no app id yet");
  const secret = key.open(
    { ciphertext: bot.credential, keyId: bot.credentialKeyId },
    botAad(bot.organizationId, bot.id),
  );
  if (!bot.config.certificate || !secret.private_key) throw new Error("The bot has no certificate");
  return {
    botAppId: bot.botAppId,
    tenantId: bot.tenantId,
    certificate: bot.config.certificate,
    privateKey: secret.private_key,
  };
}

const reply = (status: number, body?: unknown) =>
  new Response(body === undefined ? null : JSON.stringify(body), {
    status,
    headers: body === undefined ? {} : { "content-type": "application/json" },
  });

export type TeamsRuntimeDeps = {
  store: Store;
  teams: TeamsStore;
  key: CredentialKey;
  auth: BotAuth;
  bots: BotClient;
  // By id: what answers each trigger.
  runners: Map<string, TriggerRunner>;
  // Null when the deployment leaves Teams out: then no organization can connect it.
  connector: Connector | null;
  // This service's public address (see createTeamsAdmin); the trigger's agent must use it.
  publicUrl: string | null;
};

export function createTeamsRuntime(deps: TeamsRuntimeDeps) {
  const { teams, key, bots } = deps;

  async function post(
    bot: BotRow,
    serviceUrl: string,
    conversation: string,
    activity: Record<string, unknown>,
    replyTo?: string,
  ) {
    const path = `/v3/conversations/${encodeURIComponent(conversation)}/activities${replyTo ? `/${encodeURIComponent(replyTo)}` : ""}`;
    const res = await bots.connector(botCredential(key, bot), serviceUrl, {
      method: "POST",
      path,
      body: replyTo ? { ...activity, replyToId: replyTo } : activity,
    });
    // The bot was removed or blocked in a chat [4]. A team's channel can refuse a post for
    // other reasons (moderation), so teams are only removed by their uninstall event.
    if (res.status === 403) {
      const chat = (await teams.conversations(bot.id)).find(
        (c) => c.conversationId === conversation && c.kind !== "team",
      );
      if (chat) await teams.removeConversation(bot.id, chat.conversationId);
    }
    return res;
  }

  // ---- inbound ----

  /** POST /teams/bots/:botAppId/messages. Proven by Microsoft's token for this bot, queued,
   *  answered at once. */
  async function inbound(botAppId: string, req: Request): Promise<Response> {
    if (req.headers.has("origin")) return reply(403);
    let signed = false;
    try {
      // The token first: nothing from the body is read without a valid Bot Framework token.
      const matches = await deps.auth.token(req.headers.get("authorization"), botAppId);
      signed = true;
      const raw = await readCapped(req);
      if (raw === null) return reply(413);
      let activity: Activity;
      try {
        activity = JSON.parse(raw) as Activity;
      } catch {
        return reply(400);
      }
      const serviceUrl = activity.serviceUrl ?? "";
      matches({ serviceUrl, channelId: activity.channelId ?? "" });
      const bot = await teams.botByAppId(botAppId);
      if (!bot || bot.status === "awaiting_setup") throw new Unauthorized("unknown bot");
      if (activity.recipient?.id !== `28:${botAppId}`) throw new Unauthorized("wrong recipient");
      if (activity.channelData?.tenant?.id?.toLowerCase() !== bot.tenantId)
        throw new Unauthorized("wrong tenant");
      if (!validServiceUrl(serviceUrl)) throw new Unauthorized("not a Teams service address");
      // A card button: answered in this response, never queued, so it starts no run whatever
      // the trigger says.
      if (activity.type === "invoke" && activity.name === "adaptiveCard/action")
        return reply(200, await cardAction(bot, activity));
      if (!activity.id) return reply(200);
      const place = placeOf(activity);
      const sealed = key.seal(
        { activity: JSON.stringify(activity) },
        eventAad(bot.id, activity.id),
      );
      const fresh = await teams.insertEvent({
        botId: bot.id,
        activityId: activity.id,
        type: activity.type ?? "unknown",
        conversationId: place?.conversationId ?? activity.conversation?.id ?? "",
        threadId: place?.threadId ?? null,
        fromAadId: activity.from?.aadObjectId ?? null,
        payload: sealed.ciphertext,
        payloadKeyId: sealed.keyId,
      });
      if (fresh) kick();
      return reply(200);
    } catch (error) {
      if (!(error instanceof Unauthorized)) throw error;
      console.warn(`[teams] refused a request for bot ${botAppId}: ${error.message}`);
      // Only a request with a valid token is shown to the admin; anyone can send the rest.
      const why = REFUSED[error.message];
      if (signed && why) await teams.refused(botAppId, why);
      return reply(403);
    }
  }

  // ---- approval cards (approval.ts) ----

  /** An Action.Execute click on an approval card, after the token, bot, recipient, tenant and
   *  service address checks (the request is authenticated by a Bot Framework token; its body is
   *  not signed). Counts only from the approver (matched on the Entra object id Teams sends), in
   *  the team or chat the card was posted in, once, while pending and unexpired; every click is
   *  audited, accepted or not. */
  async function cardAction(bot: BotRow, a: Activity) {
    const value = (a.value ?? {}) as {
      action?: { type?: string; verb?: string; data?: Record<string, unknown> };
    };
    const verb = value.action?.verb;
    const data = value.action?.data ?? {};
    const id = guid(data.approval);
    const clicker = guid(a.from?.aadObjectId);
    // The team (for a channel) or the chat the click came from.
    const conversationId = placeOf(a)?.conversationId ?? null;
    const audit = async (decision: "allowed" | "denied", reason: string) => {
      const state = await teams.agentState(bot.agentId);
      await deps.store.audit({
        organizationId: bot.organizationId,
        agentId: bot.agentId,
        agentName: state?.agentName ?? "",
        connectorId: "teams",
        tool: "teams:approval",
        target: id,
        decision,
        reason: reason.slice(0, 500),
        vendorStatus: null,
      });
    };
    if (
      value.action?.type !== "Action.Execute" ||
      (verb !== "approve" && verb !== "reject") ||
      !id
    ) {
      await audit("denied", `unknown card action clicked by ${clicker ?? "unknown"}`);
      return invokeError("This card action is not one this agent knows");
    }
    const approval = conversationId ? await teams.approvalForBot(bot.id, id, conversationId) : null;
    if (!approval) {
      await audit(
        "denied",
        `${verb} clicked by ${clicker ?? "unknown"} in ${conversationId ?? "an unknown place"}: no such approval for this bot there`,
      );
      return invokeMessage("This request is not known to this agent.");
    }
    if (!clicker || clicker !== approval.approverAadId) {
      await audit("denied", `${verb} clicked by ${clicker ?? "unknown"}, not the approver`);
      return invokeMessage(`Only ${approval.approverName} can answer this request.`);
    }
    if (approval.status !== "pending") {
      await audit("denied", `${verb} clicked again: already ${approval.status}`);
      return invokeCard(approvalCard(approval));
    }
    if (approval.expired) {
      await audit("denied", `${verb} clicked after it expired`);
      await updateCard(bot, approval);
      return invokeCard(approvalCard(approval));
    }
    const comment =
      typeof data.comment === "string" ? data.comment.trim().slice(0, MAX_COMMENT) || null : null;
    const decided = await teams.decideApproval(
      bot.id,
      approval.id,
      approval.conversationId,
      {
        status: verb === "approve" ? "approved" : "rejected",
        byAadId: clicker,
        byName: (a.from?.name?.trim() || approval.approverName).slice(0, 200),
      },
      comment,
    );
    if (!decided) {
      // Answered or expired in the meantime.
      const now =
        (await teams.approvalForBot(bot.id, approval.id, approval.conversationId)) ?? approval;
      await audit("denied", `${verb} clicked again: already ${now.status}`);
      return invokeCard(approvalCard(now));
    }
    await audit("allowed", `${decided.status} by ${clicker}${comment ? " with a comment" : ""}`);
    await updateCard(bot, decided);
    return invokeCard(approvalCard(decided));
  }

  // The answer in the invoke changes only the clicker's view; editing the bot's message changes
  // everyone's. A failed edit is logged, never undoes the decision.
  async function updateCard(bot: BotRow, approval: ApprovalRow) {
    if (!approval.threadId || !approval.activityId) return;
    try {
      const res = await bots.connector(botCredential(key, bot), approval.serviceUrl, {
        method: "PUT",
        path: `/v3/conversations/${encodeURIComponent(approval.threadId)}/activities/${encodeURIComponent(approval.activityId)}`,
        body: {
          type: "message",
          id: approval.activityId,
          attachments: [cardAttachment(approvalCard(approval))],
        },
      });
      if (res.status < 200 || res.status >= 300)
        console.error(`[teams] editing approval card ${approval.id}: ${res.status}`);
    } catch (error) {
      console.error(`[teams] editing approval card ${approval.id}`, error);
    }
  }

  // ---- the queue ----

  const queue = drainer("teams", async () => {
    const event = await teams.claimEvent();
    if (!event) return false;
    try {
      await handle(event);
    } catch (error) {
      console.error(`[teams] event ${event.id} failed`, error);
      await teams.finishEvent(event.id, "failed", "Internal error");
    }
    return true;
  });
  const kick = queue.kick;

  async function handle(event: EventRow) {
    const bot = await teams.botById(event.botId);
    if (!bot || !event.payload || !event.payloadKeyId)
      return teams.finishEvent(event.id, "failed", "bot or payload gone");
    const activity = JSON.parse(
      key.open(
        { ciphertext: event.payload, keyId: event.payloadKeyId },
        eventAad(bot.id, event.activityId),
      ).activity,
    ) as Activity;
    const place = placeOf(activity);
    if (!place) return teams.finishEvent(event.id, "ignored", "unknown conversation type");
    const installed = installChange(activity, `28:${bot.botAppId}`);
    if (installed === false) {
      await teams.removeConversation(bot.id, place.conversationId);
      return teams.finishEvent(event.id, "done", "removed");
    }
    await remember(bot, activity, place);
    if (activity.type !== "message")
      return teams.finishEvent(event.id, "done", installed ? "installed" : (activity.type ?? null));
    return respond(bot, event, activity, place);
  }

  async function remember(bot: BotRow, a: Activity, place: Place) {
    await teams.upsertConversation({
      botId: bot.id,
      conversationId: place.conversationId,
      kind: place.kind,
      serviceUrl: a.serviceUrl!,
      teamId: place.teamId,
      teamGroupId: place.teamGroupId,
      name: place.name,
      userAadId: place.kind === "personal" ? (a.from?.aadObjectId ?? null) : null,
    });
  }

  async function respond(bot: BotRow, event: EventRow, a: Activity, place: Place) {
    const [state, trigger] = await Promise.all([
      teams.agentState(bot.agentId),
      teams.trigger(bot.agentId),
    ]);
    const runner = trigger ? deps.runners.get(trigger.runner) : undefined;
    const runnerUnready = !trigger?.enabled
      ? null
      : runner
        ? await runner.unready(bot.organizationId)
        : `runner ${trigger.runner} is not offered by this deployment`;
    const capabilities =
      deps.connector && state?.grant && state.connection?.status === "connected"
        ? effectiveCapabilities(
            deps.connector,
            state.grant.capabilities,
            state.connection.capabilities,
          ).map((c) => c.id)
        : [];
    // Where this team is installed, its group id is known even if the message lacks it.
    if (place.kind === "team" && !place.teamGroupId)
      place.teamGroupId =
        (await teams.conversations(bot.id)).find((c) => c.conversationId === place.conversationId)
          ?.teamGroupId ?? null;
    const decision = decide({
      activity: a,
      place,
      botAppId: bot.botAppId!,
      tenantId: bot.tenantId,
      agentActive: state?.agentStatus === "active",
      capabilities,
      limits: state?.grant?.limits ?? {},
      trigger,
      runnerUnready,
      counts: await teams.runCounts(bot.agentId, bot.id),
    });
    if (decision.action === "ignore")
      return teams.finishEvent(event.id, "ignored", decision.reason);
    if (decision.action === "limit") {
      await limitNotice(bot, a.serviceUrl!, place, place.threadId, decision.notice);
      return teams.finishEvent(event.id, "ignored", decision.reason);
    }
    // Not named in the allowed people: only a member of the organization, never a guest.
    if (decision.checkMember && !(await member(bot, a, place)))
      return teams.finishEvent(event.id, "ignored", "sender is not a member of the organization");

    const token = newReplyToken();
    const latest = await teams.latestRun(bot.id, place.threadId);
    // A session holds what it found for the person who asked; it is not reused for anyone else,
    // nor by another runner than the one that started it.
    const reusable =
      latest?.state === "idle" &&
      latest.runner === trigger!.runner &&
      latest.requestedBy === decision.sender &&
      latest.sessionId &&
      latest.turns < SESSION_MAX_TURNS &&
      Date.now() - Date.parse(latest.startedAt) < SESSION_MAX_AGE_MS;
    let run: RunRow | null = null;
    if (reusable) {
      if (!(await teams.continueRun(latest.id, a.id!, hashReplyToken(token))))
        return teams.finishEvent(event.id, "held", "conversation busy");
      run = latest;
    } else {
      run = await teams.startRun({
        agentId: bot.agentId,
        botId: bot.id,
        conversationId: place.conversationId,
        threadId: place.threadId,
        triggerActivityId: a.id!,
        replyTokenHash: hashReplyToken(token),
        requestedBy: decision.sender,
        runner: trigger!.runner,
      });
      if (!run) return teams.finishEvent(event.id, "held", "conversation busy");
    }

    try {
      const secret = openTriggerSecret(key, bot.organizationId, trigger!);
      await verified(runner!, bot, state?.agentName ?? "", trigger!, secret);
      await deps.store.audit({
        organizationId: bot.organizationId,
        agentId: bot.agentId,
        agentName: state?.agentName ?? "",
        connectorId: "teams",
        tool: "teams:run",
        target: place.threadId,
        decision: "allowed",
        reason: `run ${run.id} ${reusable ? "continued" : "started"} for ${decision.sender}`,
        vendorStatus: null,
      });
      await post(bot, a.serviceUrl!, place.threadId, { type: "typing" });
      const channelName = a.channelData?.channel?.name ?? null;
      const context = reusable ? null : await threadContext(bot, place, a.id!);
      if (context?.note) await teams.updateRun(run.id, { note: context.note });
      const text = promptText({
        botName: bot.displayName,
        sender: { name: a.from?.name ?? "Someone", aadObjectId: decision.sender },
        place,
        channelName,
        thread: context?.messages ?? null,
        message: plainText(a.text ?? ""),
        replyToken: token,
        continuing: Boolean(reusable),
      });
      const input = {
        org: bot.organizationId,
        agentId: bot.agentId,
        runId: run.id,
        botId: bot.id,
        requestedBy: decision.sender,
        text,
        replyToken: token,
        budgetCents: trigger!.budgetCents,
        trigger: trigger!,
        secret,
      };
      if (reusable) {
        await runner!.continue(latest.sessionId!, input);
        await teams.updateRun(run.id, { state: "running" });
      } else {
        const sessionId = await runner!.start(input);
        await teams.updateRun(run.id, { state: "running", sessionId });
      }
      await teams.finishEvent(event.id, "run", null);
    } catch (error) {
      console.error(`[teams] run ${run.id} failed to start`, error);
      await teams.updateRun(run.id, {
        state: "failed",
        error: (error instanceof Error ? error.message : String(error)).slice(0, 500),
      });
      await notice(bot, a.serviceUrl!, place.threadId, FAILED_NOTICE, a.id).catch((e) =>
        console.error(`[teams] notice for run ${run.id}`, e),
      );
      await teams.finishEvent(event.id, "failed", "could not start the response");
    }
  }

  // The trigger, checked again before each answer (e.g. the Claude agent and vaults, see
  // verifyTrigger). A refusal fails the run, which the admin sees with its reason, and is
  // audited as denied.
  const verifiedAt = new Map<string, number>();
  async function verified(
    runner: TriggerRunner,
    bot: BotRow,
    agentName: string,
    trigger: TriggerRow,
    secret: string | null,
  ) {
    const k = JSON.stringify([
      bot.agentId,
      trigger.runner,
      trigger.settings,
      trigger.managedAgentId,
      trigger.vaultIds,
    ]);
    if (Date.now() - (verifiedAt.get(k) ?? 0) < VERIFY_TTL_MS) return;
    try {
      await runner.verify(trigger, bot.organizationId, secret, "answer");
    } catch (error) {
      if (error instanceof ConnectError)
        await deps.store.audit({
          organizationId: bot.organizationId,
          agentId: bot.agentId,
          agentName,
          connectorId: "teams",
          tool: "teams:run",
          target: trigger.managedAgentId ?? trigger.runner,
          decision: "denied",
          reason: `trigger check failed: ${error.message}`.slice(0, 500),
          vendorStatus: null,
        });
      throw error;
    }
    verifiedAt.set(k, Date.now());
  }

  // A limit reached, said at most once an hour in the conversation: every message past it
  // would otherwise get one.
  async function limitNotice(
    bot: BotRow,
    serviceUrl: string,
    place: { conversationId: string },
    threadId: string,
    text: string,
  ) {
    if (await teams.claimLimitNotice(bot.id, place.conversationId))
      await notice(bot, serviceUrl, threadId, text);
  }

  // What became of one response (failed to start, budget, no reply, stuck): once per run, so
  // always said.
  async function notice(
    bot: BotRow,
    serviceUrl: string,
    threadId: string,
    text: string,
    replyTo?: string,
  ) {
    await post(bot, serviceUrl, threadId, { type: "message", text }, replyTo);
  }

  // The sender as Teams knows them: in the bot's own tenant, and a member, not a guest.
  async function member(bot: BotRow, a: Activity, place: Place): Promise<boolean> {
    if (!a.from?.id) return false;
    const res = await bots.connector(botCredential(key, bot), a.serviceUrl!, {
      method: "GET",
      path: `/v3/conversations/${encodeURIComponent(place.conversationId)}/members/${encodeURIComponent(a.from.id)}`,
    });
    const m = res.data as { tenantId?: string; userRole?: string } | null;
    return (
      res.status === 200 && m?.tenantId?.toLowerCase() === bot.tenantId && m.userRole === "user"
    );
  }

  // Earlier messages in the thread or group chat, through Graph with the bot's own token.
  // Direct chats cannot be read this way; their earlier turns are in the reused session.
  async function threadContext(
    bot: BotRow,
    place: Place,
    activityId: string,
  ): Promise<{ messages: ThreadMessage[] | null; note?: string } | null> {
    if (place.kind === "personal") return null;
    const blocked = (status: number) =>
      status === 403
        ? {
            messages: null,
            note: "Could not read earlier messages: your organization blocks apps from reading them.",
          }
        : { messages: null };
    if (place.kind === "team") {
      if (!place.teamGroupId || !place.channelId) return null;
      const root = /;messageid=(\d+)/.exec(place.threadId)?.[1];
      if (!root || root === activityId) return null;
      const base = `/v1.0/teams/${place.teamGroupId}/channels/${encodeURIComponent(place.channelId)}/messages/${root}`;
      const [first, replies] = await Promise.all([
        bots.graph(botCredential(key, bot), base),
        bots.graph(botCredential(key, bot), `${base}/replies?$top=${THREAD_CONTEXT}`),
      ]);
      if (first.status !== 200) return blocked(first.status);
      if (replies.status !== 200) return blocked(replies.status);
      const list = [
        first.data as GraphMessage,
        ...((replies.data as { value?: GraphMessage[] }).value ?? []).reverse(),
      ];
      return {
        messages: simplify(list.filter((m) => m.id !== activityId)).slice(-THREAD_CONTEXT),
      };
    }
    const path = `/v1.0/chats/${encodeURIComponent(place.conversationId)}/messages?$top=${THREAD_CONTEXT + 1}&$orderby=createdDateTime desc`;
    const res = await bots.graph(botCredential(key, bot), path);
    if (res.status !== 200) return blocked(res.status);
    const list = ((res.data as { value?: GraphMessage[] }).value ?? []).reverse();
    return { messages: simplify(list.filter((m) => m.id !== activityId)).slice(-THREAD_CONTEXT) };
  }

  // ---- watching the sessions ----

  /** One pass: finished sessions are recorded; messages held while a thread was busy go back
   *  into the queue once it is not. */
  async function watch() {
    const stale = await teams.failStaleStarts();
    if (stale) console.error(`[teams] ${stale} run(s) failed after not starting`);
    for (const run of await teams.runningRuns()) {
      try {
        const bot = await teams.botById(run.botId);
        if (!bot) continue;
        const runner = deps.runners.get(run.runner);
        // The session itself is left running: no runner here can stop it.
        if (!runner) {
          await teams.updateRun(run.id, {
            state: "failed",
            error: `Runner ${run.runner} is not offered by this deployment`,
          });
          continue;
        }
        if (!(await teams.trigger(run.agentId))) {
          await teams.updateRun(run.id, { state: "failed", error: "Its trigger was removed" });
          continue;
        }
        const s = await runner.state(run, bot.organizationId);
        if ("gone" in s) {
          await teams.updateRun(run.id, { state: "failed", error: s.gone });
          continue;
        }
        const costCents = s.costCents ?? undefined;
        const tell = (text: string) => runNotice(bot, run, text);
        if (s.status === "idle" && s.atBudget) {
          await teams.updateRun(run.id, {
            state: "failed",
            error: "Stopped at its spending limit",
            costCents,
          });
          await tell(BUDGET_NOTICE);
        } else if (s.status === "idle" && run.replyCount === 0) {
          await teams.updateRun(run.id, {
            state: "no_reply",
            error: "Finished without replying in Teams",
            costCents,
          });
          await tell(NO_REPLY_NOTICE);
        } else if (s.status === "idle") await teams.updateRun(run.id, { state: "idle", costCents });
        else if (s.status === "terminated")
          await teams.updateRun(run.id, { state: "ended", costCents });
        else if (Date.now() - Date.parse(run.replyTokenAt) > STUCK_MS) {
          await runner
            .stop(run.sessionId!, bot.organizationId)
            .catch((e) => console.error(`[teams] interrupting run ${run.id}`, e));
          await teams.updateRun(run.id, {
            state: "failed",
            error: "Stopped after 30 minutes without finishing",
            costCents,
          });
          await tell(FAILED_NOTICE);
        }
      } catch (error) {
        console.error(`[teams] watching run ${run.id}`, error);
      }
    }
    if (await teams.requeueHeld()) kick();
  }

  async function runNotice(bot: BotRow, run: RunRow, text: string) {
    const place = (await teams.conversations(bot.id)).find(
      (c) => c.conversationId === run.conversationId,
    );
    if (place)
      await notice(bot, place.serviceUrl, run.threadId, text, run.triggerActivityId).catch((e) =>
        console.error(`[teams] notice for run ${run.id}`, e),
      );
  }

  let timers: NodeJS.Timeout[] = [];
  return {
    inbound,
    kick,
    watch,
    /** Drains the queue now and then (events from a worker that stopped) and watches sessions
     *  every few seconds. */
    start() {
      queue.start();
      timers = [
        setInterval(() => {
          kick();
          watch().catch((error) => console.error("[teams] watch", error));
        }, 5000),
        setInterval(
          () => teams.pruneEvents().catch((error) => console.error("[teams] prune", error)),
          3600_000,
        ),
      ];
      for (const t of timers) t.unref();
    },
    /** Stops claiming events and waits for the one in hand, so a rollout neither strands it
     *  in 'processing' nor answers it twice. */
    async stop() {
      for (const t of timers) clearInterval(t);
      await queue.stop();
    },
    /** What an agent's Teams tools reach: only its own bot. */
    access(bot: BotRow): TeamsAccess {
      return {
        places: () => teams.conversations(bot.id),
        graph: (path) => bots.graph(botCredential(key, bot), path),
        listChannels: (place) =>
          bots.connector(botCredential(key, bot), place.serviceUrl, {
            method: "GET",
            path: `/v3/teams/${encodeURIComponent(place.teamId!)}/conversations`,
          }),
        post: (serviceUrl, conversation, text, replyTo) =>
          post(
            bot,
            serviceUrl,
            conversation,
            { type: "message", text, textFormat: "markdown" },
            replyTo,
          ),
        newThread: (place, channelId, text) =>
          bots.connector(botCredential(key, bot), place.serviceUrl, {
            method: "POST",
            path: "/v3/conversations",
            body: {
              isGroup: true,
              channelData: { channel: { id: channelId }, tenant: { id: bot.tenantId } },
              activity: { type: "message", text, textFormat: "markdown" },
            },
          }),
        postCard: (serviceUrl, conversation, card) =>
          post(bot, serviceUrl, conversation, {
            type: "message",
            attachments: [cardAttachment(card)],
          }),
        newThreadCard: (place, channelId, card) =>
          bots.connector(botCredential(key, bot), place.serviceUrl, {
            method: "POST",
            path: "/v3/conversations",
            body: {
              isGroup: true,
              channelData: { channel: { id: channelId }, tenant: { id: bot.tenantId } },
              activity: { type: "message", attachments: [cardAttachment(card)] },
            },
          }),
        createApproval: (input) =>
          teams.createApproval({ ...input, agentId: bot.agentId, botId: bot.id }),
        // A click that landed before the card's id was kept could not edit the card for
        // everyone; it is edited now.
        approvalPosted: async (id, threadId, activityId) => {
          const row = await teams.approvalPosted(id, threadId, activityId);
          if (row && row.status !== "pending") await updateCard(bot, row);
        },
        discardApproval: (id) => teams.discardApproval(id),
        approval: (tokenHash) => teams.approvalForAgent(bot.agentId, tokenHash),
        runForReply: (tokenHash) => teams.runForReply(bot.agentId, tokenHash),
        claimReply: (runId) => teams.claimReply(bot.agentId, runId),
        releaseReply: (runId) => teams.releaseReply(runId),
      };
    },
    /** Settles the queue (tests). */
    idle: queue.idle,
  };
}

export type TeamsRuntime = ReturnType<typeof createTeamsRuntime>;

export type GraphMessage = {
  id?: string;
  createdDateTime?: string;
  messageType?: string;
  from?: {
    user?: { displayName?: string; id?: string; userIdentityType?: string };
    application?: { displayName?: string; id?: string };
  } | null;
  body?: { content?: string; contentType?: string };
  replyToId?: string | null;
};

export function simplify(list: GraphMessage[]): ThreadMessage[] {
  return list
    .filter((m) => (m.messageType ?? "message") === "message")
    .map((m) => ({
      from: m.from?.user?.displayName ?? m.from?.application?.displayName ?? "Someone",
      text:
        m.body?.contentType === "html" ? plainText(m.body.content ?? "") : (m.body?.content ?? ""),
    }));
}

export type TeamsAccess = {
  places(): Promise<ConversationRow[]>;
  graph(path: string): Promise<VendorResponse>;
  listChannels(team: ConversationRow): Promise<VendorResponse>;
  post(
    serviceUrl: string,
    conversation: string,
    text: string,
    replyTo?: string,
  ): Promise<VendorResponse>;
  newThread(team: ConversationRow, channelId: string, text: string): Promise<VendorResponse>;
  // Approval cards: an Adaptive Card posted as the bot (no text beside it, so Teams answers with
  // the card's own activity id), and the agent's own approvals.
  postCard(
    serviceUrl: string,
    conversation: string,
    card: Record<string, unknown>,
  ): Promise<VendorResponse>;
  newThreadCard(
    team: ConversationRow,
    channelId: string,
    card: Record<string, unknown>,
  ): Promise<VendorResponse>;
  createApproval(input: {
    tokenHash: string;
    conversationId: string;
    serviceUrl: string;
    approverAadId: string;
    approverName: string;
    title: string;
    body: string;
    facts: ApprovalFact[];
    days: number;
  }): Promise<ApprovalRow>;
  approvalPosted(id: string, threadId: string, activityId: string): Promise<void>;
  discardApproval(id: string): Promise<void>;
  approval(tokenHash: string): Promise<ApprovalRow | null>;
  runForReply(tokenHash: string): Promise<RunRow | null>;
  claimReply(runId: string): Promise<boolean>;
  releaseReply(runId: string): Promise<void>;
};
