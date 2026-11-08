import { randomUUID } from "node:crypto";
import { effectiveCapabilities, type Connector } from "../../../connector-kit/src/connector.js";
import { credentialAad, type CredentialKey } from "../secrets.js";
import type { Store } from "../store.js";
import { ConnectError, type VendorResponse } from "../../../connector-kit/src/types.js";
import {
  decide,
  directChatFiles,
  graphFiles,
  guid,
  installChange,
  placeOf,
  plainText,
  promptText,
  newReplyToken,
  hashReplyToken,
  type Activity,
  type Place,
  type TeamsFile,
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
import {
  consentAnswer,
  consentCard,
  downloadFile,
  fileInfoCard,
  fileStorageUrl,
  offerAad,
  OFFER_HOURS,
} from "./offer.js";
import { drainer } from "../queue.js";
import { noReport, type Report } from "../report.js";
import {
  nextAttempt,
  openTriggerSecret,
  RETRY_WINDOW_MS,
  RetryLater,
  type RunCounts,
  type RunInput,
  type TriggerRunner,
} from "./runner.js";
import { Unauthorized, type BotAuth } from "./jwt.js";
import type {
  AgentTeamsState,
  BotRow,
  ConversationRow,
  EventRow,
  RetryRow,
  RunRow,
  TeamsStore,
  TriggerRow,
} from "./store.js";

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
const runFilesAad = (runId: string) => `agent-connect:teams-run-files:${runId}`;
const deliveryAad = (runId: string) => `agent-connect:teams-delivery:${runId}`;

// What a run's next hand-over to its runner sends again, sealed on the run meanwhile.
type Delivery = {
  deliveryId: string;
  eventId: string;
  // The message, so each hand-over decides on it again as the agent and trigger stand then.
  activity: Activity;
  requestedBy: string;
  text: string;
  replyToken: string;
};

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
  // Fetches the files an agent sends and puts them in the recipient's OneDrive.
  fetch?: typeof fetch;
  report?: Report;
  // How long a start the runner did not take is tried again (runner.ts nextAttempt).
  retryWindowMs?: number;
};

export function createTeamsRuntime(deps: TeamsRuntimeDeps) {
  const { teams, key, bots } = deps;
  const fetchImpl = deps.fetch ?? fetch;
  const report = deps.report ?? noReport;
  const retryWindowMs = deps.retryWindowMs ?? RETRY_WINDOW_MS;
  const botTags = (bot: BotRow) => ({
    app: "teams",
    org: bot.organizationId,
    agent: bot.agentId,
    bot: bot.id,
  });
  // Uploads of accepted files, which outlive the invoke that started them.
  const uploads = new Set<Promise<void>>();

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
      if (activity.type === "invoke" && activity.name === "fileConsent/invoke") {
        await fileConsent(bot, activity);
        return reply(200);
      }
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
    const audit = (decision: "allowed" | "denied", reason: string) =>
      clickAudit(bot, "teams:approval", id, decision, reason);
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

  async function clickAudit(
    bot: BotRow,
    tool: string,
    target: string | null,
    decision: "allowed" | "denied",
    reason: string,
  ) {
    const state = await teams.agentState(bot.agentId);
    await deps.store.audit({
      organizationId: bot.organizationId,
      agentId: bot.agentId,
      agentName: state?.agentName ?? "",
      connectorId: "teams",
      tool,
      target,
      decision,
      reason: reason.slice(0, 500),
      vendorStatus: null,
    });
  }

  // ---- files sent in a direct chat (offer.ts) ----

  /** Accept or Decline on a file consent card, after the same request checks as a card click.
   *  Counts once, from the person the file was offered to, in that chat, unexpired. On accept
   *  the file goes to the upload address Teams gives (their OneDrive), after this answers. */
  async function fileConsent(bot: BotRow, a: Activity) {
    const answer = consentAnswer(a.value);
    const clicker = guid(a.from?.aadObjectId);
    const place = placeOf(a);
    const offerId = guid(answer?.offer);
    const audit = (decision: "allowed" | "denied", reason: string) =>
      clickAudit(bot, "teams:file", offerId, decision, reason);
    if (!answer || !offerId || !clicker || place?.kind !== "personal") {
      await audit("denied", `unknown file answer from ${clicker ?? "unknown"}`);
      return;
    }
    const target = fileStorageUrl(answer.upload?.uploadUrl);
    const upload = answer.upload;
    if (answer.action === "accept" && (!upload || !target)) {
      await audit("denied", `accepted by ${clicker} with no OneDrive upload address`);
      return;
    }
    const offer = await teams.takeFileOffer(bot.id, offerId, place.conversationId, clicker);
    if (!offer) {
      await audit("denied", `${answer.action} by ${clicker}: no such file waiting for them there`);
      return;
    }
    const card = a.replyToId;
    const dropCard = () =>
      card
        ? bots
            .connector(botCredential(key, bot), offer.serviceUrl, {
              method: "DELETE",
              path: `/v3/conversations/${encodeURIComponent(place.conversationId)}/activities/${encodeURIComponent(card)}`,
            })
            .catch((e) => console.error(`[teams] removing file card ${offer.id}`, e))
        : null;
    if (answer.action === "decline") {
      await audit("allowed", `${offer.name} declined`);
      await dropCard();
      return;
    }
    const say = (text: string) =>
      notice(bot, offer.serviceUrl, place.conversationId, text).catch((e) =>
        console.error(`[teams] notice for file ${offer.id}`, e),
      );
    const work = (async () => {
      let bytes: Buffer;
      try {
        bytes = Buffer.from(
          key.open({ ciphertext: offer.content, keyId: offer.contentKeyId }, offerAad(offer.id))
            .content,
          "base64",
        );
      } catch (e) {
        console.error(`[teams] opening file ${offer.id}`, e);
        await audit("denied", `${offer.name}: the waiting file could not be opened`);
        await say(`I couldn't send ${offer.name}. Ask me to send it again.`);
        await dropCard();
        return;
      }
      let failed: string | null = null;
      try {
        const res = await fetchImpl(target!, {
          method: "PUT",
          redirect: "error",
          headers: {
            "content-type": "application/octet-stream",
            "content-range": `bytes 0-${bytes.length - 1}/${bytes.length}`,
          },
          body: bytes,
        });
        await res.body?.cancel();
        // Teams' upload contract: only 200 and 201 mean the file is complete.
        if (res.status !== 200 && res.status !== 201) failed = `OneDrive answered ${res.status}`;
      } catch (e) {
        failed = `the upload failed: ${e instanceof Error ? e.message : String(e)}`;
      }
      if (failed) {
        // The consent card stays, so the person can accept again.
        await teams.restoreFileOffer(offer);
        await audit("denied", `${offer.name}: ${failed}`);
        await say(
          `I couldn't save ${offer.name} to your OneDrive. Select Accept on it to try again, or ask me to send it again.`,
        );
        return;
      }
      const sent = await post(
        bot,
        offer.serviceUrl,
        place.conversationId,
        fileInfoCard(upload!, offer.name),
      ).catch((e: unknown) => {
        console.error(`[teams] file card for ${offer.id}`, e);
        return { status: 0 };
      });
      if (sent.status < 200 || sent.status >= 300) {
        report("teams_send_failed", null, {
          ...botTags(bot),
          what: "file_card",
          status: sent.status,
        });
        await audit(
          "denied",
          `${offer.name} saved, but Teams refused its file card: ${sent.status}`,
        );
        await say(`I saved ${offer.name} to your OneDrive, but couldn't show it here.`);
        await dropCard();
        return;
      }
      await audit("allowed", `${offer.name} saved to their OneDrive`);
      await dropCard();
    })().catch((e) => {
      console.error(`[teams] sending file ${offer.id}`, e);
      report("teams_send_failed", e, { ...botTags(bot), what: "file" });
    });
    uploads.add(work);
    void work.finally(() => uploads.delete(work));
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
      if (res.status < 200 || res.status >= 300) {
        console.error(`[teams] editing approval card ${approval.id}: ${res.status}`);
        report("teams_send_failed", null, {
          ...botTags(bot),
          what: "approval_card",
          status: res.status,
        });
      }
    } catch (error) {
      console.error(`[teams] editing approval card ${approval.id}`, error);
      report("teams_send_failed", error, { ...botTags(bot), what: "approval_card" });
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
      report("teams_event_failed", error, { app: "teams", bot: event.botId });
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

  /** decide() on the message as the agent, its grant and its trigger stand now. */
  async function authorize(
    bot: BotRow,
    a: Activity,
    place: Place,
    state: AgentTeamsState | null,
    trigger: TriggerRow | null,
    counts: RunCounts,
    // Without text, only files make it a message; a channel's or group chat's are on Graph. The
    // caller passes them when it has read them already.
    early?: TeamsFile[] | null,
  ) {
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
    const files =
      early !== undefined ? early : a.text?.trim() ? null : await messageFiles(bot, place, a);
    return decide({
      activity: a,
      place,
      botAppId: bot.botAppId!,
      tenantId: bot.tenantId,
      agentActive: state?.agentStatus === "active",
      capabilities,
      limits: state?.grant?.limits ?? {},
      trigger,
      runnerUnready,
      counts,
      hasFiles: Boolean(files?.length),
    });
  }

  async function respond(bot: BotRow, event: EventRow, a: Activity, place: Place) {
    const [state, trigger] = await Promise.all([
      teams.agentState(bot.agentId),
      teams.trigger(bot.agentId),
    ]);
    const runner = trigger ? deps.runners.get(trigger.runner) : undefined;
    const counts = await teams.runCounts(bot.agentId, bot.id);
    const early = a.text?.trim() ? null : await messageFiles(bot, place, a);
    const decision = await authorize(bot, a, place, state, trigger, counts, early);
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
    const since = Date.now();
    let run: RunRow | null = null;
    if (reusable) {
      if (!(await teams.continueRun(latest.id, a.id!, hashReplyToken(token))))
        return teams.finishEvent(event.id, "held", "conversation busy");
      run = { ...latest, triggerActivityId: a.id! };
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

    let input: RunInput;
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
      const files = early ?? (await messageFiles(bot, place, a));
      await teams.setRunFiles(
        run.id,
        files.length ? key.seal({ files: JSON.stringify(files) }, runFilesAad(run.id)) : null,
      );
      const text = promptText({
        botName: bot.displayName,
        sender: { name: a.from?.name ?? "Someone", aadObjectId: decision.sender },
        place,
        channelName,
        thread: context?.messages ?? null,
        message: plainText(a.text ?? ""),
        replyToken: token,
        continuing: Boolean(reusable),
        files,
      });
      input = {
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
        deliveryId: randomUUID(),
      };
    } catch (error) {
      return failStart(bot, run, event.id, error, 0);
    }
    await deliver(bot, run, event.id, a, input, reusable ? latest.sessionId : null, 1, since);
  }

  /** Hands the message to the runner: started, or tried again later while the runner may yet
   *  take it (RetryLater, within the retry window; the run stays starting, so the thread's
   *  later messages stay held behind it), or failed. */
  async function deliver(
    bot: BotRow,
    run: RunRow,
    eventId: string,
    activity: Activity,
    input: RunInput,
    sessionId: string | null,
    attempt: number,
    since: number,
  ) {
    const runner = deps.runners.get(run.runner)!;
    try {
      if (sessionId) {
        await runner.continue(sessionId, input);
        await teams.updateRun(run.id, { state: "running" });
      } else {
        await teams.updateRun(run.id, { state: "running", sessionId: await runner.start(input) });
      }
      await teams.finishEvent(eventId, "run", null);
    } catch (error) {
      const at =
        error instanceof RetryLater
          ? nextAttempt(attempt, since, retryWindowMs, error.retryAfterMs)
          : null;
      if (at === null) return failStart(bot, run, eventId, error, attempt);
      const why = (error as Error).message;
      console.warn(`[teams] run ${run.id} attempt ${attempt}: ${why}; again later`);
      const delivery: Delivery = {
        deliveryId: input.deliveryId,
        eventId,
        activity,
        requestedBy: input.requestedBy,
        text: input.text,
        replyToken: input.replyToken,
      };
      await teams.scheduleRetry(
        run.id,
        attempt,
        new Date(at),
        key.seal({ delivery: JSON.stringify(delivery) }, deliveryAad(run.id)),
      );
      await teams.finishEvent(eventId, "run", "waiting to hand the message to the runner again");
    }
  }

  // The one place a response that could not start is told: the run fails with the reason, the
  // person is told once, onError once.
  async function failStart(
    bot: BotRow,
    run: RunRow,
    eventId: string | null,
    error: unknown,
    attempts: number,
  ) {
    console.error(`[teams] run ${run.id} failed to start`, error);
    report("teams_run_start_failed", error, { ...botTags(bot), runner: run.runner });
    await teams.updateRun(run.id, {
      state: "failed",
      error: (error instanceof Error ? error.message : String(error)).slice(0, 500),
    });
    await runNotice(bot, run, FAILED_NOTICE);
    if (eventId) await teams.finishEvent(eventId, "failed", "could not start the response");
  }

  const retries = drainer("teams-retries", async () => {
    const row = await teams.claimRetry();
    if (!row) return false;
    // Whatever stops a hand-over fails its run: left starting, it would be claimed again forever
    // and hold its thread.
    await retry(row)
      .catch(async (error) => {
        const bot = await teams.botById(row.botId);
        if (bot) await failStart(bot, row, null, error, row.attempts);
      })
      .catch((error) => console.error(`[teams] retrying run ${row.id}`, error));
    return true;
  });

  // The next hand-over of a run the runner did not take, as the trigger stands now.
  async function retry(row: RetryRow) {
    const bot = await teams.botById(row.botId);
    if (!bot) return;
    const sent = JSON.parse(
      key.open({ ciphertext: row.delivery, keyId: row.deliveryKeyId }, deliveryAad(row.id))
        .delivery,
    ) as Delivery;
    const [state, trigger] = await Promise.all([
      teams.agentState(row.agentId),
      teams.trigger(row.agentId),
    ]);
    const runner = deps.runners.get(row.runner);
    const fail = (why: string) =>
      failStart(bot, row, sent.eventId, new ConnectError(why), row.attempts);
    const since = Date.parse(row.replyTokenAt);
    if (Date.now() > since + retryWindowMs) return fail("The runner did not take it in time");
    if (!trigger || trigger.runner !== row.runner)
      return fail("Its trigger was turned off or changed runner");
    // Access may have changed while it waited. The caps were checked when it first came, and it
    // has held its place among them since, so they are not counted again.
    const place = placeOf(sent.activity)!;
    const none = { active: 0, lastHour: 0, spentToday: 0 };
    const decision = await authorize(bot, sent.activity, place, state, trigger, none);
    if (decision.action !== "respond") return fail(`No longer allowed: ${decision.reason}`);
    if (decision.checkMember && !(await member(bot, sent.activity, place)))
      return fail("No longer allowed: sender is not a member of the organization");
    if (!runner || !state) return fail(`Runner ${row.runner} is not offered by this deployment`);
    const secret = openTriggerSecret(key, bot.organizationId, trigger);
    try {
      await verified(runner, bot, state.agentName, trigger, secret);
    } catch (error) {
      return failStart(bot, row, sent.eventId, error, row.attempts);
    }
    const input: RunInput = {
      org: bot.organizationId,
      agentId: row.agentId,
      runId: row.id,
      botId: bot.id,
      requestedBy: sent.requestedBy,
      text: sent.text,
      replyToken: sent.replyToken,
      budgetCents: trigger.budgetCents,
      trigger,
      secret,
      deliveryId: sent.deliveryId,
    };
    const attempt = row.attempts + 1;
    await deliver(bot, row, sent.eventId, sent.activity, input, row.sessionId, attempt, since);
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

  // The files sent with a message. A direct chat's come with the activity; a channel's or group
  // chat's only with Graph's copy of it, read with the bot's own token. Unreadable: none listed.
  async function messageFiles(bot: BotRow, place: Place, a: Activity): Promise<TeamsFile[]> {
    if (place.kind === "personal") return directChatFiles(a);
    let path = `/v1.0/chats/${encodeURIComponent(place.conversationId)}/messages/${a.id}`;
    if (place.kind === "team") {
      if (!place.teamGroupId || !place.channelId) return [];
      const root = /;messageid=(\d+)/.exec(place.threadId)?.[1];
      const base = `/v1.0/teams/${place.teamGroupId}/channels/${encodeURIComponent(place.channelId)}/messages`;
      path = root && root !== a.id ? `${base}/${root}/replies/${a.id}` : `${base}/${a.id}`;
    }
    const res = await bots.graph(botCredential(key, bot), path).catch((error: unknown) => {
      console.error(`[teams] reading the files of message ${a.id}`, error);
      return null;
    });
    if (res?.status !== 200) return [];
    return graphFiles(res.data as GraphMessage, place.kind === "team" ? "channel" : "groupChat");
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
      await notice(bot, place.serviceUrl, run.threadId, text, run.triggerActivityId).catch((e) => {
        console.error(`[teams] notice for run ${run.id}`, e);
        report("teams_send_failed", e, { ...botTags(bot), what: "notice" });
      });
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
      retries.start();
      timers = [
        setInterval(() => {
          kick();
          retries.kick();
          watch().catch((error) => console.error("[teams] watch", error));
        }, 5000),
        setInterval(
          () =>
            Promise.all([teams.pruneEvents(), teams.pruneFileOffers()]).catch((error) =>
              console.error("[teams] prune", error),
            ),
          3600_000,
        ),
      ];
      for (const t of timers) t.unref();
    },
    /** Stops claiming events and waits for the one in hand, so a rollout neither strands it
     *  in 'processing' nor answers it twice. */
    async stop() {
      for (const t of timers) clearInterval(t);
      await Promise.all([queue.stop(), retries.stop()]);
    },
    /** Hands over every run whose next attempt is due, now (tests). */
    async retryDue() {
      retries.kick();
      await retries.idle();
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
        files: async (run) => {
          const sealed = await teams.runFiles(bot.agentId, run.id);
          if (!sealed) return [];
          return JSON.parse(key.open(sealed, runFilesAad(run.id)).files) as TeamsFile[];
        },
        download: (url, max) => downloadFile(fetchImpl, url, max),
        offerFile: async ({ conversation, recipientAadId, name, bytes }) => {
          const id = randomUUID();
          await teams.createFileOffer({
            id,
            botId: bot.id,
            conversationId: conversation.conversationId,
            serviceUrl: conversation.serviceUrl,
            recipientAadId,
            name,
            size: bytes.length,
            sealed: key.seal({ content: bytes.toString("base64") }, offerAad(id)),
            hours: OFFER_HOURS,
          });
          try {
            const res = await post(
              bot,
              conversation.serviceUrl,
              conversation.conversationId,
              consentCard(id, name, bytes.length),
            );
            if (res.status < 200 || res.status >= 300) await teams.discardFileOffer(id);
            return res;
          } catch (error) {
            await teams.discardFileOffer(id);
            throw error;
          }
        },
        claimReply: (runId) => teams.claimReply(bot.agentId, runId),
        releaseReply: (runId) => teams.releaseReply(runId),
      };
    },
    /** Settles the queue and the uploads of accepted files (tests). */
    async idle() {
      await queue.idle();
      await Promise.all(uploads);
    },
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
  attachments?: { contentType?: string; contentUrl?: string; name?: string }[];
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
  // The files sent with the run's latest message.
  files(run: RunRow): Promise<TeamsFile[]>;
  claimReply(runId: string): Promise<boolean>;
  // A file's bytes from its SharePoint or OneDrive download link (offer.ts).
  download(url: string, max: number): Promise<Buffer>;
  // Offers a file in a direct chat with a consent card, for that chat's person to accept.
  offerFile(input: {
    conversation: ConversationRow;
    recipientAadId: string;
    name: string;
    bytes: Buffer;
  }): Promise<VendorResponse>;
  releaseReply(runId: string): Promise<void>;
};
