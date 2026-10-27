import { z } from "zod";
import { hashReplyToken, looksLikeReplyToken } from "../../packages/server/src/teams/activity.js";
import {
  APPROVAL_DAYS,
  approvalCard,
  approvalStatus,
  hashApprovalToken,
  looksLikeApprovalToken,
  newApprovalToken,
} from "../../packages/server/src/teams/approval.js";
import {
  simplify,
  type GraphMessage,
  type TeamsAccess,
} from "../../packages/server/src/teams/runtime.js";
import type { ConversationRow } from "../../packages/server/src/teams/store.js";
import type { VendorResponse } from "../../packages/connector-kit/src/types.js";
import {
  audited,
  Denied,
  json,
  listLimit,
  VendorFailed,
  type ToolContext,
  type Toolset,
} from "../../packages/connector-kit/src/common.js";

// Teams, as the agent's own bot. Every tool reaches only where that bot is installed, within
// the key's `teams` and `channels` limits. A key with either limit reaches teams only, never
// chats. No tool takes a bot: the key's bot is the only one.

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const CHANNEL = /^19:[A-Za-z0-9_.-]+@thread\.(tacv2|skype)$/;
const CHAT = /^(19:[A-Za-z0-9_.-]+@(thread\.v2|unq\.gbl\.spaces)|a:[A-Za-z0-9_.-]+)$/;
const MESSAGE = /^\d{1,20}$/;
const MAX_TEXT = 20_000;

function access(ctx: ToolContext): TeamsAccess {
  if (!ctx.teams) throw new Error("Teams tools need the agent's bot");
  return ctx.teams;
}

const teamsLimit = (ctx: ToolContext) => listLimit(ctx.limits, "teams");
const channelsLimit = (ctx: ToolContext) => listLimit(ctx.limits, "channels");
const confined = (ctx: ToolContext) => Boolean(teamsLimit(ctx) || channelsLimit(ctx));

async function places(ctx: ToolContext) {
  const all = await access(ctx).places();
  const allowed = teamsLimit(ctx);
  return {
    teams: all.filter(
      (p) =>
        p.kind === "team" &&
        p.teamGroupId &&
        (!allowed || allowed.includes(p.teamGroupId.toLowerCase())),
    ),
    chats: confined(ctx) ? [] : all.filter((p) => p.kind !== "team"),
  };
}

async function team(ctx: ToolContext, id: string): Promise<ConversationRow> {
  if (!GUID.test(id)) throw new Denied("Give a team id from teams_list_teams");
  const found = (await places(ctx)).teams.find((t) => t.teamGroupId === id.toLowerCase());
  if (!found) throw new Denied("This agent cannot reach that team");
  return found;
}

// Posts and reads address a channel by its id alone, so it must be one of the team's
// channels, or a key limited to one team could reach another team's channel.
async function channel(
  ctx: ToolContext,
  place: ConversationRow,
  id: string,
  mark: (s: number) => void,
): Promise<string> {
  if (!CHANNEL.test(id)) throw new Denied("Give a channel id from teams_list_channels");
  const allowed = channelsLimit(ctx);
  if (allowed && !allowed.includes(id)) throw new Denied("This agent cannot reach that channel");
  const data = ok(await access(ctx).listChannels(place), mark);
  if (!((data?.conversations ?? []) as { id: string }[]).some((c) => c.id === id))
    throw new Denied("That channel is not in this team");
  return id;
}

async function chat(ctx: ToolContext, id: string, { readable = false } = {}) {
  if (!CHAT.test(id)) throw new Denied("Give a chat id from teams_list_teams");
  const found = (await places(ctx)).chats.find((c) => c.conversationId === id);
  if (!found) throw new Denied("This agent cannot reach that chat");
  if (readable && found.kind === "personal")
    throw new Denied("Direct chats can't be read; earlier turns are in this session");
  return found;
}

function ok(res: VendorResponse, mark: (s: number) => void) {
  mark(res.status);
  if (res.status < 200 || res.status >= 300) {
    const d = res.data as { error?: { message?: string; code?: string }; message?: string } | null;
    const message = d?.error?.message ?? d?.message;
    throw new VendorFailed(
      res.status,
      `Microsoft Teams answered ${res.status}${message ? `: ${message}` : ""}`,
    );
  }
  return res.data as any;
}

// Who sent a message: Graph's chatMessageFromIdentitySet names a user (teamworkUserIdentity, its
// id the person's Entra object id) or an application (a bot). A display name can be anything, so
// from_id is what identifies a person; it is left null for anyone Graph does not call an Entra
// user (anonymous, federated, consumer or phone users: their id is no object in this directory).
const ENTRA_USER = new Set(["aadUser", "onPremiseAadUser"]);
function sender(m: GraphMessage) {
  const user = m.from?.user;
  if (user)
    return {
      from_id:
        typeof user.id === "string" &&
        GUID.test(user.id) &&
        (user.userIdentityType === undefined || ENTRA_USER.has(user.userIdentityType))
          ? user.id.toLowerCase()
          : null,
      from_kind: "user" as const,
    };
  return { from_id: null, from_kind: m.from?.application ? ("application" as const) : null };
}

const message = (m: GraphMessage & { replies?: GraphMessage[] }) => ({
  id: m.id,
  created: m.createdDateTime ?? null,
  ...simplify([m])[0],
  ...sender(m),
  ...(m.replies ? { replies: m.replies.length } : {}),
});

// A place given as a team and channel, or as a chat.
const where = {
  team: z.string().optional().describe("Team id, with channel."),
  channel: z.string().optional().describe("Channel id in that team."),
  chat: z.string().optional().describe("Chat id, instead of a team and channel."),
};

async function resolve(
  ctx: ToolContext,
  input: { team?: string; channel?: string; chat?: string },
  mark: (s: number) => void,
  opts: { readable?: boolean } = {},
) {
  if (input.chat && !input.team && !input.channel)
    return { kind: "chat" as const, place: await chat(ctx, input.chat, opts) };
  if (input.team && input.channel && !input.chat) {
    const place = await team(ctx, input.team);
    return {
      kind: "channel" as const,
      place,
      channel: await channel(ctx, place, input.channel, mark),
    };
  }
  throw new Denied("Give either a team and a channel, or a chat");
}

// A sign-in name (UPN, guests' #EXT# form too) or an Entra object id.
const SIGN_IN = /^[A-Za-z0-9._%+'#-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$/;

// The approver as the Microsoft connection's directory knows them (User.ReadBasic.All, which
// every Teams ceiling expects): their object id, whatever was given, and their name. The agent
// names them, so this tells it whether a sign-in name exists here (docs/teams.md "Approval
// cards").
async function approver(ctx: ToolContext, value: string, mark: (s: number) => void) {
  const given = value.trim();
  if (!GUID.test(given) && !SIGN_IN.test(given))
    throw new Denied("Give the approver's sign-in name (e.g. ann@contoso.com) or Entra object id");
  const res = await ctx.call({
    method: "GET",
    path: `/v1.0/users/${encodeURIComponent(given)}?$select=id,displayName`,
  });
  if (res.status === 404) {
    mark(res.status);
    throw new Denied(`No one in this organization signs in as ${given}`);
  }
  const user = ok(res, mark) as { id?: string; displayName?: string };
  if (!user?.id || !GUID.test(user.id))
    throw new Error("Microsoft Graph answered a user with no id");
  return { id: user.id.toLowerCase(), name: (user.displayName ?? given).slice(0, 200) };
}

function text(value: string) {
  if (!value.trim()) throw new Denied("The message is empty");
  if (value.length > MAX_TEXT) throw new Denied(`Messages are at most ${MAX_TEXT} characters`);
  return value;
}

export const teamsTools: Toolset = {
  names: [
    "teams_list_teams",
    "teams_list_channels",
    "teams_read_messages",
    "teams_read_thread",
    "teams_post_message",
    "teams_reply",
    "teams_request_approval",
    "teams_get_approval",
  ],

  register(server, ctx, allowed) {
    const readOnly = { readOnlyHint: true, openWorldHint: true } as const;
    const sends = { readOnlyHint: false, destructiveHint: false, openWorldHint: true } as const;

    if (allowed.has("teams_list_teams"))
      server.registerTool(
        "teams_list_teams",
        {
          description:
            "List the teams and chats this agent's bot was added to and may use. Team ids are what other Teams tools take. " +
            "A direct chat carries user: the Entra object id of the person in it (Teams gives such chats no name).",
          inputSchema: z.object({}),
          annotations: readOnly,
        },
        () =>
          audited(ctx, "teams_list_teams", null, async () => {
            const { teams, chats } = await places(ctx);
            return json({
              teams: teams.map((t) => ({ team: t.teamGroupId, name: t.name })),
              chats: chats.map((c) => ({
                chat: c.conversationId,
                kind: c.kind === "personal" ? "direct" : "group",
                name: c.name,
                user: c.kind === "personal" ? c.userAadId : null,
              })),
            });
          }),
      );

    if (allowed.has("teams_list_channels"))
      server.registerTool(
        "teams_list_channels",
        {
          description: "List the channels of a team this agent may use.",
          inputSchema: z.object({ team: z.string().describe("Team id from teams_list_teams.") }),
          annotations: readOnly,
        },
        ({ team: id }) =>
          audited(ctx, "teams_list_channels", id, async (mark) => {
            const place = await team(ctx, id);
            const data = ok(await access(ctx).listChannels(place), mark);
            const allowedChannels = channelsLimit(ctx);
            const channels = ((data?.conversations ?? []) as { id: string; name?: string }[])
              .filter((c) => !allowedChannels || allowedChannels.includes(c.id))
              .map((c) => ({ channel: c.id, name: c.name ?? "General" }));
            return json({ channels });
          }),
      );

    if (allowed.has("teams_read_messages"))
      server.registerTool(
        "teams_read_messages",
        {
          description:
            "Read recent messages in a channel (top-level posts, newest first, with reply counts) or a group chat. Direct chats can't be read. " +
            "Each message has from (a display name, which anyone can set), from_id (the sender's Microsoft Entra object id; null for apps) and from_kind (user or application).",
          inputSchema: z.object({
            ...where,
            max_results: z.number().int().min(1).max(50).default(20),
          }),
          annotations: readOnly,
        },
        (input) =>
          audited(
            ctx,
            "teams_read_messages",
            input.chat ?? `${input.team ?? ""}/${input.channel ?? ""}`,
            async (mark) => {
              const r = await resolve(ctx, input, mark, { readable: true });
              const path =
                r.kind === "chat"
                  ? `/v1.0/chats/${encodeURIComponent(r.place.conversationId)}/messages?$top=${input.max_results}&$orderby=createdDateTime desc`
                  : `/v1.0/teams/${r.place.teamGroupId}/channels/${encodeURIComponent(r.channel)}/messages?$top=${input.max_results}&$expand=replies`;
              const data = ok(await access(ctx).graph(path), mark);
              const messages = ((data?.value ?? []) as GraphMessage[])
                .filter((m) => (m.messageType ?? "message") === "message")
                .map(message);
              return json({ messages, more: Boolean(data?.["@odata.nextLink"]) });
            },
          ),
      );

    if (allowed.has("teams_read_thread"))
      server.registerTool(
        "teams_read_thread",
        {
          description:
            "Read a channel post and its replies, oldest first. Each message has from, from_id (the sender's Microsoft Entra object id; null for apps) and from_kind (user or application).",
          inputSchema: z.object({
            team: z.string().describe("Team id from teams_list_teams."),
            channel: z.string().describe("Channel id from teams_list_channels."),
            message_id: z.string().describe("The post's id, from teams_read_messages."),
          }),
          annotations: readOnly,
        },
        (input) =>
          audited(
            ctx,
            "teams_read_thread",
            `${input.channel}/${input.message_id}`,
            async (mark) => {
              const place = await team(ctx, input.team);
              const ch = await channel(ctx, place, input.channel, mark);
              if (!MESSAGE.test(input.message_id))
                throw new Denied("Give a message id from teams_read_messages");
              const base = `/v1.0/teams/${place.teamGroupId}/channels/${encodeURIComponent(ch)}/messages/${input.message_id}`;
              const root = ok(await access(ctx).graph(base), mark) as GraphMessage;
              const replies = ok(await access(ctx).graph(`${base}/replies?$top=50`), mark);
              const list = ((replies?.value ?? []) as GraphMessage[]).reverse();
              return json({
                post: message(root),
                replies: list
                  .filter((m) => (m.messageType ?? "message") === "message")
                  .map(message),
              });
            },
          ),
      );

    if (allowed.has("teams_post_message"))
      server.registerTool(
        "teams_post_message",
        {
          description:
            "Post a message as this agent's bot: a new post in a channel (or a reply to a post with reply_to), or a message in a chat. Markdown is supported.",
          inputSchema: z.object({
            ...where,
            text: z.string(),
            reply_to: z.string().optional().describe("A channel post's id, to reply under it."),
          }),
          annotations: sends,
        },
        (input) =>
          audited(
            ctx,
            "teams_post_message",
            input.chat ?? `${input.channel ?? ""}${input.reply_to ? `/${input.reply_to}` : ""}`,
            async (mark) => {
              const body = text(input.text);
              const r = await resolve(ctx, input, mark);
              const teams = access(ctx);
              let data;
              if (r.kind === "chat") {
                if (input.reply_to) throw new Denied("reply_to is for channel posts");
                data = ok(await teams.post(r.place.serviceUrl, r.place.conversationId, body), mark);
              } else if (input.reply_to) {
                if (!MESSAGE.test(input.reply_to))
                  throw new Denied("Give a post id from teams_read_messages");
                data = ok(
                  await teams.post(
                    r.place.serviceUrl,
                    `${r.channel};messageid=${input.reply_to}`,
                    body,
                    input.reply_to,
                  ),
                  mark,
                );
              } else data = ok(await teams.newThread(r.place, r.channel, body), mark);
              return json({ posted: true, id: data?.activityId ?? data?.id ?? null });
            },
          ),
      );

    if (allowed.has("teams_request_approval"))
      server.registerTool(
        "teams_request_approval",
        {
          description:
            "Ask one person to approve or reject something: posts a card as this agent's bot, in a channel (a new post) or a chat, with Approve and Reject buttons and an optional comment. " +
            "Only that person's click counts, in that team or chat: the click arrives in a request authenticated by a Bot Framework token, and the approver is matched on the Entra object id Teams sends. " +
            "You choose the approver; nothing here checks that they are the right person to approve this. The guarantee is that the person you named clicked, not that this is the right person. " +
            "Returns approval_id: read the answer later with teams_get_approval. Requests expire after 7 days.",
          inputSchema: z.object({
            ...where,
            approver: z
              .string()
              .max(320)
              .describe("The person's sign-in name (e.g. ann@contoso.com) or Entra object id."),
            title: z.string().min(1).max(200),
            text: z.string().min(1).max(4000).describe("What is asked, shown on the card."),
            facts: z
              .array(z.object({ name: z.string().min(1).max(100), value: z.string().max(1000) }))
              .max(20)
              .optional()
              .describe("Name and value pairs shown on the card, e.g. Amount: 1,200 USD."),
          }),
          annotations: sends,
        },
        (input) =>
          audited(
            ctx,
            "teams_request_approval",
            `${input.chat ?? `${input.team ?? ""}/${input.channel ?? ""}`} → ${input.approver}`,
            async (mark, retarget) => {
              if (!input.title.trim() || !input.text.trim()) throw new Denied("The card is empty");
              const r = await resolve(ctx, input, mark);
              const person = await approver(ctx, input.approver, mark);
              const where = r.kind === "chat" ? r.place.conversationId : r.channel;
              retarget(`${where} → ${person.id}`);
              const teams = access(ctx);
              const token = newApprovalToken();
              const row = await teams.createApproval({
                tokenHash: hashApprovalToken(token),
                conversationId: r.place.conversationId,
                serviceUrl: r.place.serviceUrl,
                approverAadId: person.id,
                approverName: person.name,
                title: input.title.trim(),
                body: input.text.trim(),
                facts: input.facts ?? [],
                days: APPROVAL_DAYS,
              });
              let threadId: string;
              let messageId: string | null;
              try {
                if (r.kind === "chat") {
                  const data = ok(
                    await teams.postCard(
                      r.place.serviceUrl,
                      r.place.conversationId,
                      approvalCard(row),
                    ),
                    mark,
                  );
                  threadId = r.place.conversationId;
                  messageId = data?.id ?? null;
                } else {
                  const data = ok(
                    await teams.newThreadCard(r.place, r.channel, approvalCard(row)),
                    mark,
                  );
                  messageId = data?.activityId ?? null;
                  threadId =
                    data?.id ?? (messageId ? `${r.channel};messageid=${messageId}` : r.channel);
                }
              } catch (error) {
                await teams.discardApproval(row.id);
                throw error;
              }
              // Without the message id the card cannot be edited later; the answer still counts.
              if (messageId) await teams.approvalPosted(row.id, threadId, messageId);
              return json({ approval_id: token, message_id: messageId, expires: row.expiresAt });
            },
          ),
      );

    if (allowed.has("teams_get_approval"))
      server.registerTool(
        "teams_get_approval",
        {
          description:
            "Read the answer to an approval this agent asked for with teams_request_approval: pending, approved, rejected or expired, who answered (their Entra object id and name), when, and their comment.",
          inputSchema: z.object({ approval_id: z.string().max(100) }),
          annotations: readOnly,
        },
        (input) =>
          audited(ctx, "teams_get_approval", null, async (_mark, retarget) => {
            if (!looksLikeApprovalToken(input.approval_id))
              throw new Denied("That approval id is not valid");
            // Another agent's id finds nothing, the same as a made-up one.
            const a = await access(ctx).approval(hashApprovalToken(input.approval_id));
            if (!a) throw new Denied("No approval of this agent has that id");
            retarget(a.id);
            return json({
              status: approvalStatus(a),
              title: a.title,
              approver: { id: a.approverAadId, name: a.approverName },
              decided_by: a.decidedByAadId ? { id: a.decidedByAadId, name: a.decidedByName } : null,
              decided_at: a.decidedAt,
              comment: a.comment,
              expires: a.expiresAt,
            });
          }),
      );

    if (allowed.has("teams_reply"))
      server.registerTool(
        "teams_reply",
        {
          description:
            "Answer the Teams message that started this session, in the same thread. Use the reply_token given with the message.",
          inputSchema: z.object({ reply_token: z.string(), text: z.string() }),
          annotations: sends,
        },
        (input) =>
          audited(ctx, "teams_reply", null, async (mark, retarget) => {
            const body = text(input.text);
            if (!looksLikeReplyToken(input.reply_token))
              throw new Denied("That reply token is not valid");
            const teams = access(ctx);
            // Another agent's token, an old one or a made-up one all find nothing.
            const run = await teams.runForReply(hashReplyToken(input.reply_token));
            if (!run) throw new Denied("That reply token is not valid or has expired");
            retarget(run.threadId);
            const place = (await teams.places()).find(
              (p) => p.conversationId === run.conversationId,
            );
            if (!place) throw new Denied("The bot is no longer in that conversation");
            // The key's limits may have changed since the message arrived.
            if (place.kind === "team") {
              await channel(
                ctx,
                await team(ctx, place.teamGroupId ?? ""),
                run.threadId.split(";")[0],
                mark,
              );
            } else if (confined(ctx)) throw new Denied("This agent cannot reach that chat");
            // A reply is used only by a post that went through.
            if (!(await teams.claimReply(run.id)))
              throw new Denied("That reply token is not valid or has expired");
            try {
              ok(
                await teams.post(place.serviceUrl, run.threadId, body, run.triggerActivityId),
                mark,
              );
            } catch (error) {
              await teams.releaseReply(run.id);
              throw error;
            }
            return json({ replied: true });
          }),
      );
  },
};
