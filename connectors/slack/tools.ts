import { z } from "zod";
import { channelQuery, slackError, type Channel } from "./vendor.js";
import {
  audited,
  Denied,
  json,
  listLimit,
  VendorFailed,
  type ToolContext,
  type Toolset,
} from "../../packages/connector-kit/src/common.js";

// Slack as the app's bot: it reaches only channels it is a member of (Slack refuses the rest
// with not_in_channel), within the key's `channels` limit. Direct messages are out of reach:
// the manifest asks for no im or mpim scope.

const CHANNEL = /^[CG][A-Z0-9]{2,20}$/;
const TS = /^\d{9,11}\.\d{6}$/;
const MAX_TEXT = 40_000;
const MAX_FILE_BYTES = 1_000_000;
// users.info calls one answer makes at most to put names on people; the rest keep only ids.
const MAX_NAME_LOOKUPS = 20;

async function ok(
  ctx: ToolContext,
  mark: (s: number) => void,
  req: Parameters<ToolContext["call"]>[0],
) {
  const res = await ctx.call(req);
  mark(res.status);
  if (res.status < 200 || res.status >= 300) throw new VendorFailed(res.status, slackError(res));
  return res.data as any;
}

const query = (params: Record<string, string | number | undefined>) =>
  new URLSearchParams(
    Object.entries(params).flatMap(([k, v]) =>
      v === undefined || v === "" ? [] : [[k, String(v)]],
    ),
  ).toString();

function channel(ctx: ToolContext, id: string): string {
  if (!CHANNEL.test(id))
    throw new Denied("Give a channel id such as C0123ABCD from slack_list_channels");
  const allowed = listLimit(ctx.limits, "channels");
  if (allowed && !allowed.includes(id)) throw new Denied("This agent may not use that channel");
  return id;
}

function timestamp(ts: string, what = "message"): string {
  if (!TS.test(ts)) throw new Denied(`Give the ${what}'s ts, such as 1730000000.123456`);
  return ts;
}

// A Slack ts or an ISO 8601 time, as the ts Slack's history bounds take.
function bound(value: string | undefined): string | undefined {
  if (value === undefined) return undefined;
  if (TS.test(value) || /^\d{9,11}$/.test(value)) return value;
  const ms = Date.parse(value);
  if (Number.isNaN(ms)) throw new Denied(`${value} is neither a Slack ts nor an ISO 8601 time`);
  return (ms / 1000).toFixed(6);
}

// Slack's mrkdwn markup to plain text: mentions and channel links by name, links as
// "label (url)", and its three escaped characters back.
export function plainText(text: string, names: Map<string, string>): string {
  return text
    .replace(/<([^<>]+)>/g, (_, inner: string) => {
      const bar = inner.indexOf("|");
      const target = bar < 0 ? inner : inner.slice(0, bar);
      const label = bar < 0 ? null : inner.slice(bar + 1);
      if (target.startsWith("@"))
        return `@${names.get(target.slice(1)) ?? label ?? target.slice(1)}`;
      if (target.startsWith("#")) return `#${label ?? target.slice(1)}`;
      if (target.startsWith("!subteam^")) return label ?? "@group";
      if (target.startsWith("!date^")) return label ?? target;
      if (target.startsWith("!")) return `@${target.slice(1)}`;
      return label && label !== target ? `${label} (${target})` : target;
    })
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&amp;", "&");
}

type Message = {
  ts: string;
  user?: string;
  bot_id?: string;
  username?: string;
  bot_profile?: { name?: string };
  text?: string;
  subtype?: string;
  thread_ts?: string;
  reply_count?: number;
  edited?: unknown;
  reactions?: { name: string; count: number }[];
  files?: { name?: string; title?: string }[];
  attachments?: { fallback?: string; text?: string }[];
};

// People's names for the authors and mentions of these messages. A failed lookup leaves the
// id alone: a name is a convenience, never worth failing the read.
async function names(ctx: ToolContext, messages: Message[]): Promise<Map<string, string>> {
  const ids = new Set<string>();
  for (const m of messages) {
    if (m.user) ids.add(m.user);
    for (const [, id] of (m.text ?? "").matchAll(/<@([UW][A-Z0-9]+)/g)) ids.add(id);
  }
  const found = new Map<string, string>();
  for (const id of [...ids].slice(0, MAX_NAME_LOOKUPS)) {
    const res = await ctx.call({ method: "GET", path: `/users.info?${query({ user: id })}` });
    const u = (res.data as { user?: any } | null)?.user;
    if (res.status === 200 && u)
      found.set(id, u.profile?.display_name || u.real_name || u.name || id);
  }
  return found;
}

function message(m: Message, who: Map<string, string>) {
  const text = m.text || (m.attachments ?? []).map((a) => a.fallback ?? a.text ?? "").join("\n");
  return {
    ts: m.ts,
    at: new Date(Number(m.ts) * 1000).toISOString(),
    ...(m.user ? { user_id: m.user, user: who.get(m.user) ?? null } : {}),
    ...(m.bot_id ? { bot: m.bot_profile?.name ?? m.username ?? m.bot_id } : {}),
    ...(m.subtype ? { subtype: m.subtype } : {}),
    text: plainText(text, who),
    ...(m.thread_ts ? { thread_ts: m.thread_ts } : {}),
    ...(m.reply_count ? { reply_count: m.reply_count } : {}),
    ...(m.edited ? { edited: true } : {}),
    ...(m.reactions?.length
      ? { reactions: m.reactions.map((r) => ({ emoji: r.name, count: r.count })) }
      : {}),
    ...(m.files?.length ? { files: m.files.map((f) => f.title || f.name || "file") } : {}),
  };
}

async function messages(ctx: ToolContext, list: Message[]) {
  const who = await names(ctx, list);
  return list.map((m) => message(m, who));
}

const channelId = z.string().describe("Channel id, e.g. C0123ABCD, from slack_list_channels.");
const ts = z.string().describe("The message's ts, e.g. 1730000000.123456.");
const cursor = z.string().max(500).optional().describe("The next value of the previous answer.");

export const slackTools: Toolset = {
  names: [
    "slack_list_channels",
    "slack_read_channel",
    "slack_read_thread",
    "slack_get_permalink",
    "slack_post_message",
    "slack_add_reaction",
    "slack_upload_text_file",
    "slack_find_user_by_email",
    "slack_delete_message",
  ],

  register(server, ctx, allowed) {
    const readOnly = { readOnlyHint: true, openWorldHint: true } as const;
    const writes = { readOnlyHint: false, destructiveHint: false, openWorldHint: true } as const;

    if (allowed.has("slack_list_channels"))
      server.registerTool(
        "slack_list_channels",
        {
          description:
            "List the Slack channels this agent can use: those the app was added to, within the agent's limits.",
          inputSchema: z.object({ cursor }),
          annotations: readOnly,
        },
        ({ cursor }) =>
          audited(ctx, "slack_list_channels", null, async (mark) => {
            const limit = listLimit(ctx.limits, "channels");
            const data = await ok(ctx, mark, {
              method: "GET",
              path: `/users.conversations?${channelQuery(200, cursor)}`,
            });
            const channels = ((data.channels ?? []) as Channel[])
              .filter((c) => !limit || limit.includes(c.id))
              .map((c) => ({
                id: c.id,
                name: c.name ?? null,
                private: Boolean(c.is_private),
                topic: c.topic?.value || null,
                purpose: c.purpose?.value || null,
              }));
            return json({ channels, next: data.response_metadata?.next_cursor || null });
          }),
      );

    if (allowed.has("slack_read_channel"))
      server.registerTool(
        "slack_read_channel",
        {
          description:
            "Read a Slack channel's messages, newest first. A message with reply_count has a thread: read it with slack_read_thread. oldest and latest bound the window (a ts or an ISO 8601 time).",
          inputSchema: z.object({
            channel: channelId,
            limit: z.number().int().min(1).max(100).default(20),
            oldest: z.string().optional(),
            latest: z.string().optional(),
            cursor,
          }),
          annotations: readOnly,
        },
        (args) =>
          audited(ctx, "slack_read_channel", args.channel, async (mark) => {
            const data = await ok(ctx, mark, {
              method: "GET",
              path: `/conversations.history?${query({
                channel: channel(ctx, args.channel),
                limit: args.limit,
                oldest: bound(args.oldest),
                latest: bound(args.latest),
                cursor: args.cursor,
              })}`,
            });
            return json({
              messages: await messages(ctx, data.messages ?? []),
              next: data.response_metadata?.next_cursor || null,
            });
          }),
      );

    if (allowed.has("slack_read_thread"))
      server.registerTool(
        "slack_read_thread",
        {
          description:
            "Read a Slack thread, oldest first: the message that started it, then its replies.",
          inputSchema: z.object({
            channel: channelId,
            thread_ts: z.string().describe("The ts of the message that started the thread."),
            limit: z.number().int().min(1).max(200).default(50),
            cursor,
          }),
          annotations: readOnly,
        },
        (args) =>
          audited(ctx, "slack_read_thread", `${args.channel} ${args.thread_ts}`, async (mark) => {
            const data = await ok(ctx, mark, {
              method: "GET",
              path: `/conversations.replies?${query({
                channel: channel(ctx, args.channel),
                ts: timestamp(args.thread_ts, "thread"),
                limit: args.limit,
                cursor: args.cursor,
              })}`,
            });
            return json({
              messages: await messages(ctx, data.messages ?? []),
              next: data.response_metadata?.next_cursor || null,
            });
          }),
      );

    if (allowed.has("slack_get_permalink"))
      server.registerTool(
        "slack_get_permalink",
        {
          description: "Get a link people can open to one Slack message.",
          inputSchema: z.object({ channel: channelId, ts }),
          annotations: readOnly,
        },
        (args) =>
          audited(ctx, "slack_get_permalink", `${args.channel} ${args.ts}`, async (mark) => {
            const data = await ok(ctx, mark, {
              method: "GET",
              path: `/chat.getPermalink?${query({
                channel: channel(ctx, args.channel),
                message_ts: timestamp(args.ts),
              })}`,
            });
            return json({ permalink: data.permalink });
          }),
      );

    if (allowed.has("slack_post_message"))
      server.registerTool(
        "slack_post_message",
        {
          description:
            "Post a message as the app in a Slack channel, or reply in a thread with thread_ts. Slack mrkdwn works: *bold*, _italic_, `code`, <@U0123ABCD> mentions a person.",
          inputSchema: z.object({
            channel: channelId,
            text: z.string().min(1).max(MAX_TEXT),
            thread_ts: z.string().optional().describe("Reply in this message's thread."),
            also_send_to_channel: z
              .boolean()
              .default(false)
              .describe("With thread_ts: also show the reply in the channel."),
          }),
          annotations: writes,
        },
        (args) =>
          audited(ctx, "slack_post_message", args.channel, async (mark) => {
            const data = await ok(ctx, mark, {
              method: "POST",
              path: "/chat.postMessage",
              body: {
                channel: channel(ctx, args.channel),
                text: args.text,
                ...(args.thread_ts
                  ? {
                      thread_ts: timestamp(args.thread_ts, "thread"),
                      reply_broadcast: args.also_send_to_channel,
                    }
                  : {}),
              },
            });
            return json({ channel: data.channel, ts: data.ts, thread_ts: args.thread_ts ?? null });
          }),
      );

    if (allowed.has("slack_add_reaction"))
      server.registerTool(
        "slack_add_reaction",
        {
          description:
            "Add an emoji reaction as the app to a Slack message, e.g. eyes or white_check_mark.",
          inputSchema: z.object({
            channel: channelId,
            ts,
            emoji: z.string().min(1).max(100).describe("Emoji name, with or without colons."),
          }),
          annotations: writes,
        },
        (args) =>
          audited(ctx, "slack_add_reaction", `${args.channel} ${args.ts}`, async (mark) => {
            const name = args.emoji.replace(/^:|:$/g, "");
            const res = await ctx.call({
              method: "POST",
              path: "/reactions.add",
              body: { channel: channel(ctx, args.channel), timestamp: timestamp(args.ts), name },
            });
            mark(res.status);
            // Already there is what the agent asked for.
            if ((res.data as { error?: string } | null)?.error === "already_reacted")
              return json({ reacted: true, already: true });
            if (res.status !== 200) throw new VendorFailed(res.status, slackError(res));
            return json({ reacted: true, already: false });
          }),
      );

    if (allowed.has("slack_upload_text_file"))
      server.registerTool(
        "slack_upload_text_file",
        {
          description:
            "Share a text file (up to 1 MB), such as a report or a CSV, as the app in a Slack channel or thread.",
          inputSchema: z.object({
            channel: channelId,
            filename: z.string().min(1).max(255).describe("e.g. report.csv"),
            content: z.string().min(1),
            title: z.string().max(255).optional(),
            comment: z
              .string()
              .max(MAX_TEXT)
              .optional()
              .describe("A message posted with the file."),
            thread_ts: z.string().optional().describe("Share it in this message's thread."),
          }),
          annotations: writes,
        },
        (args) =>
          audited(
            ctx,
            "slack_upload_text_file",
            `${args.channel} ${args.filename}`,
            async (mark) => {
              const id = channel(ctx, args.channel);
              if (/[/\\\x00-\x1f]/.test(args.filename))
                throw new Denied("Give a file name without slashes or control characters");
              const thread = args.thread_ts ? timestamp(args.thread_ts, "thread") : undefined;
              const length = Buffer.byteLength(args.content, "utf8");
              if (length > MAX_FILE_BYTES) throw new Denied("The file is over 1 MB");
              // Slack's upload: an address for the bytes, the bytes, then sharing it.
              const slot = await ok(ctx, mark, {
                method: "POST",
                path: `/files.getUploadURLExternal?${query({ filename: args.filename, length })}`,
              });
              await ok(ctx, mark, { method: "POST", path: slot.upload_url, body: args.content });
              await ok(ctx, mark, {
                method: "POST",
                path: "/files.completeUploadExternal",
                body: {
                  files: [{ id: slot.file_id, title: args.title ?? args.filename }],
                  channel_id: id,
                  ...(args.comment ? { initial_comment: args.comment } : {}),
                  ...(thread ? { thread_ts: thread } : {}),
                },
              });
              return json({ file_id: slot.file_id, channel: id, thread_ts: thread ?? null });
            },
          ),
      );

    if (allowed.has("slack_find_user_by_email"))
      server.registerTool(
        "slack_find_user_by_email",
        {
          description:
            "Find a person in the Slack workspace by email address, e.g. to mention them as <@id> in a message.",
          inputSchema: z.object({ email: z.email() }),
          annotations: readOnly,
        },
        ({ email }) =>
          audited(ctx, "slack_find_user_by_email", email, async (mark) => {
            const res = await ctx.call({
              method: "GET",
              path: `/users.lookupByEmail?${query({ email })}`,
            });
            mark(res.status);
            if (res.status === 404) return json({ found: false });
            if (res.status !== 200) throw new VendorFailed(res.status, slackError(res));
            const u = (res.data as { user: any }).user;
            return json({
              found: true,
              id: u.id,
              name: u.profile?.display_name || u.real_name || u.name,
              real_name: u.real_name ?? null,
              title: u.profile?.title || null,
              time_zone: u.tz ?? null,
              is_bot: Boolean(u.is_bot),
              deactivated: Boolean(u.deleted),
            });
          }),
      );

    if (allowed.has("slack_delete_message"))
      server.registerTool(
        "slack_delete_message",
        {
          description:
            "Delete a message the app posted. Slack cannot bring it back, and the app cannot delete anyone else's message.",
          inputSchema: z.object({ channel: channelId, ts }),
          annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
        },
        (args) =>
          audited(ctx, "slack_delete_message", `${args.channel} ${args.ts}`, async (mark) => {
            await ok(ctx, mark, {
              method: "POST",
              path: "/chat.delete",
              body: { channel: channel(ctx, args.channel), ts: timestamp(args.ts) },
            });
            return json({ deleted: true, channel: args.channel, ts: args.ts });
          }),
      );
  },
};
