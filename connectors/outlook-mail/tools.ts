import { z } from "zod";
import {
  audited,
  Denied,
  json,
  textLimit,
  VendorFailed,
  type ToolContext,
  type Toolset,
} from "../../packages/connector-kit/src/common.js";

// Every call goes to the mailbox in the agent's grant. No tool takes a mailbox argument, so
// an agent cannot name another one.

const MOVE_TO = ["inbox", "archive"];
const MESSAGE_ID = /^[A-Za-z0-9_=-]{1,1024}$/;
const LIST_FIELDS = "id,subject,from,toRecipients,receivedDateTime,isRead,bodyPreview";
const MESSAGE_FIELDS = `${LIST_FIELDS},ccRecipients,body,conversationId,hasAttachments`;

// Where a listing continues: Graph pages a folder or calendar with $skip and a search with $skiptoken, so
// the cursor says which ("s" or "t") and carries only that value, never a path.
export function nextPage(data: { "@odata.nextLink"?: string }): string | null {
  const q = URL.parse(data["@odata.nextLink"] ?? "")?.searchParams;
  const token = q?.get("$skiptoken");
  const skip = q?.get("$skip");
  return token ? `t${token}` : skip ? `s${skip}` : null;
}
export function pageParam(cursor: string | undefined): string[] {
  if (cursor === undefined) return [];
  if (/^s\d{1,6}$/.test(cursor)) return [`$skip=${cursor.slice(1)}`];
  if (/^t.+$/.test(cursor)) return [`$skiptoken=${encodeURIComponent(cursor.slice(1))}`];
  throw new Denied("Give cursor as the next value from the previous call");
}

type Address = { emailAddress?: { name?: string; address?: string } };
type Message = {
  id: string;
  subject?: string;
  from?: Address;
  toRecipients?: Address[];
  ccRecipients?: Address[];
  receivedDateTime?: string;
  isRead?: boolean;
  bodyPreview?: string;
  body?: { content?: string };
  conversationId?: string;
  hasAttachments?: boolean;
};

export function mailboxPath(ctx: ToolContext): string {
  const mailbox = textLimit(ctx.limits, "mailbox");
  if (!mailbox)
    throw new Denied("No mailbox is set for this agent. Ask an administrator to set one.");
  return `/v1.0/users/${encodeURIComponent(mailbox)}`;
}

function messagePath(ctx: ToolContext, id: string): string {
  if (!MESSAGE_ID.test(id)) throw new Denied("Give a message id from mail_list_messages");
  return `${mailboxPath(ctx)}/messages/${encodeURIComponent(id)}`;
}

export async function ok(
  ctx: ToolContext,
  mark: (s: number) => void,
  req: Parameters<ToolContext["call"]>[0],
) {
  const res = await ctx.call(req);
  mark(res.status);
  if (res.status < 200 || res.status >= 300) {
    const message = (res.data as { error?: { message?: string } } | null)?.error?.message;
    throw new VendorFailed(
      res.status,
      `Microsoft 365 answered ${res.status}${message ? `: ${message}` : ""}`,
    );
  }
  return res.data as any;
}

const address = (a?: Address) => a?.emailAddress?.address ?? null;
const recipients = (list: string[]) => list.map((address) => ({ emailAddress: { address } }));

function summary(m: Message) {
  return {
    id: m.id,
    subject: m.subject ?? "",
    from: address(m.from),
    to: (m.toRecipients ?? []).map(address),
    received: m.receivedDateTime ?? null,
    read: m.isRead ?? null,
    preview: m.bodyPreview ?? "",
  };
}

export const microsoftTools: Toolset = {
  names: [
    "mail_list_messages",
    "mail_get_message",
    "mail_send",
    "mail_reply",
    "mail_move",
    "mail_flag",
    "mail_mark_read",
    "mail_delete",
  ],

  register(server, ctx, allowed) {
    const readOnly = { readOnlyHint: true, openWorldHint: true } as const;
    const sends = { readOnlyHint: false, destructiveHint: false, openWorldHint: true } as const;
    const emails = z.array(z.email()).max(50);

    if (allowed.has("mail_list_messages"))
      server.registerTool(
        "mail_list_messages",
        {
          description:
            "List recent messages in this agent's mailbox, newest first. A page can hold few or none while more follow: when `more` is true, call again with `cursor` set to `next`.",
          inputSchema: z.object({
            folder: z.enum(["inbox", "sentitems", "drafts", "archive"]).default("inbox"),
            search: z.string().max(200).optional().describe("Words to search for."),
            unread_only: z.boolean().default(false),
            max_results: z.number().int().min(1).max(50).default(20),
            cursor: z
              .string()
              .max(4000)
              .optional()
              .describe(
                "The `next` value from the previous call with the same folder, search and unread_only: the following page.",
              ),
          }),
          annotations: readOnly,
        },
        ({ folder, search, unread_only, max_results, cursor }) =>
          audited(ctx, "mail_list_messages", folder, async (mark) => {
            const query = [`$top=${max_results}`, `$select=${LIST_FIELDS}`, ...pageParam(cursor)];
            // Microsoft does not combine $search with $filter on messages.
            const words = search?.replaceAll('"', "").trim();
            if (words) query.push(`$search=${encodeURIComponent(`"${words}"`)}`);
            else if (unread_only) query.push(`$filter=${encodeURIComponent("isRead eq false")}`);
            const data = await ok(ctx, mark, {
              method: "GET",
              path: `${mailboxPath(ctx)}/mailFolders/${folder}/messages?${query.join("&")}`,
            });
            const messages = (data.value as Message[])
              .filter((m) => !unread_only || m.isRead === false)
              .map(summary);
            return json({ messages, more: Boolean(data["@odata.nextLink"]), next: nextPage(data) });
          }),
      );

    if (allowed.has("mail_get_message"))
      server.registerTool(
        "mail_get_message",
        {
          description: "Read one message in this agent's mailbox, with its body as plain text.",
          inputSchema: z.object({ message_id: z.string().describe("Id from mail_list_messages.") }),
          annotations: readOnly,
        },
        ({ message_id }) =>
          audited(ctx, "mail_get_message", message_id, async (mark) => {
            const m: Message = await ok(ctx, mark, {
              method: "GET",
              path: `${messagePath(ctx, message_id)}?$select=${MESSAGE_FIELDS}`,
              headers: { prefer: 'outlook.body-content-type="text"' },
            });
            return json({
              ...summary(m),
              cc: (m.ccRecipients ?? []).map(address),
              conversation: m.conversationId ?? null,
              attachments: m.hasAttachments ?? false,
              body: m.body?.content ?? "",
            });
          }),
      );

    if (allowed.has("mail_send"))
      server.registerTool(
        "mail_send",
        {
          description: "Send a plain-text email from this agent's mailbox.",
          inputSchema: z.object({
            to: emails.min(1),
            cc: emails.default([]),
            subject: z.string().min(1).max(255),
            body: z.string().min(1).max(100000),
          }),
          annotations: sends,
        },
        ({ to, cc, subject, body }) =>
          audited(ctx, "mail_send", [...to, ...cc].join(", "), async (mark) => {
            await ok(ctx, mark, {
              method: "POST",
              path: `${mailboxPath(ctx)}/sendMail`,
              body: {
                message: {
                  subject,
                  body: { contentType: "Text", content: body },
                  toRecipients: recipients(to),
                  ccRecipients: recipients(cc),
                },
                saveToSentItems: true,
              },
            });
            return json({ sent: true });
          }),
      );

    if (allowed.has("mail_reply"))
      server.registerTool(
        "mail_reply",
        {
          description: "Reply to the sender of a message in this agent's mailbox.",
          inputSchema: z.object({
            message_id: z.string().describe("Id from mail_list_messages."),
            body: z.string().min(1).max(100000),
          }),
          annotations: sends,
        },
        ({ message_id, body }) =>
          audited(ctx, "mail_reply", message_id, async (mark) => {
            await ok(ctx, mark, {
              method: "POST",
              path: `${messagePath(ctx, message_id)}/reply`,
              body: { comment: body },
            });
            return json({ sent: true });
          }),
      );

    const organizes = { readOnlyHint: false, destructiveHint: false, openWorldHint: true } as const;
    const messageId = z.string().describe("Id from mail_list_messages.");

    if (allowed.has("mail_move"))
      server.registerTool(
        "mail_move",
        {
          // Deleted Items and Junk Email (purged automatically) would be deleting, which
          // mail_delete covers. Checked inside audited() so a refusal is logged.
          description: "Move a message in this agent's mailbox to another folder.",
          inputSchema: z.object({
            message_id: messageId,
            folder: z.string().max(50).describe("inbox or archive"),
          }),
          annotations: organizes,
        },
        ({ message_id, folder }) =>
          audited(ctx, "mail_move", `${message_id} → ${folder}`, async (mark) => {
            if (!MOVE_TO.includes(folder))
              throw new Denied("Messages can be moved to inbox or archive");
            const moved = await ok(ctx, mark, {
              method: "POST",
              path: `${messagePath(ctx, message_id)}/move`,
              body: { destinationId: folder },
            });
            // Moving gives the message a new id.
            return json({ moved: true, message_id: moved?.id ?? null });
          }),
      );

    if (allowed.has("mail_flag"))
      server.registerTool(
        "mail_flag",
        {
          description:
            "Flag a message in this agent's mailbox, mark it complete, or clear the flag.",
          inputSchema: z.object({
            message_id: messageId,
            flag: z.enum(["flagged", "complete", "notFlagged"]),
          }),
          annotations: organizes,
        },
        ({ message_id, flag }) =>
          audited(ctx, "mail_flag", `${message_id}: ${flag}`, async (mark) => {
            await ok(ctx, mark, {
              method: "PATCH",
              path: messagePath(ctx, message_id),
              body: { flag: { flagStatus: flag } },
            });
            return json({ flag });
          }),
      );

    if (allowed.has("mail_mark_read"))
      server.registerTool(
        "mail_mark_read",
        {
          description: "Mark a message in this agent's mailbox read or unread.",
          inputSchema: z.object({ message_id: messageId, read: z.boolean().default(true) }),
          annotations: organizes,
        },
        ({ message_id, read }) =>
          audited(
            ctx,
            "mail_mark_read",
            `${message_id}: ${read ? "read" : "unread"}`,
            async (mark) => {
              await ok(ctx, mark, {
                method: "PATCH",
                path: messagePath(ctx, message_id),
                body: { isRead: read },
              });
              return json({ read });
            },
          ),
      );

    if (allowed.has("mail_delete"))
      server.registerTool(
        "mail_delete",
        {
          description:
            "Delete a message in this agent's mailbox: it moves to Deleted Items, where it can be restored until Microsoft 365 empties the folder.",
          inputSchema: z.object({ message_id: messageId }),
          annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
        },
        ({ message_id }) =>
          audited(ctx, "mail_delete", message_id, async (mark) => {
            // A DELETE skips Deleted Items (it goes straight to Recoverable Items, which only
            // an administrator can reach); a move keeps it where people look for it.
            const moved = await ok(ctx, mark, {
              method: "POST",
              path: `${messagePath(ctx, message_id)}/move`,
              body: { destinationId: "deleteditems" },
            });
            return json({ deleted: true, folder: "deleteditems", message_id: moved.id });
          }),
      );
  },
};
