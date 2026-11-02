import { z } from "zod";
import {
  audited,
  Denied,
  json,
  listLimit,
  VendorFailed,
  type ToolContext,
  type Toolset,
} from "../../packages/connector-kit/src/common.js";
import type { VendorRequest } from "../../packages/connector-kit/src/types.js";
import { cursorArg } from "../sharepoint-files/tools.js";
import { gongError } from "./vendor.js";

// Gong calls, transcripts, users and workspaces, read only. A key's limits keep it to calls in
// chosen workspaces and, with `domains` or `crm_accounts`, to calls with chosen customers (a
// participant's email domain or a linked CRM account). Gong cannot filter by either, so every
// call is read through /v2/calls/extensive (parties and CRM context) and checked before it is
// shown, including one asked for by id. Gong pages with records.cursor and has no page size, so
// `next` carries Gong's cursor, how many calls of that page were already gone through and the
// date window. Calls marked private in Gong are never returned. Gong answers 404 for an empty
// listing. Request shapes: Gong API v2 documentation (read October 27, 2026).

const CALL_ID = /^\d{1,25}$/;
const USER_ID = /^\d{1,25}$/;
const WORKSPACE_ID = /^\d{1,25}$/;
const EMAIL = /^[^\s@]{1,64}@[^\s@]{1,190}$/;
// A filtered listing reads at most this many of Gong's pages (100 calls or users each) per call.
export const MAX_PAGES = 5;
// An email lookup reads at most this many pages of users.
export const MAX_USER_PAGES = 20;
const TEXT = 4000;
const NO_CALL = "No call with that id that this agent may read";

type Mark = (status: number) => void;
type Obj = Record<string, any>;
type Cursor = { c: string | null; s: number; f?: string; t?: string };

async function ok(ctx: ToolContext, mark: Mark, req: VendorRequest, empty?: Obj) {
  const res = await ctx.call(req);
  mark(res.status);
  if (res.status === 404 && empty) return empty;
  if (res.status < 200 || res.status >= 300) throw new VendorFailed(res.status, gongError(res));
  return res.data as Obj;
}

const encode = (c: Cursor) => Buffer.from(JSON.stringify(c)).toString("base64url");
function decode(raw: string | undefined): Cursor {
  if (!raw) return { c: null, s: 0 };
  try {
    const c = JSON.parse(Buffer.from(raw, "base64url").toString());
    if ((c.c === null || typeof c.c === "string") && Number.isInteger(c.s) && c.s >= 0) return c;
  } catch {}
  throw new Denied("Give `next` from the previous call as `cursor`");
}

/**
 * Gong's pages from `cursor` on, kept to what `keep` allows, until `max` are kept; `next`
 * resumes at the first item not yet gone through.
 */
async function keptPages<T>(
  cursor: Cursor,
  max: number,
  pages: number,
  read: (gongCursor: string | null) => Promise<{ items: Obj[]; cursor: string | null }>,
  keep: (o: Obj) => T | null,
) {
  const kept: T[] = [];
  let at = cursor.c;
  let skip = cursor.s;
  for (let n = 0; n < pages; n++) {
    const page = await read(at);
    for (let i = skip; i < page.items.length; i++) {
      if (kept.length === max) return { kept, next: { ...cursor, c: at, s: i } };
      const item = keep(page.items[i]);
      if (item !== null) kept.push(item);
    }
    skip = 0;
    if (!page.cursor) return { kept, next: null };
    if (kept.length === max) return { kept, next: { ...cursor, c: page.cursor, s: 0 } };
    at = page.cursor;
  }
  return { kept, next: { ...cursor, c: at, s: 0 } };
}

const domainOf = (email: unknown) =>
  typeof email === "string" ? email.slice(email.lastIndexOf("@") + 1).toLowerCase() : "";
const inDomain = (email: unknown, domain: string) => {
  const d = domainOf(email);
  return d === domain || d.endsWith(`.${domain}`);
};

const crmObjects = (call: Obj): Obj[] =>
  (call.context ?? []).flatMap((s: Obj) =>
    (s.objects ?? []).map((o: Obj) => ({ ...o, system: s.system })),
  );

/** Whether the key's limits reach this call (from /v2/calls/extensive). */
export function reachable(limits: ToolContext["limits"], call: Obj): boolean {
  const meta = call.metaData ?? {};
  if (meta.isPrivate === true) return false;
  const workspaces = listLimit(limits, "workspaces");
  if (workspaces && !workspaces.includes(String(meta.workspaceId))) return false;
  const domains = listLimit(limits, "domains")?.map((d) => d.toLowerCase());
  const accounts = listLimit(limits, "crm_accounts");
  if (!domains && !accounts) return true;
  return (
    (call.parties ?? []).some((p: Obj) => domains?.some((d) => inDomain(p.emailAddress, d))) ||
    crmObjects(call).some(
      (o) => String(o.objectType).toLowerCase() === "account" && accounts?.includes(o.objectId),
    )
  );
}

const seconds = (ms: unknown) => (typeof ms === "number" ? Math.round(ms / 1000) : null);
const clip = (s: unknown, n = TEXT) => (typeof s === "string" ? s.slice(0, n) : null);

const participant = (p: Obj) => ({
  name: p.name ?? null,
  email: p.emailAddress ?? null,
  title: p.title ?? null,
  affiliation: p.affiliation ?? null,
  gong_user_id: p.userId ?? null,
});

function crm(call: Obj, fields: boolean) {
  return crmObjects(call)
    .slice(0, 20)
    .map((o) => ({
      system: o.system ?? null,
      type: o.objectType ?? null,
      id: o.objectId ?? null,
      ...(fields
        ? {
            fields: Object.fromEntries(
              (o.fields ?? [])
                .slice(0, 15)
                .filter((f: Obj) => typeof f.name === "string")
                .map((f: Obj) => [
                  f.name,
                  typeof f.value === "string" ? clip(f.value, 300) : (f.value ?? null),
                ]),
            ),
          }
        : {
            name: (o.fields ?? []).find((f: Obj) => /^name$/i.test(f.name))?.value ?? null,
          }),
    }));
}

const m = (call: Obj): Obj => call.metaData ?? {};

function summary(call: Obj) {
  const m = call.metaData ?? {};
  const parties: Obj[] = call.parties ?? [];
  return {
    id: m.id,
    title: m.title ?? null,
    started: m.started ?? null,
    duration_seconds: m.duration ?? null,
    direction: m.direction ?? null,
    scope: m.scope ?? null,
    workspace_id: m.workspaceId ?? null,
    host_user_id: m.primaryUserId ?? null,
    url: m.url ?? null,
    participants: parties.slice(0, 30).map(participant),
    more_participants: parties.length > 30,
  };
}

function details(call: Obj) {
  const c = call.content ?? {};
  return {
    ...summary(call),
    purpose: m(call).purpose ?? null,
    crm: crm(call, true),
    brief: clip(c.brief),
    key_points: (c.keyPoints ?? []).slice(0, 30).map((k: Obj) => clip(k.text, 500)),
    highlights: (c.highlights ?? []).slice(0, 20).map((h: Obj) => ({
      title: h.title ?? null,
      items: (h.items ?? []).slice(0, 20).map((i: Obj) => clip(i.text, 500)),
    })),
    outline: (c.outline ?? []).slice(0, 30).map((s: Obj) => ({
      section: s.section ?? null,
      start_seconds: s.startTime ?? null,
      items: (s.items ?? []).slice(0, 20).map((i: Obj) => clip(i.text, 500)),
    })),
    topics: (c.topics ?? [])
      .filter((t: Obj) => t.duration > 0)
      .map((t: Obj) => ({ name: t.name, seconds: t.duration })),
    outcome: c.callOutcome?.name ?? null,
  };
}
const extensive = (filter: Obj, content: boolean, cursor: string | null) => ({
  method: "POST" as const,
  path: "/v2/calls/extensive",
  body: {
    ...(cursor ? { cursor } : {}),
    filter,
    contentSelector: {
      context: "Extended",
      exposedFields: {
        parties: true,
        ...(content
          ? {
              content: {
                brief: true,
                outline: true,
                highlights: true,
                keyPoints: true,
                topics: true,
                callOutcome: true,
              },
            }
          : {}),
      },
    },
  },
});

/** One call by id, refused like a missing one when the key's limits do not reach it. */
async function openCall(ctx: ToolContext, mark: Mark, id: string, content: boolean) {
  if (!CALL_ID.test(id)) throw new Denied("Give a call id from gong_list_calls");
  const data = await ok(ctx, mark, extensive({ callIds: [id] }, content, null), { calls: [] });
  const call = (data.calls ?? []).find((c: Obj) => String(c.metaData?.id) === id);
  if (!call) throw new VendorFailed(404, NO_CALL);
  if (!reachable(ctx.limits, call)) throw new Denied(NO_CALL);
  return call as Obj;
}

function day(raw: string | undefined, fallback: Date) {
  if (raw === undefined) return fallback.toISOString();
  const t = Date.parse(raw);
  if (!Number.isFinite(t)) throw new Denied("Give dates as ISO 8601, such as 2026-10-01");
  return new Date(t).toISOString();
}

const user = (u: Obj) => ({
  id: u.id,
  email: u.emailAddress ?? null,
  name: [u.firstName, u.lastName].filter(Boolean).join(" ") || null,
  title: u.title ?? null,
  active: u.active ?? null,
  manager_id: u.managerId ?? null,
});

export const gongTools: Toolset = {
  names: [
    "gong_list_calls",
    "gong_get_call",
    "gong_get_transcript",
    "gong_list_workspaces",
    "gong_list_users",
    "gong_get_user",
  ],

  register(server, ctx, allowed) {
    const readOnly = { readOnlyHint: true, openWorldHint: true } as const;
    const max = (most: number, fallback: number) =>
      z.number().int().min(1).max(most).default(fallback);
    const callId = z.string().max(25).describe("A call's id, from gong_list_calls.");
    const users = (gongCursor: string | null, mark: Mark) =>
      ok(
        ctx,
        mark,
        {
          method: "GET",
          path: `/v2/users${gongCursor ? `?cursor=${encodeURIComponent(gongCursor)}` : ""}`,
        },
        { users: [] },
      ).then((d) => ({
        items: (d.users ?? []) as Obj[],
        cursor: d.records?.cursor ?? null,
      }));

    if (allowed.has("gong_list_calls"))
      server.registerTool(
        "gong_list_calls",
        {
          description:
            "List recorded calls and meetings between `from` and `to` (ISO 8601 dates or times; the last 7 days by default), oldest first: title, start, length, participants (name, email, title, Internal or External) and linked CRM records. `participant` keeps calls with that email address or email domain (e.g. acme.com). When `next` is not null, call again with `cursor` set to it.",
          inputSchema: z.object({
            from: z.string().max(40).optional(),
            to: z.string().max(40).optional(),
            participant: z.string().max(254).optional(),
            workspace_id: z.string().max(25).optional(),
            host_user_id: z.string().max(25).optional(),
            max_results: max(100, 25),
            cursor: cursorArg,
          }),
          annotations: readOnly,
        },
        ({ from, to, participant: who, workspace_id, host_user_id, max_results, cursor }) =>
          audited(ctx, "gong_list_calls", who ?? workspace_id ?? from ?? null, async (mark) => {
            const at = decode(cursor);
            const until = at.t ?? day(to, new Date());
            const since = at.f ?? day(from, new Date(Date.parse(until) - 7 * 86400_000));
            if (since >= until) throw new Denied("`from` must be before `to`");
            const workspaces = listLimit(ctx.limits, "workspaces");
            if (workspace_id !== undefined && !WORKSPACE_ID.test(workspace_id))
              throw new Denied("Give a workspace id from gong_list_workspaces");
            if (workspace_id && workspaces && !workspaces.includes(workspace_id))
              throw new Denied("This agent may not read that workspace");
            if (host_user_id !== undefined && !USER_ID.test(host_user_id))
              throw new Denied("Give a Gong user id from gong_list_users");
            const needle = who?.trim().toLowerCase();
            const filter = {
              fromDateTime: since,
              toDateTime: until,
              ...(workspace_id || workspaces?.length === 1
                ? { workspaceId: workspace_id ?? workspaces![0] }
                : {}),
              ...(host_user_id ? { primaryUserIds: [host_user_id] } : {}),
            };
            const { kept, next } = await keptPages(
              { ...at, f: since, t: until },
              max_results,
              MAX_PAGES,
              async (c) => {
                const d = await ok(ctx, mark, extensive(filter, false, c), {
                  calls: [],
                });
                return {
                  items: d.calls ?? [],
                  cursor: d.records?.cursor ?? null,
                };
              },
              (call) =>
                reachable(ctx.limits, call) &&
                (!needle ||
                  (call.parties ?? []).some((p: Obj) =>
                    needle.includes("@")
                      ? String(p.emailAddress ?? "").toLowerCase() === needle
                      : inDomain(p.emailAddress, needle),
                  ))
                  ? { ...summary(call), crm: crm(call, false) }
                  : null,
            );
            return json({
              from: since,
              to: until,
              calls: kept,
              next: next && encode(next),
            });
          }),
      );

    if (allowed.has("gong_get_call"))
      server.registerTool(
        "gong_get_call",
        {
          description:
            "Read one call: title, start, length, participants (Internal or External), the linked CRM account and opportunity with their fields, and what Gong made of it: brief, key points, highlights (such as next steps), outline, topics and outcome. For what was actually said, use gong_get_transcript.",
          inputSchema: z.object({ call_id: callId }),
          annotations: readOnly,
        },
        ({ call_id }) =>
          audited(ctx, "gong_get_call", call_id, async (mark) =>
            json(details(await openCall(ctx, mark, call_id, true))),
          ),
      );

    if (allowed.has("gong_get_transcript"))
      server.registerTool(
        "gong_get_transcript",
        {
          description:
            "Read what was said on a call, in order, one entry per turn: speaker, Internal or External, start and end in seconds from the call's start, and the words. Narrow it with `from_seconds` and `to_seconds`, or with `contains` (only sentences with that phrase, any case). When `next` is not null, call again with `cursor` set to it and the same arguments.",
          inputSchema: z.object({
            call_id: callId,
            from_seconds: z.number().int().min(0).optional(),
            to_seconds: z.number().int().min(0).optional(),
            contains: z.string().min(1).max(200).optional(),
            max_turns: max(200, 50),
            cursor: cursorArg,
          }),
          annotations: readOnly,
        },
        ({ call_id, from_seconds, to_seconds, contains, max_turns, cursor }) =>
          audited(ctx, "gong_get_transcript", call_id, async (mark) => {
            const offset = cursor === undefined ? 0 : Number(cursor);
            if (!Number.isInteger(offset) || offset < 0)
              throw new Denied("Give `next` from the previous call as `cursor`");
            const call = await openCall(ctx, mark, call_id, false);
            const data = await ok(
              ctx,
              mark,
              {
                method: "POST",
                path: "/v2/calls/transcript",
                body: { filter: { callIds: [call_id] } },
              },
              { callTranscripts: [] },
            );
            const speakers = new Map<string, Obj>(
              (call.parties ?? []).map((p: Obj) => [String(p.speakerId), p]),
            );
            const lo = (from_seconds ?? 0) * 1000;
            const hi = to_seconds === undefined ? Infinity : to_seconds * 1000;
            const phrase = contains?.toLowerCase();
            const monologues: Obj[] =
              (data.callTranscripts ?? []).find((t: Obj) => String(t.callId) === call_id)
                ?.transcript ?? [];
            const turns = monologues.flatMap((mono) => {
              const said = (mono.sentences ?? []).filter(
                (s: Obj) =>
                  s.start >= lo &&
                  s.start <= hi &&
                  (!phrase || String(s.text).toLowerCase().includes(phrase)),
              );
              if (!said.length) return [];
              const who = speakers.get(String(mono.speakerId));
              return [
                {
                  speaker: who?.name ?? "Unknown speaker",
                  affiliation: who?.affiliation ?? null,
                  start_seconds: seconds(said[0].start),
                  end_seconds: seconds(said.at(-1).end),
                  text: clip(said.map((s: Obj) => s.text).join(" ")),
                },
              ];
            });
            const end = offset + max_turns;
            return json({
              call: {
                id: call_id,
                title: m(call).title ?? null,
                started: m(call).started ?? null,
              },
              turns: turns.slice(offset, end),
              total_turns: turns.length,
              next: end < turns.length ? String(end) : null,
            });
          }),
      );

    if (allowed.has("gong_list_workspaces"))
      server.registerTool(
        "gong_list_workspaces",
        {
          description: "List the Gong workspaces this agent may read calls in: id and name.",
          inputSchema: z.object({}),
          annotations: readOnly,
        },
        () =>
          audited(ctx, "gong_list_workspaces", null, async (mark) => {
            const scope = listLimit(ctx.limits, "workspaces");
            const data = await ok(ctx, mark, {
              method: "GET",
              path: "/v2/workspaces",
            });
            return json({
              workspaces: (data.workspaces ?? [])
                .filter((w: Obj) => !scope || scope.includes(String(w.id)))
                .map((w: Obj) => ({
                  id: w.id,
                  name: w.name ?? null,
                  description: w.description ?? null,
                })),
            });
          }),
      );

    if (allowed.has("gong_list_users"))
      server.registerTool(
        "gong_list_users",
        {
          description:
            "List the people with Gong accounts in your company (your own staff, not customers): id, email, name, title, whether active, and manager's id. `query` keeps those whose name or email contains it. When `next` is not null, call again with `cursor` set to it and the same `query`.",
          inputSchema: z.object({
            query: z.string().max(100).optional(),
            max_results: max(100, 25),
            cursor: cursorArg,
          }),
          annotations: readOnly,
        },
        ({ query, max_results, cursor }) =>
          audited(ctx, "gong_list_users", query ?? null, async (mark) => {
            const q = query?.trim().toLowerCase();
            const { kept, next } = await keptPages(
              decode(cursor),
              max_results,
              q ? MAX_PAGES : 1,
              (c) => users(c, mark),
              (u) => {
                const shown = user(u);
                return !q || `${shown.name ?? ""} ${shown.email ?? ""}`.toLowerCase().includes(q)
                  ? shown
                  : null;
              },
            );
            return json({ users: kept, next: next && encode(next) });
          }),
      );

    if (allowed.has("gong_get_user"))
      server.registerTool(
        "gong_get_user",
        {
          description:
            "Look up one person with a Gong account by Gong user id (e.g. a call's host_user_id) or email address.",
          inputSchema: z.object({
            user: z.string().max(254).describe("A Gong user id or an email address."),
          }),
          annotations: readOnly,
        },
        ({ user: ref }) =>
          audited(ctx, "gong_get_user", ref, async (mark) => {
            const value = ref.trim().toLowerCase();
            if (USER_ID.test(value)) {
              const data = await ok(ctx, mark, {
                method: "GET",
                path: `/v2/users/${value}`,
              });
              return json(user(data.user ?? {}));
            }
            if (!EMAIL.test(value)) throw new Denied("Give a Gong user id or an email address");
            const { kept } = await keptPages(
              { c: null, s: 0 },
              1,
              MAX_USER_PAGES,
              (c) => users(c, mark),
              (u) =>
                [u.emailAddress, ...(u.emailAliases ?? [])].some(
                  (e) => String(e ?? "").toLowerCase() === value,
                )
                  ? user(u)
                  : null,
            );
            if (!kept.length) throw new VendorFailed(404, "No Gong user with that email address");
            return json(kept[0]);
          }),
      );
  },
};
