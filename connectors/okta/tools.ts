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
import { cursorArg } from "../sharepoint-files/tools.js";

// People, groups, apps and sign-in activity in Okta, and the help desk's actions on a person.
// A key's `groups` limit (Okta group ids) keeps it to the members of those groups: every tool
// that names a person reads the person's groups first, and every write checks before it calls.
// Okta has no nested groups, so membership is direct. Request shapes: developer.okta.com/docs/
// api/openapi/okta-management/management/tag/User, UserLifecycle, UserSessions, UserFactor,
// Group, Application and SystemLog (read October 27, 2026).

const USER_ID = /^00u[A-Za-z0-9]{17}$/;
const GROUP_ID = /^00g[A-Za-z0-9]{17}$/;
const FACTOR_ID = /^[A-Za-z0-9]{20}$/;
const LOGIN = /^[^\s/\\?#%"]{1,100}$/;
const EVENT_TYPE = /^[a-z0-9_.]{1,100}$/;
const STATUSES = [
  "ACTIVE",
  "LOCKED_OUT",
  "PASSWORD_EXPIRED",
  "RECOVERY",
  "SUSPENDED",
  "PROVISIONED",
  "STAGED",
  "DEPROVISIONED",
] as const;
// A person's groups are read at most this many pages of 200 for the limit check; past it the
// person reads as outside the key's groups (refused, never widened).
export const MAX_GROUP_PAGES = 5;
// Under a groups limit a search reads on, at most this many pages, until it has `max_results`.
export const MAX_SCOPED_PAGES = 5;
// Okta keeps the System Log 90 days.
const MAX_HOURS = 90 * 24;

type Mark = (status: number) => void;
type Obj = Record<string, any>;
type Page = { items: Obj[]; after: string | null };

function oktaError(res: { status: number; data: unknown }) {
  const d = res.data as { errorSummary?: string; errorCauses?: { errorSummary?: string }[] };
  const detail = [d?.errorSummary, ...(d?.errorCauses ?? []).map((c) => c.errorSummary)]
    .filter(Boolean)
    .join("; ");
  return new VendorFailed(res.status, `Okta answered ${res.status}${detail ? `: ${detail}` : ""}`);
}

async function ok(ctx: ToolContext, mark: Mark, req: Parameters<ToolContext["call"]>[0]) {
  const res = await ctx.call(req);
  mark(res.status);
  if (res.status < 200 || res.status >= 300) throw oktaError(res);
  return res.data as any;
}

const page = (ctx: ToolContext, mark: Mark, path: string) =>
  ok(ctx, mark, { method: "GET", path }) as Promise<Page>;

const after = (cursor: string | null | undefined) =>
  cursor ? `&after=${encodeURIComponent(cursor)}` : "";

/** The key's group ids; null when it has none. */
export function groupScope(ctx: ToolContext): Set<string> | null {
  const groups = listLimit(ctx.limits, "groups");
  return groups ? new Set(groups) : null;
}

async function groupsOf(ctx: ToolContext, mark: Mark, userId: string) {
  const groups: Obj[] = [];
  let cursor: string | null = null;
  for (let read = 0; read < MAX_GROUP_PAGES; read++) {
    const data = await page(ctx, mark, `/api/v1/users/${userId}/groups?limit=200${after(cursor)}`);
    groups.push(...data.items);
    cursor = data.after;
    if (!cursor) break;
  }
  return groups;
}

const inScope = (scope: Set<string> | null, groups: Obj[]) =>
  !scope || groups.some((g) => scope.has(g.id));

function userRef(raw: string): string {
  const value = raw.trim();
  if (!USER_ID.test(value) && !LOGIN.test(value))
    throw new Denied("Give a person's Okta id or username, such as ada@acme.com");
  return encodeURIComponent(value);
}

/** The person, read and checked against the key's groups; `groups` only when read. */
async function openUser(ctx: ToolContext, mark: Mark, raw: string, withGroups = false) {
  const scope = groupScope(ctx);
  const denied = new Denied("This agent may not look up or act on that person");
  let user: Obj;
  try {
    user = await ok(ctx, mark, { method: "GET", path: `/api/v1/users/${userRef(raw)}` });
  } catch (error) {
    // A kept key must not learn which usernames exist outside its groups.
    if (scope && error instanceof VendorFailed && error.status === 404) throw denied;
    throw error;
  }
  if (!USER_ID.test(user.id ?? "")) throw new Error("Okta answered a user without an id");
  const groups = scope || withGroups ? await groupsOf(ctx, mark, user.id) : [];
  if (!inScope(scope, groups)) throw denied;
  return { user, groups, scope };
}

const person = (u: Obj) => ({
  id: u.id,
  login: u.profile?.login ?? null,
  email: u.profile?.email ?? null,
  name: [u.profile?.firstName, u.profile?.lastName].filter(Boolean).join(" ") || null,
  title: u.profile?.title ?? null,
  department: u.profile?.department ?? null,
  status: u.status ?? null,
});

const group = (g: Obj) => ({
  id: g.id,
  name: g.profile?.name ?? null,
  description: g.profile?.description ?? null,
  type: g.type ?? null,
});

/** A listing page kept to what `keep` allows; `next` continues after the last page read. */
async function keptPage<T>(
  ctx: ToolContext,
  mark: Mark,
  path: (cursor: string) => string,
  max: number,
  cursor: string | undefined,
  scoped: boolean,
  keep: (o: Obj) => Promise<T | null>,
) {
  const kept: T[] = [];
  let next: string | null = cursor ?? null;
  for (let read = 0; read < (scoped ? MAX_SCOPED_PAGES : 1); read++) {
    const data = await page(ctx, mark, path(after(next)));
    for (const item of await Promise.all(data.items.map(keep))) if (item !== null) kept.push(item);
    next = data.after;
    if (!next || kept.length >= max) break;
  }
  return { kept, more: Boolean(next), next };
}

// Search words go inside Okta's quoted filter strings.
const words = (raw: string) => raw.replace(/["\\]/g, "").trim();

// The log window rides in the cursor, so a later page reads the same window.
const ISO = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d{1,3})?Z$/;
function logWindow(cursor: string | undefined) {
  if (!cursor) return { until: new Date().toISOString(), after: null };
  const split = cursor.indexOf("~");
  const [until, rest] = [cursor.slice(0, split), cursor.slice(split + 1)];
  if (split < 0 || !ISO.test(until) || !rest)
    throw new Denied("Give `next` from the previous call as `cursor`");
  return { until, after: rest };
}

function event(e: Obj, userId: string, scoped: boolean) {
  const actor = e.actor ?? {};
  const ua = e.client?.userAgent ?? {};
  const geo = e.client?.geographicalContext ?? {};
  const known = (v: unknown) => (typeof v === "string" && v && v !== "UNKNOWN" ? v : null);
  return {
    time: e.published ?? null,
    event_type: e.eventType ?? null,
    message: e.displayMessage ?? null,
    outcome: e.outcome?.result ?? null,
    reason: e.outcome?.reason ?? null,
    // Under a groups limit another person's name could be outside the key's groups.
    by:
      actor.id === userId
        ? "this person"
        : scoped && actor.type === "User"
          ? "another person"
          : (actor.displayName ?? null),
    ip: e.client?.ipAddress ?? null,
    location: [geo.city, geo.state, geo.country].filter(known).join(", ") || null,
    client: [known(ua.browser), known(ua.os)].filter(Boolean).join(" on ") || null,
  };
}

export const oktaTools: Toolset = {
  names: [
    "okta_get_user",
    "okta_search_users",
    "okta_list_factors",
    "okta_search_groups",
    "okta_list_group_members",
    "okta_list_apps",
    "okta_get_user_activity",
    "okta_add_group_member",
    "okta_remove_group_member",
    "okta_unlock_user",
    "okta_send_password_reset",
    "okta_reset_factor",
    "okta_suspend_user",
    "okta_unsuspend_user",
    "okta_clear_sessions",
    "okta_deactivate_user",
  ],

  register(server, ctx, allowed) {
    const readOnly = { readOnlyHint: true, openWorldHint: true } as const;
    const write = (destructive: boolean) =>
      ({ readOnlyHint: false, destructiveHint: destructive, openWorldHint: true }) as const;
    const user = z
      .string()
      .max(100)
      .describe("A person's Okta id (00u...) or username, such as ada@acme.com.");
    const groupId = z.string().max(20).describe("A group's id (00g...), from okta_search_groups.");
    const max = (most: number, fallback: number) =>
      z.number().int().min(1).max(most).default(fallback);
    const checkGroup = (scope: Set<string> | null, id: string) => {
      if (!GROUP_ID.test(id)) throw new Denied("Give a group id from okta_search_groups");
      if (scope && !scope.has(id)) throw new Denied("This agent may not use that group");
    };

    if (allowed.has("okta_get_user"))
      server.registerTool(
        "okta_get_user",
        {
          description:
            "Look up one person: username, email, name, title, department, account status (ACTIVE, LOCKED_OUT, SUSPENDED, PASSWORD_EXPIRED, ...), last sign-in, last password change and their groups.",
          inputSchema: z.object({ user }),
          annotations: readOnly,
        },
        ({ user: u }) =>
          audited(ctx, "okta_get_user", u, async (mark) => {
            const { user: p, groups, scope } = await openUser(ctx, mark, u, true);
            const shown = groups.filter((g) => !scope || scope.has(g.id));
            return json({
              ...person(p),
              locked_out: p.status === "LOCKED_OUT",
              last_login: p.lastLogin ?? null,
              password_changed: p.passwordChanged ?? null,
              status_changed: p.statusChanged ?? null,
              created: p.created ?? null,
              groups: shown.slice(0, 100).map(group),
              more_groups: shown.length > 100,
            });
          }),
      );

    if (allowed.has("okta_search_users"))
      server.registerTool(
        "okta_search_users",
        {
          description:
            "Find people whose username, email, first or last name starts with `query`, or everyone with a status (e.g. LOCKED_OUT). When `more` is true, call again with `cursor` set to `next`.",
          inputSchema: z.object({
            query: z.string().max(100).optional(),
            status: z.enum(STATUSES).optional(),
            max_results: max(50, 25),
            cursor: cursorArg,
          }),
          annotations: readOnly,
        },
        ({ query, status, max_results, cursor }) =>
          audited(ctx, "okta_search_users", query ?? status ?? null, async (mark) => {
            const w = words(query ?? "");
            if (!w && !status) throw new Denied("Give a few letters of a name, or a status");
            const parts = [
              ...(w
                ? [
                    `(${["login", "email", "firstName", "lastName"].map((f) => `profile.${f} sw "${w}"`).join(" or ")})`,
                  ]
                : []),
              ...(status ? [`status eq "${status}"`] : []),
            ];
            const search = encodeURIComponent(parts.join(" and "));
            const scope = groupScope(ctx);
            const { kept, more, next } = await keptPage(
              ctx,
              mark,
              (c) => `/api/v1/users?search=${search}&limit=${max_results}${c}`,
              max_results,
              cursor,
              Boolean(scope),
              async (u) =>
                inScope(scope, scope ? await groupsOf(ctx, mark, u.id) : []) ? person(u) : null,
            );
            return json({ people: kept, more, next });
          }),
      );

    if (allowed.has("okta_list_factors"))
      server.registerTool(
        "okta_list_factors",
        {
          description:
            "List a person's enrolled MFA factors (authenticators): id, type, provider and status. Use an id with okta_reset_factor.",
          inputSchema: z.object({ user }),
          annotations: readOnly,
        },
        ({ user: u }) =>
          audited(ctx, "okta_list_factors", u, async (mark) => {
            const { user: p } = await openUser(ctx, mark, u);
            const data = await page(ctx, mark, `/api/v1/users/${p.id}/factors`);
            return json({
              user: person(p),
              factors: data.items.map((f) => ({
                id: f.id,
                type: f.factorType ?? null,
                provider: f.provider ?? null,
                name: f.vendorName ?? null,
                status: f.status ?? null,
                enrolled: f.created ?? null,
                last_updated: f.lastUpdated ?? null,
              })),
            });
          }),
      );

    if (allowed.has("okta_search_groups"))
      server.registerTool(
        "okta_search_groups",
        {
          description:
            "Find groups whose name starts with `query` (all groups without it). When `more` is true, call again with `cursor` set to `next`.",
          inputSchema: z.object({
            query: z.string().max(100).optional(),
            max_results: max(100, 25),
            cursor: cursorArg,
          }),
          annotations: readOnly,
        },
        ({ query, max_results, cursor }) =>
          audited(ctx, "okta_search_groups", query ?? null, async (mark) => {
            const w = words(query ?? "");
            const scope = groupScope(ctx);
            if (scope) {
              // The key's own groups, read directly: scanning every group to find them could miss.
              const found = await Promise.all(
                [...scope].map(async (id) => {
                  const res = await ctx.call({ method: "GET", path: `/api/v1/groups/${id}` });
                  mark(res.status);
                  if (res.status === 404) return null;
                  if (res.status < 200 || res.status >= 300) throw oktaError(res);
                  return res.data as Obj;
                }),
              );
              const groups = found
                .filter((g): g is Obj => g !== null)
                .filter((g) => (g.profile?.name ?? "").toLowerCase().startsWith(w.toLowerCase()))
                .map(group);
              return json({ groups, more: false, next: null });
            }
            const data = await page(
              ctx,
              mark,
              `/api/v1/groups?${w ? `q=${encodeURIComponent(w)}&` : ""}limit=${max_results}${after(cursor)}`,
            );
            return json({
              groups: data.items.map(group),
              more: Boolean(data.after),
              next: data.after,
            });
          }),
      );

    if (allowed.has("okta_list_group_members"))
      server.registerTool(
        "okta_list_group_members",
        {
          description:
            "List the people in a group. When `more` is true, call again with `cursor` set to `next`.",
          inputSchema: z.object({
            group_id: groupId,
            max_results: max(200, 50),
            cursor: cursorArg,
          }),
          annotations: readOnly,
        },
        ({ group_id, max_results, cursor }) =>
          audited(ctx, "okta_list_group_members", group_id, async (mark) => {
            checkGroup(groupScope(ctx), group_id);
            const data = await page(
              ctx,
              mark,
              `/api/v1/groups/${group_id}/users?limit=${max_results}${after(cursor)}`,
            );
            return json({
              people: data.items.map(person),
              more: Boolean(data.after),
              next: data.after,
            });
          }),
      );

    if (allowed.has("okta_list_apps"))
      server.registerTool(
        "okta_list_apps",
        {
          description:
            "List apps in Okta: those assigned to `user` when given, and with a name starting with `query` when given. When `more` is true, call again with `cursor` set to `next`.",
          inputSchema: z.object({
            user: user.optional(),
            query: z.string().max(100).optional(),
            max_results: max(100, 25),
            cursor: cursorArg,
          }),
          annotations: readOnly,
        },
        ({ user: u, query, max_results, cursor }) =>
          audited(ctx, "okta_list_apps", u ?? query ?? null, async (mark) => {
            const w = words(query ?? "");
            const filter = u
              ? `filter=${encodeURIComponent(`user.id eq "${(await openUser(ctx, mark, u)).user.id}"`)}&`
              : "";
            const q = w ? `q=${encodeURIComponent(w)}&` : "";
            const data = await page(
              ctx,
              mark,
              `/api/v1/apps?${filter}${q}limit=${max_results}${after(cursor)}`,
            );
            return json({
              apps: data.items.map((a) => ({
                id: a.id,
                label: a.label ?? null,
                app: a.name ?? null,
                status: a.status ?? null,
                sign_on_mode: a.signOnMode ?? null,
              })),
              more: Boolean(data.after),
              next: data.after,
            });
          }),
      );

    if (allowed.has("okta_get_user_activity"))
      server.registerTool(
        "okta_get_user_activity",
        {
          description:
            "A person's recent System Log events, newest first: sign-ins, MFA challenges, lockouts, password and account changes, with outcome, IP address, place and browser. Also gives their status and last sign-in. `event_type` narrows to one Okta event type, e.g. user.session.start or user.account.lock. When `more` is true, call again with `cursor` set to `next` and the same `hours`.",
          inputSchema: z.object({
            user,
            hours: z.number().int().min(1).max(MAX_HOURS).default(24),
            event_type: z.string().max(100).optional(),
            max_results: max(100, 25),
            cursor: cursorArg,
          }),
          annotations: readOnly,
        },
        ({ user: u, hours, event_type, max_results, cursor }) =>
          audited(ctx, "okta_get_user_activity", u, async (mark) => {
            if (event_type !== undefined && !EVENT_TYPE.test(event_type))
              throw new Denied("Give an Okta event type such as user.session.start");
            const window = logWindow(cursor);
            const { user: p, scope } = await openUser(ctx, mark, u);
            const since = new Date(Date.parse(window.until) - hours * 3600_000).toISOString();
            const filter = `(actor.id eq "${p.id}" or target.id eq "${p.id}")${event_type ? ` and eventType eq "${event_type}"` : ""}`;
            const data = await page(
              ctx,
              mark,
              `/api/v1/logs?since=${since}&until=${window.until}&sortOrder=DESCENDING&filter=${encodeURIComponent(filter)}&limit=${max_results}${after(window.after)}`,
            );
            return json({
              user: person(p),
              last_login: p.lastLogin ?? null,
              events: data.items.map((e) => event(e, p.id, Boolean(scope))),
              more: Boolean(data.after),
              next: data.after ? `${window.until}~${data.after}` : null,
            });
          }),
      );

    const membership = (tool: string, add: boolean) =>
      server.registerTool(
        tool,
        {
          description: add
            ? "Add a person to a group. Membership can give them apps and access."
            : "Remove a person from a group. They lose what the group gave them.",
          inputSchema: z.object({ user, group_id: groupId }),
          annotations: write(false),
        },
        ({ user: u, group_id }) =>
          audited(ctx, tool, `${u} ${add ? "to" : "from"} ${group_id}`, async (mark) => {
            checkGroup(groupScope(ctx), group_id);
            const { user: p } = await openUser(ctx, mark, u);
            await ok(ctx, mark, {
              method: add ? "PUT" : "DELETE",
              path: `/api/v1/groups/${group_id}/users/${p.id}`,
            });
            return json({ user: person(p), group_id, member: add });
          }),
      );
    if (allowed.has("okta_add_group_member")) membership("okta_add_group_member", true);
    if (allowed.has("okta_remove_group_member")) membership("okta_remove_group_member", false);

    // One action on one person: read and checked first, then the call on their id.
    const action = (
      tool: string,
      description: string,
      destructive: boolean,
      path: (id: string) => { method: "POST" | "DELETE"; path: string },
      done: string,
    ) =>
      server.registerTool(
        tool,
        { description, inputSchema: z.object({ user }), annotations: write(destructive) },
        ({ user: u }) =>
          audited(ctx, tool, u, async (mark) => {
            const { user: p } = await openUser(ctx, mark, u);
            await ok(ctx, mark, path(p.id));
            return json({ user: person(p), done });
          }),
      );
    const lifecycle = (id: string, step: string, query = "") => ({
      method: "POST" as const,
      path: `/api/v1/users/${id}/lifecycle/${step}${query}`,
    });

    if (allowed.has("okta_unlock_user"))
      action(
        "okta_unlock_user",
        "Unlock a person whose account is locked out (status LOCKED_OUT) so they can sign in again with their password.",
        false,
        (id) => lifecycle(id, "unlock"),
        "Unlocked",
      );
    if (allowed.has("okta_send_password_reset"))
      action(
        "okta_send_password_reset",
        "Email a person a link to set a new password. Their current password stops working until they do. The link goes only to them.",
        false,
        // Always emailed: without sendEmail Okta would answer the reset link itself.
        (id) => lifecycle(id, "reset_password", "?sendEmail=true"),
        "Okta emailed them a password reset link",
      );
    if (allowed.has("okta_suspend_user"))
      action(
        "okta_suspend_user",
        "Suspend a person: they cannot sign in until unsuspended. Their apps and groups are kept.",
        true,
        (id) => lifecycle(id, "suspend"),
        "Suspended",
      );
    if (allowed.has("okta_unsuspend_user"))
      action(
        "okta_unsuspend_user",
        "Unsuspend a suspended person so they can sign in again.",
        false,
        (id) => lifecycle(id, "unsuspend"),
        "Unsuspended",
      );
    if (allowed.has("okta_clear_sessions"))
      action(
        "okta_clear_sessions",
        "Sign a person out everywhere: ends their Okta sessions and revokes the tokens apps got through Okta.",
        true,
        (id) => ({ method: "DELETE", path: `/api/v1/users/${id}/sessions?oauthTokens=true` }),
        "Signed out everywhere",
      );
    if (allowed.has("okta_deactivate_user"))
      action(
        "okta_deactivate_user",
        "Deactivate a person: they can no longer sign in, and apps set up for it remove their accounts. Only an administrator can reactivate them.",
        true,
        (id) => lifecycle(id, "deactivate", "?sendEmail=false"),
        "Deactivated",
      );

    if (allowed.has("okta_reset_factor"))
      server.registerTool(
        "okta_reset_factor",
        {
          description:
            "Remove one of a person's MFA factors (id from okta_list_factors), e.g. after a lost phone. They set it up again at their next sign-in.",
          inputSchema: z.object({ user, factor_id: z.string().max(20) }),
          annotations: write(true),
        },
        ({ user: u, factor_id }) =>
          audited(ctx, "okta_reset_factor", `${u} factor ${factor_id}`, async (mark) => {
            if (!FACTOR_ID.test(factor_id))
              throw new Denied("Give a factor id from okta_list_factors");
            const { user: p } = await openUser(ctx, mark, u);
            await ok(ctx, mark, {
              method: "DELETE",
              path: `/api/v1/users/${p.id}/factors/${factor_id}`,
            });
            return json({ user: person(p), factor_id, done: "Factor removed" });
          }),
      );
  },
};
