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
import { cursorArg, nextCursor, ok, pageQuery } from "../sharepoint-files/tools.js";

// People and groups in Microsoft Entra ID, read-only. A key's `groups` and `units` limits
// (group and administrative unit ids) keep it to the people and groups inside them, checked on
// every call, lookups by id included, with checkMemberObjects (transitive). Request shapes:
// learn.microsoft.com/graph/api/user-get, user-list, user-list-transitivememberof,
// user-list-directreports, group-list, group-list-members, subscribedsku-list and
// directoryobject-checkmemberobjects.

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const UPN = /^[^\s/\\?#%'"]+@[^\s/\\?#%'"]+$/;
// checkMemberObjects takes at most 20 ids.
export const MAX_SCOPE_IDS = 20;
const PERSON_FIELDS =
  "id,displayName,userPrincipalName,mail,jobTitle,department,officeLocation,accountEnabled";
const BRIEF_FIELDS = "id,displayName,userPrincipalName,mail,jobTitle,department";
const GROUP_FIELDS = "id,displayName,mail,description,securityEnabled,groupTypes";
const MEMBER_FIELDS = [...new Set(`${BRIEF_FIELDS},${GROUP_FIELDS}`.split(","))].join(",");
// Directory searches and casts need this header (and $count).
const EVENTUAL = { consistencylevel: "eventual" };

type Mark = (status: number) => void;
type DirObject = Record<string, unknown> & { id: string; "@odata.type"?: string };

/** The key's group and administrative unit ids, lowercased; null when it has neither. */
export function directoryScope(ctx: ToolContext) {
  const groups = (listLimit(ctx.limits, "groups") ?? []).map((g) => g.toLowerCase());
  const units = (listLimit(ctx.limits, "units") ?? []).map((u) => u.toLowerCase());
  if (!groups.length && !units.length) return null;
  if (groups.length + units.length > MAX_SCOPE_IDS)
    throw new Denied(
      `This key names more than ${MAX_SCOPE_IDS} groups and administrative units. Ask an administrator to narrow it.`,
    );
  return { groups, units, all: [...groups, ...units] };
}

/** Which of `ids` the user or group is in, directly or through other groups. */
export async function memberOf(
  ctx: ToolContext,
  mark: Mark,
  kind: "users" | "groups",
  id: string,
  ids: string[],
): Promise<string[]> {
  const data = await ok(ctx, mark, {
    method: "POST",
    path: `/v1.0/${kind}/${encodeURIComponent(id)}/checkMemberObjects`,
    body: { ids },
  });
  return ((data?.value ?? []) as string[]).map((v) => v.toLowerCase());
}

type Scope = NonNullable<ReturnType<typeof directoryScope>>;

async function personAllowed(ctx: ToolContext, mark: Mark, scope: Scope | null, id: string) {
  return !scope || (await memberOf(ctx, mark, "users", id, scope.all)).length > 0;
}

// "all": the group, and so everyone in it, is inside the key's groups. "checked": it is only
// in one of the key's administrative units, so each member is checked on its own.
async function groupAccess(
  ctx: ToolContext,
  mark: Mark,
  scope: Scope | null,
  id: string,
): Promise<"all" | "checked" | null> {
  if (!scope || scope.groups.includes(id.toLowerCase())) return "all";
  const hits = await memberOf(ctx, mark, "groups", id, scope.all);
  if (hits.some((h) => scope.groups.includes(h))) return "all";
  return hits.length ? "checked" : null;
}

const kindOf = (o: DirObject) =>
  o["@odata.type"] === "#microsoft.graph.user"
    ? "users"
    : o["@odata.type"] === "#microsoft.graph.group"
      ? "groups"
      : null;

const brief = (u: Record<string, unknown>) => ({
  id: u.id,
  name: u.displayName ?? null,
  email: u.mail ?? u.userPrincipalName ?? null,
  title: u.jobTitle ?? null,
  department: u.department ?? null,
});

const groupBrief = (g: Record<string, unknown>) => ({
  id: g.id,
  name: g.displayName ?? null,
  email: g.mail ?? null,
  description: g.description ?? null,
  kind: (g.groupTypes as string[] | undefined)?.includes("Unified")
    ? "Microsoft 365"
    : g.securityEnabled
      ? "security"
      : "distribution",
});

/** A person given as an object id or a sign-in name, as a Graph path segment. */
function personRef(raw: string): string {
  const value = raw.trim();
  if (!GUID.test(value) && !UPN.test(value))
    throw new Denied("Give a person's id or sign-in name, such as ada@contoso.com");
  return encodeURIComponent(value);
}

const searchWords = (raw: string) => {
  const words = raw.replace(/["\\]/g, "").trim();
  if (!words) throw new Denied("Give a few letters of a name to search for");
  return words;
};

/** The person, read and checked against the key's limits. */
async function openPerson(ctx: ToolContext, mark: Mark, raw: string, query: string) {
  const scope = directoryScope(ctx);
  const denied = new Denied("This agent may not look up that person");
  let person;
  try {
    person = await ok(ctx, mark, { method: "GET", path: `/v1.0/users/${personRef(raw)}?${query}` });
  } catch (error) {
    // A kept key must not learn which sign-in names exist outside its groups.
    if (scope && error instanceof VendorFailed && error.status === 404) throw denied;
    throw error;
  }
  if (!(await personAllowed(ctx, mark, scope, person.id))) throw denied;
  return person;
}

const page = z.number().int().min(1);

// Under limits, filtering follows Graph's paging, so one page can keep nothing. A kept
// listing reads on, at most this many pages, until it has `max` or the listing ends.
export const MAX_SCOPED_PAGES = 5;

/** A listing page kept to what `keep` allows; `next` continues after the last page read. */
async function keptPage<T>(
  ctx: ToolContext,
  mark: Mark,
  path: (cursor: string | undefined) => string,
  max: number,
  cursor: string | undefined,
  scoped: boolean,
  keep: (o: DirObject) => Promise<T | null>,
) {
  const kept: T[] = [];
  let next = cursor;
  for (let read = 0; read < (scoped ? MAX_SCOPED_PAGES : 1); read++) {
    const data = await ok(ctx, mark, { method: "GET", path: path(next), headers: EVENTUAL });
    for (const item of await Promise.all((data.value as DirObject[]).map(keep)))
      if (item !== null) kept.push(item);
    next = nextCursor(data) ?? undefined;
    if (!next || kept.length >= max) break;
  }
  return { kept, more: Boolean(next), next: next ?? null };
}

export const entraTools: Toolset = {
  names: [
    "entra_get_person",
    "entra_search_people",
    "entra_person_groups",
    "entra_group_members",
    "entra_search_groups",
  ],

  register(server, ctx, allowed) {
    const readOnly = { readOnlyHint: true, openWorldHint: true } as const;
    const person = z
      .string()
      .max(320)
      .describe("A person's id or sign-in name, such as ada@contoso.com.");
    const groupId = z.string().max(36).describe("A group's id, from entra_search_groups.");

    if (allowed.has("entra_get_person"))
      server.registerTool(
        "entra_get_person",
        {
          description:
            "Look up one person: name, sign-in name, email, job title, department, office, whether the account is on, manager, direct reports and licenses (by product name).",
          inputSchema: z.object({ person }),
          annotations: readOnly,
        },
        ({ person: p }) =>
          audited(ctx, "entra_get_person", p, async (mark) => {
            const scope = directoryScope(ctx);
            const u = await openPerson(
              ctx,
              mark,
              p,
              `$select=${PERSON_FIELDS},assignedLicenses&$expand=manager($select=${BRIEF_FIELDS})`,
            );
            const reports = await ok(ctx, mark, {
              method: "GET",
              path: `/v1.0/users/${encodeURIComponent(u.id)}/directReports/microsoft.graph.user?$select=${BRIEF_FIELDS}&$top=100&$count=true`,
              headers: EVENTUAL,
            });
            const shown = [];
            for (const r of reports.value as DirObject[])
              if (await personAllowed(ctx, mark, scope, r.id)) shown.push(brief(r));
            const manager = u.manager as DirObject | undefined;
            const skuIds = ((u.assignedLicenses ?? []) as { skuId: string }[]).map((l) => l.skuId);
            let licenses: string[] = [];
            if (skuIds.length) {
              const skus = await ok(ctx, mark, {
                method: "GET",
                path: "/v1.0/subscribedSkus?$select=skuId,skuPartNumber",
              });
              const names = new Map(
                (skus.value as { skuId: string; skuPartNumber: string }[]).map((s) => [
                  s.skuId,
                  s.skuPartNumber,
                ]),
              );
              licenses = skuIds.map((id) => names.get(id) ?? id);
            }
            return json({
              ...brief(u),
              sign_in_name: u.userPrincipalName ?? null,
              office: u.officeLocation ?? null,
              account_enabled: u.accountEnabled ?? null,
              manager:
                manager?.id && (await personAllowed(ctx, mark, scope, manager.id))
                  ? brief(manager)
                  : null,
              direct_reports: shown,
              more_direct_reports: Boolean(reports["@odata.nextLink"]),
              licenses,
            });
          }),
      );

    if (allowed.has("entra_search_people"))
      server.registerTool(
        "entra_search_people",
        {
          description:
            "Find people by part of their name or email. When `more` is true, call again with `cursor` set to `next`.",
          inputSchema: z.object({
            query: z.string().max(100),
            max_results: page.max(50).default(25),
            cursor: cursorArg,
          }),
          annotations: readOnly,
        },
        ({ query, max_results, cursor }) =>
          audited(ctx, "entra_search_people", query, async (mark) => {
            const words = searchWords(query);
            const scope = directoryScope(ctx);
            const search = encodeURIComponent(`"displayName:${words}" OR "mail:${words}"`);
            const { kept, more, next } = await keptPage(
              ctx,
              mark,
              (c) =>
                `/v1.0/users?$search=${search}&$select=${BRIEF_FIELDS}&$count=true&${pageQuery(max_results, c)}`,
              max_results,
              cursor,
              Boolean(scope),
              async (u) => ((await personAllowed(ctx, mark, scope, u.id)) ? brief(u) : null),
            );
            return json({ people: kept, more, next });
          }),
      );

    if (allowed.has("entra_person_groups"))
      server.registerTool(
        "entra_person_groups",
        {
          description:
            "List the groups a person is in, directly or through other groups. When `more` is true, call again with `cursor` set to `next`.",
          inputSchema: z.object({
            person,
            max_results: page.max(100).default(50),
            cursor: cursorArg,
          }),
          annotations: readOnly,
        },
        ({ person: p, max_results, cursor }) =>
          audited(ctx, "entra_person_groups", p, async (mark) => {
            const scope = directoryScope(ctx);
            const u = await openPerson(ctx, mark, p, "$select=id");
            const { kept, more, next } = await keptPage(
              ctx,
              mark,
              (c) =>
                `/v1.0/users/${encodeURIComponent(u.id)}/transitiveMemberOf/microsoft.graph.group?$select=${GROUP_FIELDS}&$count=true&${pageQuery(max_results, c)}`,
              max_results,
              cursor,
              Boolean(scope),
              async (g) => ((await groupAccess(ctx, mark, scope, g.id)) ? groupBrief(g) : null),
            );
            return json({ groups: kept, more, next });
          }),
      );

    if (allowed.has("entra_group_members"))
      server.registerTool(
        "entra_group_members",
        {
          description:
            "List a group's direct members: people and groups inside it. When `more` is true, call again with `cursor` set to `next`.",
          inputSchema: z.object({
            group_id: groupId,
            max_results: page.max(100).default(50),
            cursor: cursorArg,
          }),
          annotations: readOnly,
        },
        ({ group_id, max_results, cursor }) =>
          audited(ctx, "entra_group_members", group_id, async (mark) => {
            if (!GUID.test(group_id)) throw new Denied("Give a group id from entra_search_groups");
            const scope = directoryScope(ctx);
            const access = await groupAccess(ctx, mark, scope, group_id);
            if (!access) throw new Denied("This agent may not read that group");
            const { kept, more, next } = await keptPage(
              ctx,
              mark,
              (c) =>
                `/v1.0/groups/${group_id}/members?$select=${MEMBER_FIELDS}&$count=true&${pageQuery(max_results, c)}`,
              max_results,
              cursor,
              access === "checked",
              async (m) => {
                const kind = kindOf(m);
                if (!kind) return null;
                if (
                  access === "checked" &&
                  !(await memberOf(ctx, mark, kind, m.id, scope!.all)).length
                )
                  return null;
                return kind === "users"
                  ? { type: "person", ...brief(m) }
                  : { type: "group", ...groupBrief(m) };
              },
            );
            return json({ members: kept, more, next });
          }),
      );

    if (allowed.has("entra_search_groups"))
      server.registerTool(
        "entra_search_groups",
        {
          description:
            "Find groups by part of their name. When `more` is true, call again with `cursor` set to `next`.",
          inputSchema: z.object({
            query: z.string().max(100),
            max_results: page.max(50).default(25),
            cursor: cursorArg,
          }),
          annotations: readOnly,
        },
        ({ query, max_results, cursor }) =>
          audited(ctx, "entra_search_groups", query, async (mark) => {
            const words = searchWords(query);
            const scope = directoryScope(ctx);
            const search = encodeURIComponent(`"displayName:${words}"`);
            const { kept, more, next } = await keptPage(
              ctx,
              mark,
              (c) =>
                `/v1.0/groups?$search=${search}&$select=${GROUP_FIELDS}&$count=true&${pageQuery(max_results, c)}`,
              max_results,
              cursor,
              Boolean(scope),
              async (g) => ((await groupAccess(ctx, mark, scope, g.id)) ? groupBrief(g) : null),
            );
            return json({ groups: kept, more, next });
          }),
      );
  },
};
