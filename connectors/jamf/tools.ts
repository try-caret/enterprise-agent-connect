import { z } from "zod";
import { vendorError } from "../../packages/connector-kit/src/oauth.js";
import {
  audited,
  Denied,
  json,
  listLimit,
  VendorFailed,
  type ToolContext,
  type Toolset,
} from "../../packages/connector-kit/src/common.js";

// Jamf Pro computers and mobile devices. A key's `sites` limit keeps it to devices (and groups
// and policies) of those Jamf sites, checked on every call, lookups by id included; a device
// with no site is outside every sites limit. Request shapes are Jamf's OpenAPI files for Jamf
// Pro 11.33 (developer.jamf.com/jamf-pro/reference, read October 27, 2026): GET
// /v4/computers-inventory and /v2/mobile-devices/detail (RSQL `filter`, `section`, `page`,
// `page-size`; neither can filter by site, so a sites limit filters here),
// /v3/computer-groups/{static,smart}-groups, /v2/mobile-device-groups/{static,smart}-groups,
// POST /v2/mdm/commands and /v2/mdm/blank-push. The Classic API where the Jamf Pro API has no
// equal: group members and static membership changes (computergroups and mobiledevicegroups by
// id; PUT with *_additions / *_deletions changes only those devices, where the Jamf Pro API's
// computer group PUT replaces the whole list), policies and a computer's policy logs
// (computerhistory subset PolicyLogs).

type Mark = (status: number) => void;
type Kind = "computer" | "mobile";
type Computer = {
  id: string;
  general?: {
    name?: string;
    lastContact?: string;
    reportDate?: string;
    lastEnrolledDate?: string;
    supervised?: boolean;
    managementId?: string;
    remoteManagement?: { managed?: boolean };
    site?: { id?: string; name?: string };
  };
  hardware?: { serialNumber?: string; model?: string };
  operatingSystem?: { name?: string; version?: string; fileVault2Status?: string };
  userAndLocation?: { username?: string; realname?: string; email?: string };
  groupMemberships?: { groupId?: string; groupName?: string; smartGroup?: boolean }[];
  applications?: { name?: string; version?: string; bundleId?: string }[];
};
type Mobile = {
  mobileDeviceId: string;
  deviceType?: string;
  general?: {
    displayName?: string;
    siteId?: string;
    osVersion?: string;
    lastContactDate?: string;
    lastInventoryUpdateDate?: string;
    lastEnrolledDate?: string;
    managed?: boolean;
    supervised?: boolean;
    managementId?: string;
  };
  hardware?: { serialNumber?: string; model?: string };
  userAndLocation?: { username?: string; realName?: string; emailAddress?: string };
  security?: { dataProtected?: boolean; passcodePresent?: boolean; lostModeEnabled?: boolean };
  groups?: { groupId?: string; groupName?: string; smart?: boolean }[];
  applications?: { name?: string; version?: string; shortVersion?: string; identifier?: string }[];
};
type Device = Computer | Mobile;
type Page<T> = { totalCount?: number; results?: T[] };
type ClassicGroup = {
  id: number;
  name?: string;
  is_smart?: boolean;
  site?: { id?: number; name?: string };
  computers?: { id: number }[];
  mobile_devices?: { id: number }[];
};
type Policy = {
  general?: {
    id?: number;
    name?: string;
    enabled?: boolean;
    frequency?: string;
    trigger_other?: string;
    category?: { name?: string };
    site?: { id?: number; name?: string };
  } & Record<string, unknown>;
  scope?: {
    all_computers?: boolean;
    computers?: unknown[];
    computer_groups?: { name?: string }[];
    exclusions?: { computer_groups?: { name?: string }[] };
  };
  self_service?: { use_for_self_service?: boolean };
  package_configuration?: { packages?: { name?: string }[] };
  scripts?: { name?: string }[];
};

const ID = /^[0-9]{1,10}$/;
// Letters, digits, spaces and . _ @ + ' -: nothing that could end an RSQL value or add a clause.
const SEARCH = /^[\p{L}\p{N} ._@+'-]{1,100}$/u;
const PIN = /^[0-9]{6}$/;
const OFFSET = /^[0-9]{1,6}$/;
// Under a sites limit a listing reads on, at most this many pages, until it has max_results.
export const MAX_SCOPED_PAGES = 5;
// ponytail: under a sites limit each listed policy is read for its site; at most this many per
// call. A site-aware policy listing if Jamf adds one.
export const MAX_POLICY_READS = 50;
const MAX_APPS = 500;

const LABEL: Record<Kind, string> = { computer: "computer", mobile: "mobile device" };

const INVENTORY = {
  computer: {
    path: "/api/v4/computers-inventory",
    id: "id",
    sort: "general.name:asc",
    sections: ["GENERAL", "HARDWARE", "OPERATING_SYSTEM", "USER_AND_LOCATION"],
    groups: "GROUP_MEMBERSHIPS",
    search: (s: string) =>
      `general.name=="*${s}*",hardware.serialNumber=="${s}",userAndLocation.username=="${s}",userAndLocation.email=="${s}"`,
  },
  mobile: {
    path: "/api/v2/mobile-devices/detail",
    id: "mobileDeviceId",
    sort: "displayName:asc",
    sections: ["GENERAL", "HARDWARE", "USER_AND_LOCATION", "SECURITY"],
    groups: "GROUPS",
    search: (s: string) =>
      `displayName=="*${s}*",serialNumber=="${s}",username=="${s}",emailAddress=="${s}"`,
  },
} as const;

const GROUP_LISTS = {
  computer: {
    static: "/api/v3/computer-groups/static-groups",
    smart: "/api/v3/computer-groups/smart-groups",
    name: "name",
  },
  mobile: {
    static: "/api/v2/mobile-device-groups/static-groups",
    smart: "/api/v2/mobile-device-groups/smart-groups",
    name: "groupName",
  },
} as const;

const CLASSIC_GROUP = {
  computer: {
    path: "/JSSResource/computergroups/id",
    root: "computer_group",
    members: "computers",
    item: "computer",
  },
  mobile: {
    path: "/JSSResource/mobiledevicegroups/id",
    root: "mobile_device_group",
    members: "mobile_devices",
    item: "mobile_device",
  },
} as const;

/** Jamf Pro API errors are {httpStatus, errors[{code, description}]}; the Classic API answers
 *  an HTML page. */
export function jamfError(data: unknown): string {
  const errors = (data as { errors?: { code?: string; description?: string }[] } | null)?.errors;
  if (Array.isArray(errors))
    return errors
      .map((e) => e.description || e.code)
      .filter(Boolean)
      .join("; ")
      .slice(0, 300);
  if (typeof data === "string")
    return data
      .replace(/<[^>]*>/g, " ")
      .replace(/\s+/g, " ")
      .trim()
      .slice(0, 300);
  return vendorError(data);
}

async function ok<T>(
  ctx: ToolContext,
  mark: Mark,
  req: Parameters<ToolContext["call"]>[0],
  hintTool?: string,
): Promise<T> {
  const res = await ctx.call(req);
  mark(res.status);
  if (res.status < 200 || res.status >= 300) {
    const detail = jamfError(res.data);
    throw new VendorFailed(
      res.status,
      `Jamf Pro answered ${res.status}${detail ? `: ${detail}` : ""}`,
      hintTool,
    );
  }
  return res.data as T;
}

function keySites(ctx: ToolContext): Set<string> | null {
  const sites = listLimit(ctx.limits, "sites");
  return sites ? new Set(sites) : null;
}

const inSites = (sites: Set<string> | null, site: unknown) =>
  !sites || (site !== undefined && site !== null && sites.has(String(site)));

const siteOf = (kind: Kind, d: Device) =>
  kind === "computer" ? (d as Computer).general?.site?.id : (d as Mobile).general?.siteId;

const deviceId = (kind: Kind, d: Device) =>
  String(kind === "computer" ? (d as Computer).id : (d as Mobile).mobileDeviceId);

function summary(kind: Kind, d: Device) {
  if (kind === "computer") {
    const c = d as Computer;
    return {
      id: c.id,
      kind,
      name: c.general?.name ?? null,
      serial_number: c.hardware?.serialNumber ?? null,
      model: c.hardware?.model ?? null,
      os: c.operatingSystem?.name ?? null,
      os_version: c.operatingSystem?.version ?? null,
      user: c.userAndLocation?.username || null,
      user_name: c.userAndLocation?.realname || null,
      email: c.userAndLocation?.email || null,
      last_contact: c.general?.lastContact ?? null,
      last_inventory: c.general?.reportDate ?? null,
      managed: c.general?.remoteManagement?.managed ?? null,
      supervised: c.general?.supervised ?? null,
      filevault: c.operatingSystem?.fileVault2Status ?? null,
      site_id: c.general?.site?.id ?? null,
    };
  }
  const m = d as Mobile;
  return {
    id: m.mobileDeviceId,
    kind,
    type: m.deviceType ?? null,
    name: m.general?.displayName ?? null,
    serial_number: m.hardware?.serialNumber ?? null,
    model: m.hardware?.model ?? null,
    os_version: m.general?.osVersion ?? null,
    user: m.userAndLocation?.username || null,
    user_name: m.userAndLocation?.realName || null,
    email: m.userAndLocation?.emailAddress || null,
    last_contact: m.general?.lastContactDate ?? null,
    last_inventory: m.general?.lastInventoryUpdateDate ?? null,
    managed: m.general?.managed ?? null,
    supervised: m.general?.supervised ?? null,
    encrypted: m.security?.dataProtected ?? null,
    passcode: m.security?.passcodePresent ?? null,
    lost_mode: m.security?.lostModeEnabled ?? null,
    site_id: m.general?.siteId ?? null,
  };
}

function inventoryPath(
  kind: Kind,
  sections: readonly string[],
  page: number,
  size: number,
  filter?: string,
) {
  const q = new URLSearchParams();
  for (const s of sections) q.append("section", s);
  q.set("page", String(page));
  q.set("page-size", String(size));
  q.set("sort", INVENTORY[kind].sort);
  if (filter) q.set("filter", filter);
  return `${INVENTORY[kind].path}?${q}`;
}

/** Pages of a Jamf Pro API listing from `page`, read on (under a sites limit) until `want`
 *  rows are kept or MAX_SCOPED_PAGES were read. */
async function scan<T>(
  ctx: ToolContext,
  mark: Mark,
  pathFor: (page: number) => string,
  keep: (row: T) => boolean,
  want: number,
  page: number,
) {
  const rows: T[] = [];
  let total = 0;
  for (let read = 0; read < MAX_SCOPED_PAGES; read++) {
    const data = await ok<Page<T>>(ctx, mark, { method: "GET", path: pathFor(page) });
    const results = data.results ?? [];
    total = data.totalCount ?? 0;
    rows.push(...results.filter(keep));
    page++;
    if (rows.length >= want || page * want >= total || !results.length) break;
  }
  return { rows, total, next: page * want < total ? String(page) : null };
}

/** The devices with these ids, by id; ids Jamf does not know are missing. */
async function readDevices(
  ctx: ToolContext,
  mark: Mark,
  kind: Kind,
  ids: string[],
  sections: readonly string[] = INVENTORY[kind].sections,
  hintTool?: string,
): Promise<Map<string, Device>> {
  if (!ids.length) return new Map();
  const data = await ok<Page<Device>>(
    ctx,
    mark,
    {
      method: "GET",
      path: inventoryPath(
        kind,
        sections,
        0,
        ids.length,
        `${INVENTORY[kind].id}=in=(${ids.join(",")})`,
      ),
    },
    hintTool,
  );
  return new Map((data.results ?? []).map((d) => [deviceId(kind, d), d]));
}

/** One device, read and checked against the key's sites. A device Jamf does not know and one
 *  outside the sites get the same refusal, so a kept key cannot learn which ids exist.
 *  `hintTool`: the read is the always-on read capability's permission. */
async function openDevice(
  ctx: ToolContext,
  mark: Mark,
  kind: Kind,
  id: string,
  sections: readonly string[] = INVENTORY[kind].sections,
): Promise<Device> {
  if (!ID.test(id)) throw new Denied(`Give a ${LABEL[kind]} id from jamf_search_devices`);
  const device = (await readDevices(ctx, mark, kind, [id], sections, "jamf_get_device")).get(id);
  if (!device || !inSites(keySites(ctx), siteOf(kind, device)))
    throw new Denied(`No ${LABEL[kind]} ${id} that this agent may use`);
  return device;
}

/** Several devices, each checked as openDevice does; refuses naming every one that is not. */
async function openDevices(ctx: ToolContext, mark: Mark, kind: Kind, ids: string[]) {
  const bad = ids.find((id) => !ID.test(id));
  if (bad !== undefined) throw new Denied(`${bad} is not a ${LABEL[kind]} id`);
  const sites = keySites(ctx);
  const found = await readDevices(ctx, mark, kind, ids, ["GENERAL"], "jamf_get_device");
  const refused = ids.filter((id) => {
    const d = found.get(id);
    return !d || !inSites(sites, siteOf(kind, d));
  });
  if (refused.length)
    throw new Denied(`No ${LABEL[kind]}s ${refused.join(", ")} that this agent may use`);
  return found;
}

async function openGroup(
  ctx: ToolContext,
  mark: Mark,
  kind: Kind,
  id: string,
): Promise<ClassicGroup> {
  if (!ID.test(id)) throw new Denied("Give a group id from jamf_list_groups");
  const g = CLASSIC_GROUP[kind];
  const refused = new Denied(`No ${LABEL[kind]} group ${id} that this agent may use`);
  const res = await ctx.call({ method: "GET", path: `${g.path}/${id}` });
  mark(res.status);
  if (res.status === 404) throw refused;
  if (res.status !== 200) {
    const detail = jamfError(res.data);
    throw new VendorFailed(
      res.status,
      `Jamf Pro answered ${res.status}${detail ? `: ${detail}` : ""}`,
    );
  }
  const group = (res.data as Record<string, ClassicGroup>)[g.root];
  if (!group || !inSites(keySites(ctx), group.site?.id)) throw refused;
  return group;
}

const groupMembers = (kind: Kind, group: ClassicGroup) =>
  (group[CLASSIC_GROUP[kind].members] ?? []).map((m) => String(m.id));

const offset = (cursor: string | undefined) => {
  if (cursor === undefined) return 0;
  if (!OFFSET.test(cursor)) throw new Denied("Give the `next` value from the previous call");
  return Number(cursor);
};

function managementId(kind: Kind, d: Device): string {
  const id = d.general?.managementId;
  if (!id)
    throw new Denied(
      `Jamf Pro has no management id for ${LABEL[kind]} ${deviceId(kind, d)}: it is not enrolled in MDM, so it cannot take commands.`,
    );
  return id;
}

async function mdmCommand(
  ctx: ToolContext,
  mark: Mark,
  kind: Kind,
  device: Device,
  commandData: Record<string, unknown>,
) {
  const queued = await ok<{ id?: string }[]>(ctx, mark, {
    method: "POST",
    path: "/api/v2/mdm/commands",
    body: { clientData: [{ managementId: managementId(kind, device) }], commandData },
  });
  const s = summary(kind, device);
  return json({
    device_id: s.id,
    name: s.name,
    command: commandData.commandType,
    command_id: queued?.[0]?.id ?? null,
    status: "queued",
    note: "Jamf Pro sent the command; the device carries it out when it next checks in.",
  });
}

const TRIGGERS = [
  "trigger_checkin",
  "trigger_enrollment_complete",
  "trigger_login",
  "trigger_logout",
  "trigger_network_state_changed",
  "trigger_startup",
];

async function readPolicy(ctx: ToolContext, mark: Mark, id: string, subset = "") {
  const refused = new Denied(`No policy ${id} that this agent may use`);
  const res = await ctx.call({ method: "GET", path: `/JSSResource/policies/id/${id}${subset}` });
  mark(res.status);
  if (res.status === 404) throw refused;
  if (res.status !== 200) {
    const detail = jamfError(res.data);
    throw new VendorFailed(
      res.status,
      `Jamf Pro answered ${res.status}${detail ? `: ${detail}` : ""}`,
    );
  }
  const policy = (res.data as { policy?: Policy }).policy;
  if (!policy) throw refused;
  return { policy, inScope: inSites(keySites(ctx), policy.general?.site?.id) };
}

export const jamfTools: Toolset = {
  names: [
    "jamf_search_devices",
    "jamf_get_device",
    "jamf_list_groups",
    "jamf_get_group",
    "jamf_change_group_members",
    "jamf_send_blank_push",
    "jamf_restart_device",
    "jamf_lock_device",
    "jamf_erase_device",
    "jamf_list_policies",
    "jamf_get_policy",
    "jamf_computer_policy_logs",
  ],

  register(server, ctx, allowed) {
    const readOnly = { readOnlyHint: true, openWorldHint: true } as const;
    const kind = z
      .enum(["computer", "mobile"])
      .describe("computer: Macs. mobile: iPhones, iPads, Apple TVs and other mobile devices.");
    const id = z.string().max(10).describe("A device id, from jamf_search_devices.");
    const cursor = z
      .string()
      .max(10)
      .optional()
      .describe("The `next` value from the previous call with the same arguments.");

    if (allowed.has("jamf_search_devices"))
      server.registerTool(
        "jamf_search_devices",
        {
          description:
            "Find computers or mobile devices in Jamf Pro: by part of the device name, or by exact serial number, username or email. Without `search`, lists them all by name. Each has its id, serial number, model, system version, user, last contact and inventory time, whether it is managed and supervised, and FileVault (computers) or encryption, passcode and Lost Mode (mobile devices). When `next` is set, call again with `cursor` set to it.",
          inputSchema: z.object({
            kind,
            search: z.string().max(100).optional(),
            max_results: z.number().int().min(1).max(100).default(25),
            cursor,
          }),
          annotations: readOnly,
        },
        ({ kind, search, max_results, cursor }) =>
          audited(ctx, "jamf_search_devices", `${kind} ${search ?? ""}`.trim(), async (mark) => {
            const s = search?.trim();
            if (s && !SEARCH.test(s))
              throw new Denied(
                "Search with letters, digits, spaces and . _ @ + ' - only (a name, serial number, username or email)",
              );
            const sites = keySites(ctx);
            const found = await scan<Device>(
              ctx,
              mark,
              (page) =>
                inventoryPath(
                  kind,
                  INVENTORY[kind].sections,
                  page,
                  max_results,
                  s ? INVENTORY[kind].search(s) : undefined,
                ),
              (d) => inSites(sites, siteOf(kind, d)),
              max_results,
              offset(cursor),
            );
            return json({
              devices: found.rows.map((d) => summary(kind, d)),
              next: found.next,
            });
          }),
      );

    if (allowed.has("jamf_get_device"))
      server.registerTool(
        "jamf_get_device",
        {
          description:
            "One computer or mobile device: everything jamf_search_devices shows, plus when it last enrolled and the groups it is in. With `include_apps`, also the apps installed on it (name, version, bundle id), at most 500.",
          inputSchema: z.object({
            kind,
            device_id: id,
            include_apps: z.boolean().default(false),
          }),
          annotations: readOnly,
        },
        ({ kind, device_id, include_apps }) =>
          audited(ctx, "jamf_get_device", `${kind} ${device_id}`, async (mark) => {
            const sections = [
              ...INVENTORY[kind].sections,
              INVENTORY[kind].groups,
              ...(include_apps ? ["APPLICATIONS"] : []),
            ];
            const d = await openDevice(ctx, mark, kind, device_id, sections);
            const groups =
              kind === "computer"
                ? ((d as Computer).groupMemberships ?? []).map((g) => ({
                    id: g.groupId ?? null,
                    name: g.groupName ?? null,
                    smart: g.smartGroup ?? null,
                  }))
                : ((d as Mobile).groups ?? []).map((g) => ({
                    id: g.groupId ?? null,
                    name: g.groupName ?? null,
                    smart: g.smart ?? null,
                  }));
            const apps = !include_apps
              ? undefined
              : kind === "computer"
                ? ((d as Computer).applications ?? []).map((a) => ({
                    name: a.name ?? null,
                    version: a.version ?? null,
                    bundle_id: a.bundleId ?? null,
                  }))
                : ((d as Mobile).applications ?? []).map((a) => ({
                    name: a.name ?? null,
                    version: a.shortVersion || a.version || null,
                    bundle_id: a.identifier ?? null,
                  }));
            return json({
              ...summary(kind, d),
              ...(kind === "computer" ? { site: (d as Computer).general?.site?.name ?? null } : {}),
              last_enrolled: d.general?.lastEnrolledDate ?? null,
              groups,
              ...(apps ? { apps: apps.slice(0, MAX_APPS), apps_total: apps.length } : {}),
            });
          }),
      );

    if (allowed.has("jamf_list_groups"))
      server.registerTool(
        "jamf_list_groups",
        {
          description:
            "List computer or mobile device groups: static groups (devices added by hand) or smart groups (devices matching criteria), with id, name, description, site and member count. Optionally only names containing `search`. When `next` is set, call again with `cursor` set to it.",
          inputSchema: z.object({
            kind,
            type: z.enum(["static", "smart"]),
            search: z.string().max(100).optional(),
            max_results: z.number().int().min(1).max(100).default(50),
            cursor,
          }),
          annotations: readOnly,
        },
        ({ kind, type, search, max_results, cursor }) =>
          audited(
            ctx,
            "jamf_list_groups",
            `${kind} ${type} ${search ?? ""}`.trim(),
            async (mark) => {
              const s = search?.trim();
              if (s && !SEARCH.test(s))
                throw new Denied("Search with letters, digits, spaces and . _ @ + ' - only");
              const list = GROUP_LISTS[kind];
              const sites = keySites(ctx);
              type Group = {
                id?: string;
                groupId?: string;
                name?: string;
                groupName?: string;
                description?: string;
                groupDescription?: string;
                siteId?: string;
                count?: number;
                membershipCount?: number;
              };
              const found = await scan<Group>(
                ctx,
                mark,
                (page) => {
                  const q = new URLSearchParams({
                    page: String(page),
                    "page-size": String(max_results),
                    sort: `${list.name}:asc`,
                  });
                  if (s) q.set("filter", `${list.name}=="*${s}*"`);
                  return `${list[type]}?${q}`;
                },
                (g) => inSites(sites, g.siteId),
                max_results,
                offset(cursor),
              );
              return json({
                groups: found.rows.map((g) => ({
                  id: g.id ?? g.groupId ?? null,
                  name: g.name ?? g.groupName ?? null,
                  description: (g.description ?? g.groupDescription) || null,
                  site_id: g.siteId ?? null,
                  members: g.count ?? g.membershipCount ?? null,
                })),
                next: found.next,
              });
            },
          ),
      );

    if (allowed.has("jamf_get_group"))
      server.registerTool(
        "jamf_get_group",
        {
          description:
            "One computer or mobile device group: whether it is smart or static, its site and its members (as jamf_search_devices shows them). When `next` is set, call again with `cursor` set to it.",
          inputSchema: z.object({
            kind,
            group_id: z.string().max(10).describe("A group id, from jamf_list_groups."),
            max_results: z.number().int().min(1).max(100).default(50),
            cursor,
          }),
          annotations: readOnly,
        },
        ({ kind, group_id, max_results, cursor }) =>
          audited(ctx, "jamf_get_group", `${kind} group ${group_id}`, async (mark) => {
            const group = await openGroup(ctx, mark, kind, group_id);
            const ids = groupMembers(kind, group);
            const start = offset(cursor);
            const window = ids.slice(start, start + max_results);
            const sites = keySites(ctx);
            const devices = await readDevices(
              ctx,
              mark,
              kind,
              window,
              INVENTORY[kind].sections,
              "jamf_get_device",
            );
            return json({
              id: String(group.id),
              name: group.name ?? null,
              smart: group.is_smart ?? null,
              site: group.site?.name ?? null,
              site_id: group.site?.id ?? null,
              member_count: ids.length,
              members: window
                .map((i) => devices.get(i))
                .filter((d): d is Device => Boolean(d) && inSites(sites, siteOf(kind, d!)))
                .map((d) => summary(kind, d)),
              next: start + max_results < ids.length ? String(start + max_results) : null,
            });
          }),
      );

    if (allowed.has("jamf_change_group_members"))
      server.registerTool(
        "jamf_change_group_members",
        {
          description:
            "Add devices to, or remove them from, a static computer or mobile device group (at most 50 in one call). Policies, configuration profiles and apps scoped to the group then reach, or leave, those devices. Smart groups cannot be changed: their members follow their criteria.",
          inputSchema: z.object({
            kind,
            group_id: z.string().max(10).describe("A static group id, from jamf_list_groups."),
            add: z.array(z.string().max(10)).max(50).default([]),
            remove: z.array(z.string().max(10)).max(50).default([]),
          }),
          annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
        },
        ({ kind, group_id, add, remove }) =>
          audited(
            ctx,
            "jamf_change_group_members",
            `${kind} group ${group_id}`,
            async (mark, retarget) => {
              const added = [...new Set(add)];
              const removed = [...new Set(remove)];
              if (!added.length && !removed.length)
                throw new Denied("Give at least one device id to add or remove");
              if (added.length + removed.length > 50)
                throw new Denied("Change at most 50 devices in one call");
              const both = added.filter((i) => removed.includes(i));
              if (both.length) throw new Denied(`Both added and removed: ${both.join(", ")}`);
              const group = await openGroup(ctx, mark, kind, group_id);
              if (group.is_smart)
                throw new Denied(
                  `Group ${group_id} is a smart group: its members follow its criteria. Only static groups can be changed.`,
                );
              await openDevices(ctx, mark, kind, [...added, ...removed]);
              const g = CLASSIC_GROUP[kind];
              const items = (ids: string[]) =>
                ids.map((i) => `<${g.item}><id>${i}</id></${g.item}>`).join("");
              const xml =
                `<${g.root}>` +
                (added.length ? `<${g.item}_additions>${items(added)}</${g.item}_additions>` : "") +
                (removed.length
                  ? `<${g.item}_deletions>${items(removed)}</${g.item}_deletions>`
                  : "") +
                `</${g.root}>`;
              await ok(ctx, mark, {
                method: "PUT",
                path: `${g.path}/${group_id}`,
                body: xml,
                headers: { "content-type": "application/xml" },
              });
              retarget(
                `${kind} group ${group_id}: added ${added.join(",") || "none"}; removed ${removed.join(",") || "none"}`,
              );
              return json({ group_id, name: group.name ?? null, added, removed });
            },
          ),
      );

    if (allowed.has("jamf_send_blank_push"))
      server.registerTool(
        "jamf_send_blank_push",
        {
          description:
            "Ask computers or mobile devices (at most 50) to check in with Jamf Pro now, so they pick up pending commands, profiles and apps. Changes nothing on the device.",
          inputSchema: z.object({
            kind,
            device_ids: z.array(z.string().max(10)).min(1).max(50),
          }),
          annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
        },
        ({ kind, device_ids }) =>
          audited(ctx, "jamf_send_blank_push", `${kind} ${device_ids.join(",")}`, async (mark) => {
            const ids = [...new Set(device_ids)];
            const devices = await openDevices(ctx, mark, kind, ids);
            const byManagementId = new Map(
              ids.map((i) => [managementId(kind, devices.get(i)!), i]),
            );
            const data = await ok<{ errorUuids?: string[] }>(ctx, mark, {
              method: "POST",
              path: "/api/v2/mdm/blank-push",
              body: { clientManagementIds: [...byManagementId.keys()] },
            });
            const failed = (data.errorUuids ?? []).map((u) => byManagementId.get(u) ?? u);
            return json({ sent: ids.filter((i) => !failed.includes(i)), failed });
          }),
      );

    if (allowed.has("jamf_restart_device"))
      server.registerTool(
        "jamf_restart_device",
        {
          description:
            "Restart a computer or mobile device. Whoever is using it loses unsaved work. With `notify_user` (computers), the user is told and can postpone it.",
          inputSchema: z.object({
            kind,
            device_id: id,
            notify_user: z.boolean().default(true).describe("Computers only."),
          }),
          annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
        },
        ({ kind, device_id, notify_user }) =>
          audited(ctx, "jamf_restart_device", `${kind} ${device_id}`, async (mark) => {
            const d = await openDevice(ctx, mark, kind, device_id);
            return mdmCommand(ctx, mark, kind, d, {
              commandType: "RESTART_DEVICE",
              ...(kind === "computer" ? { notifyUser: notify_user } : {}),
            });
          }),
      );

    if (allowed.has("jamf_lock_device"))
      server.registerTool(
        "jamf_lock_device",
        {
          description:
            "Lock a computer or mobile device, showing an optional message and phone number on its lock screen. A computer restarts and needs the 6-digit `pin` to unlock (required for computers; keep it for whoever unlocks it). A mobile device unlocks with its own passcode.",
          inputSchema: z.object({
            kind,
            device_id: id,
            pin: z.string().max(6).optional().describe("Computers only: 6 digits."),
            message: z.string().max(200).optional(),
            phone_number: z.string().max(40).optional(),
          }),
          annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
        },
        ({ kind, device_id, pin, message, phone_number }) =>
          audited(ctx, "jamf_lock_device", `${kind} ${device_id}`, async (mark) => {
            if (kind === "computer" && !(pin && PIN.test(pin)))
              throw new Denied("Locking a computer needs a 6-digit pin");
            if (kind === "mobile" && pin)
              throw new Denied("A mobile device unlocks with its own passcode: leave pin out");
            const d = await openDevice(ctx, mark, kind, device_id);
            return mdmCommand(ctx, mark, kind, d, {
              commandType: "DEVICE_LOCK",
              ...(pin ? { pin } : {}),
              ...(message ? { message } : {}),
              ...(phone_number ? { phoneNumber: phone_number } : {}),
            });
          }),
      );

    if (allowed.has("jamf_erase_device"))
      server.registerTool(
        "jamf_erase_device",
        {
          description:
            "Erase a computer or mobile device: everything on it is deleted for good and it returns to factory settings. Give its exact serial number as `confirm_serial_number` (from jamf_get_device); a mismatch is refused. Some older Macs need a 6-digit `pin` to start up again.",
          inputSchema: z.object({
            kind,
            device_id: id,
            confirm_serial_number: z.string().max(40),
            pin: z.string().max(6).optional().describe("Computers only: 6 digits."),
          }),
          annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
        },
        ({ kind, device_id, confirm_serial_number, pin }) =>
          audited(ctx, "jamf_erase_device", `${kind} ${device_id}`, async (mark, retarget) => {
            if (pin !== undefined && (kind === "mobile" || !PIN.test(pin)))
              throw new Denied("A pin is for computers only, and is 6 digits");
            const d = await openDevice(ctx, mark, kind, device_id);
            const serial = d.hardware?.serialNumber ?? "";
            if (!serial || serial.toLowerCase() !== confirm_serial_number.trim().toLowerCase())
              throw new Denied(
                `The serial number does not match ${LABEL[kind]} ${device_id}. Read it with jamf_get_device and check it is the device to erase.`,
              );
            retarget(`${kind} ${device_id} (serial ${serial})`);
            return mdmCommand(ctx, mark, kind, d, {
              commandType: "ERASE_DEVICE",
              ...(pin ? { pin } : {}),
            });
          }),
      );

    if (allowed.has("jamf_list_policies"))
      server.registerTool(
        "jamf_list_policies",
        {
          description:
            "List Jamf Pro policies (what runs on computers: installs, scripts, settings) by id and name. Optionally only names containing `search`. When `next` is set, call again with `cursor` set to it.",
          inputSchema: z.object({
            search: z.string().max(100).optional(),
            max_results: z.number().int().min(1).max(100).default(50),
            cursor,
          }),
          annotations: readOnly,
        },
        ({ search, max_results, cursor }) =>
          audited(ctx, "jamf_list_policies", search ?? null, async (mark) => {
            const data = await ok<{ policies?: { id: number; name?: string }[] }>(ctx, mark, {
              method: "GET",
              path: "/JSSResource/policies",
            });
            const s = search?.trim().toLowerCase();
            const all = (data.policies ?? []).filter(
              (p) => !s || (p.name ?? "").toLowerCase().includes(s),
            );
            const sites = keySites(ctx);
            let at = offset(cursor);
            const policies: { id: string; name: string | null }[] = [];
            let reads = 0;
            for (; at < all.length && policies.length < max_results; at++) {
              const p = all[at];
              if (sites) {
                if (reads++ >= MAX_POLICY_READS) break;
                const { inScope } = await readPolicy(ctx, mark, String(p.id), "/subset/General");
                if (!inScope) continue;
              }
              policies.push({ id: String(p.id), name: p.name ?? null });
            }
            return json({ policies, next: at < all.length ? String(at) : null });
          }),
      );

    if (allowed.has("jamf_get_policy"))
      server.registerTool(
        "jamf_get_policy",
        {
          description:
            "One Jamf Pro policy: whether it is on, how often and on what triggers it runs, its category and site, what it is scoped to (all computers, groups, how many computers by hand, excluded groups), its packages and scripts, and whether it is offered in Self Service.",
          inputSchema: z.object({
            policy_id: z.string().max(10).describe("A policy id, from jamf_list_policies."),
          }),
          annotations: readOnly,
        },
        ({ policy_id }) =>
          audited(ctx, "jamf_get_policy", policy_id, async (mark) => {
            if (!ID.test(policy_id)) throw new Denied("Give a policy id from jamf_list_policies");
            const { policy, inScope } = await readPolicy(ctx, mark, policy_id);
            if (!inScope) throw new Denied(`No policy ${policy_id} that this agent may use`);
            const g = policy.general ?? {};
            return json({
              id: policy_id,
              name: g.name ?? null,
              enabled: g.enabled ?? null,
              frequency: g.frequency ?? null,
              triggers: [
                ...TRIGGERS.filter((t) => g[t] === true).map((t) => t.replace("trigger_", "")),
                ...(g.trigger_other ? [g.trigger_other] : []),
              ],
              category: g.category?.name ?? null,
              site: g.site?.name ?? null,
              scope: {
                all_computers: policy.scope?.all_computers ?? null,
                computer_groups: (policy.scope?.computer_groups ?? []).map((x) => x.name),
                computers: (policy.scope?.computers ?? []).length,
                excluded_computer_groups: (policy.scope?.exclusions?.computer_groups ?? []).map(
                  (x) => x.name,
                ),
              },
              packages: (policy.package_configuration?.packages ?? []).map((x) => x.name),
              scripts: (policy.scripts ?? []).map((x) => x.name),
              self_service: policy.self_service?.use_for_self_service ?? null,
            });
          }),
      );

    if (allowed.has("jamf_computer_policy_logs"))
      server.registerTool(
        "jamf_computer_policy_logs",
        {
          description:
            "The policy runs Jamf Pro logged on one computer, newest first: policy id and name, status (such as Completed or Failed), when, and the user signed in. When `next` is set, call again with `cursor` set to it.",
          inputSchema: z.object({
            computer_id: z.string().max(10).describe("A computer id, from jamf_search_devices."),
            max_results: z.number().int().min(1).max(100).default(25),
            cursor,
          }),
          annotations: readOnly,
        },
        ({ computer_id, max_results, cursor }) =>
          audited(ctx, "jamf_computer_policy_logs", computer_id, async (mark) => {
            await openDevice(ctx, mark, "computer", computer_id, ["GENERAL"]);
            type Log = {
              policy_id?: number;
              policy_name?: string;
              username?: string;
              date_time_epoch?: string | number;
              date_time_utc?: string;
              status?: string;
            };
            const data = await ok<{ computer_history?: { policy_logs?: Log[] } }>(ctx, mark, {
              method: "GET",
              path: `/JSSResource/computerhistory/id/${computer_id}/subset/PolicyLogs`,
            });
            const logs = [...(data.computer_history?.policy_logs ?? [])].sort(
              (a, b) => Number(b.date_time_epoch ?? 0) - Number(a.date_time_epoch ?? 0),
            );
            const start = offset(cursor);
            return json({
              logs: logs.slice(start, start + max_results).map((l) => ({
                policy_id: l.policy_id ?? null,
                policy_name: l.policy_name ?? null,
                status: l.status ?? null,
                completed: l.date_time_utc ?? null,
                username: l.username || null,
              })),
              next: start + max_results < logs.length ? String(start + max_results) : null,
            });
          }),
      );
  },
};
