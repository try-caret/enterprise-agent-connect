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
import { MAX_SCOPE_IDS } from "../entra/tools.js";
import { cursorArg, nextCursor, ok, pageQuery } from "../sharepoint-files/tools.js";

// Devices Microsoft Intune manages, read-only. A key's `groups` limit keeps it to devices whose
// primary user (managedDevice.userId) is in those groups, checked on every call, lookups by id
// included. Request shapes: learn.microsoft.com/graph/api/intune-devices-manageddevice-list,
// intune-devices-manageddevice-get, resources/intune-devices-user (a user's managedDevices),
// resources/intune-devices-manageddevice (v1.0 and beta: the
// detectedApps relationship is beta only), intune-devices-detectedapp-list and
// directoryobject-checkmemberobjects. intune_autopilot_status: v1.0 windowsAutopilotDeviceIdentity
// (list; resources/intune-enrollment-windowsautopilotdeviceidentity) and, for the profile name,
// the beta `deploymentProfile` relationship (intune-shared-windowsautopilotdeploymentprofile-get),
// both DeviceManagementServiceConfig.Read.All (read October 25, 2026).

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const UPN = /^[^\s/\\?#%'"]+@[^\s/\\?#%'"]+$/;
const DEVICES = "/v1.0/deviceManagement/managedDevices";
const AUTOPILOT = "/v1.0/deviceManagement/windowsAutopilotDeviceIdentities";
// Letters, digits and dashes: nothing that could change a $filter or a path.
const SERIAL = /^[A-Za-z0-9-]{1,50}$/;
const NO_ID = "00000000-0000-0000-0000-000000000000";
const AUTOPILOT_PAGES = 5;
// Under a groups limit, an Autopilot device no person is tied to (no primary user, no assigned
// user: a spare). "minimum" answers only its serial, Autopilot state and whether Intune manages
// it and its compliance; "refuse" treats it as out of scope, as the other Intune tools treat a
// device with no user. Open question for the owner: change this one line to "refuse".
const UNASSIGNED_UNDER_GROUPS = "minimum" as "minimum" | "refuse";
export const COMPLIANCE_STATES = [
  "unknown",
  "compliant",
  "noncompliant",
  "conflict",
  "error",
  "inGracePeriod",
  "configManager",
] as const;
// ponytail: the summary reads at most this many devices; a server-side report when
// fleets outgrow it.
export const SUMMARY_MAX_DEVICES = 5000;

type Mark = (status: number) => void;
type Device = Record<string, unknown> & { id: string; userId?: string | null };

/** The key's groups, refused when Intune is not set up to check them. */
function keyGroups(ctx: ToolContext): string[] | null {
  const groups = listLimit(ctx.limits, "groups");
  if (!groups) return null;
  if (ctx.config.people_groups !== "groups")
    throw new Denied(
      'This key is kept to groups, but Microsoft Intune is not set up to check groups. An administrator has to choose "Keep keys to people\'s groups" on the Microsoft Intune page.',
    );
  if (groups.length > MAX_SCOPE_IDS)
    throw new Denied(
      `This key names more than ${MAX_SCOPE_IDS} groups. Ask an administrator to narrow it.`,
    );
  return groups;
}

/** Whether a user (id or sign-in name) is in any of the groups; one Graph does not find is in none. */
async function inGroups(ctx: ToolContext, mark: Mark, groups: string[], user: string) {
  const res = await ctx.call({
    method: "POST",
    path: `/v1.0/users/${encodeURIComponent(user)}/checkMemberObjects`,
    body: { ids: groups },
  });
  mark(res.status);
  if (res.status === 404) return false;
  if (res.status !== 200) {
    const message = (res.data as { error?: { message?: string } } | null)?.error?.message;
    throw new VendorFailed(
      res.status,
      `Microsoft Graph answered ${res.status}${message ? `: ${message}` : ""}`,
    );
  }
  return ((res.data as { value?: string[] }).value ?? []).length > 0;
}

/** Whether each device's primary user is in the key's groups, one check per user. */
export function deviceFilter(ctx: ToolContext, mark: Mark) {
  const groups = keyGroups(ctx);
  if (!groups) return async () => true;
  const seen = new Map<string, Promise<boolean>>();
  return (device: Device) => {
    const user = device.userId;
    if (!user || !GUID.test(user)) return Promise.resolve(false);
    if (!seen.has(user)) seen.set(user, inGroups(ctx, mark, groups, user));
    return seen.get(user)!;
  };
}

/** Everyone in the key's groups, read once: scanning a whole fleet then costs a few listings
 *  instead of a membership check per user. Null when the key has no groups. */
async function groupMembers(ctx: ToolContext, mark: Mark): Promise<Set<string> | null> {
  const groups = keyGroups(ctx);
  if (!groups) return null;
  const ids = new Set<string>();
  for (const group of groups) {
    let cursor: string | undefined;
    do {
      const data = await ok(ctx, mark, {
        method: "GET",
        path: `/v1.0/groups/${encodeURIComponent(group)}/transitiveMembers/microsoft.graph.user?$select=id&${pageQuery(999, cursor)}`,
      });
      for (const u of data.value as { id: string }[]) ids.add(u.id.toLowerCase());
      cursor = nextCursor(data) ?? undefined;
    } while (cursor);
  }
  return ids;
}

const summary = (d: Record<string, unknown>) => ({
  id: d.id,
  name: d.deviceName ?? null,
  user: d.userPrincipalName ?? null,
  user_name: d.userDisplayName ?? null,
  os: d.operatingSystem ?? null,
  os_version: d.osVersion ?? null,
  compliance: d.complianceState ?? null,
  last_sync: d.lastSyncDateTime ?? null,
  encrypted: d.isEncrypted ?? null,
  model: d.model ?? null,
  manufacturer: d.manufacturer ?? null,
  serial_number: d.serialNumber ?? null,
  management_state: d.managementState ?? null,
  owner: d.managedDeviceOwnerType ?? null,
});

const quoted = (v: string) => `'${v.replaceAll("'", "''")}'`;

async function allowedDevices(allowed: ReturnType<typeof deviceFilter>, devices: Device[]) {
  const out = [];
  for (const d of devices) if (await allowed(d)) out.push(d);
  return out;
}

/** One device, read and checked against the key's limits. */
async function openDevice(ctx: ToolContext, mark: Mark, id: string): Promise<Device> {
  if (!GUID.test(id)) throw new Denied("Give a device id from intune_list_devices");
  const allowed = deviceFilter(ctx, mark);
  const device = await ok(ctx, mark, { method: "GET", path: `${DEVICES}/${id}` });
  if (!(await allowed(device))) throw new Denied("This agent may not read that device");
  return device;
}

const listed = (data: { value: Device[]; "@odata.nextLink"?: string }, devices: Device[]) => ({
  devices: devices.map(summary),
  more: Boolean(data["@odata.nextLink"]),
  next: nextCursor(data),
});

/** A Graph read where 404 means "nothing there": null. */
async function okOrNone(ctx: ToolContext, mark: Mark, req: Parameters<typeof ok>[2]) {
  let notFound = false;
  const data = await ok(
    ctx,
    (status) => {
      mark(status);
      notFound = status === 404;
    },
    req,
  ).catch((error) => {
    if (notFound) return null;
    throw error;
  });
  return data as Device | null;
}

/** Where Graph says a listing continues, as a path for the next call: only another page of
 *  the same listing on Graph (`@odata.nextLink` is used as given, whatever its paging
 *  parameter), never another host or resource. */
function nextPath(link: string, path: string): string {
  const url = URL.parse(link);
  if (
    url?.origin !== "https://graph.microsoft.com" ||
    url.pathname.toLowerCase() !== path.toLowerCase()
  )
    throw new Error(`Graph gave a next page outside ${path}: ${link.slice(0, 200)}`);
  return `${url.pathname}${url.search}`;
}

/** The Autopilot identity with exactly this serial number (any case), or null. The filter is
 *  `contains` and the match is made here: Microsoft's pages for the list do not document
 *  $filter, and `contains` on serialNumber is what Intune tooling is seen to use. */
async function autopilotIdentity(ctx: ToolContext, mark: Mark, serial: string) {
  const filter = encodeURIComponent(`contains(serialNumber,${quoted(serial)})`);
  let path = `${AUTOPILOT}?$top=50&$filter=${filter}`;
  for (let page = 0; page < AUTOPILOT_PAGES; page++) {
    const data = await ok(ctx, mark, { method: "GET", path });
    const hit = (data.value as Record<string, unknown>[]).find(
      (i) => String(i.serialNumber ?? "").toLowerCase() === serial.toLowerCase(),
    );
    if (hit) return hit;
    const link = data["@odata.nextLink"];
    if (!link) return null;
    path = nextPath(link, AUTOPILOT);
  }
  throw new Denied(
    "Too many Autopilot devices match that serial number. Give the full serial number.",
  );
}

export const intuneTools: Toolset = {
  names: [
    "intune_list_devices",
    "intune_get_device",
    "intune_device_apps",
    "intune_person_devices",
    "intune_compliance_summary",
    "intune_autopilot_status",
  ],

  register(server, ctx, allowed) {
    const readOnly = { readOnlyHint: true, openWorldHint: true } as const;
    const deviceId = z.string().max(36).describe("A device id, from intune_list_devices.");

    if (allowed.has("intune_list_devices"))
      server.registerTool(
        "intune_list_devices",
        {
          description:
            "List managed devices: name, user, system and version, compliance, last check-in, encryption, model and serial number. Optionally only names containing `search`, or one compliance state. When `more` is true, call again with `cursor` set to `next`.",
          inputSchema: z.object({
            search: z.string().max(100).optional().describe("Part of the device name."),
            compliance: z.enum(COMPLIANCE_STATES).optional(),
            max_results: z.number().int().min(1).max(50).default(25),
            cursor: cursorArg,
          }),
          annotations: readOnly,
        },
        ({ search, compliance, max_results, cursor }) =>
          audited(
            ctx,
            "intune_list_devices",
            [search, compliance].filter(Boolean).join(" ") || null,
            async (mark) => {
              const allowed = deviceFilter(ctx, mark);
              const filters = [
                ...(search?.trim() ? [`contains(deviceName,${quoted(search.trim())})`] : []),
                ...(compliance ? [`complianceState eq ${quoted(compliance)}`] : []),
              ];
              const filter = filters.length
                ? `&$filter=${encodeURIComponent(filters.join(" and "))}`
                : "";
              const data = await ok(ctx, mark, {
                method: "GET",
                path: `${DEVICES}?${pageQuery(max_results, cursor)}${filter}`,
              });
              return json(listed(data, await allowedDevices(allowed, data.value)));
            },
          ),
      );

    if (allowed.has("intune_get_device"))
      server.registerTool(
        "intune_get_device",
        {
          description:
            "One device's details: everything intune_list_devices shows, plus enrollment date, compliance grace period, jailbreak state and how it was enrolled.",
          inputSchema: z.object({ device_id: deviceId }),
          annotations: readOnly,
        },
        ({ device_id }) =>
          audited(ctx, "intune_get_device", device_id, async (mark) => {
            const d = await openDevice(ctx, mark, device_id);
            return json({
              ...summary(d),
              enrolled: d.enrolledDateTime ?? null,
              compliance_grace_period_ends: d.complianceGracePeriodExpirationDateTime ?? null,
              jailbroken: d.jailBroken ?? null,
              enrollment_type: d.deviceEnrollmentType ?? null,
              entra_device_id: d.azureADDeviceId ?? null,
            });
          }),
      );

    if (allowed.has("intune_device_apps"))
      server.registerTool(
        "intune_device_apps",
        {
          description:
            "The apps Intune found on a device, with version and publisher. When `more` is true, call again with `cursor` set to `next`.",
          inputSchema: z.object({
            device_id: deviceId,
            max_results: z.number().int().min(1).max(200).default(100),
            cursor: cursorArg,
          }),
          annotations: readOnly,
        },
        ({ device_id, max_results, cursor }) =>
          audited(ctx, "intune_device_apps", device_id, async (mark) => {
            await openDevice(ctx, mark, device_id);
            // Only Graph beta has a device's detectedApps.
            const data = await ok(ctx, mark, {
              method: "GET",
              path: `/beta/deviceManagement/managedDevices/${device_id}/detectedApps?${pageQuery(max_results, cursor)}`,
            });
            return json({
              apps: (data.value as Record<string, unknown>[]).map((a) => ({
                name: a.displayName ?? null,
                version: a.version ?? null,
                publisher: a.publisher ?? null,
                platform: a.platform ?? null,
              })),
              more: Boolean(data["@odata.nextLink"]),
              next: nextCursor(data),
            });
          }),
      );

    if (allowed.has("intune_person_devices"))
      server.registerTool(
        "intune_person_devices",
        {
          description:
            "The managed devices of one person, by their sign-in name or user id. When `more` is true, call again with `cursor` set to `next`.",
          inputSchema: z.object({
            person: z
              .string()
              .max(320)
              .describe("A sign-in name such as ada@contoso.com, or a user id."),
            max_results: z.number().int().min(1).max(50).default(25),
            cursor: cursorArg,
          }),
          annotations: readOnly,
        },
        ({ person, max_results, cursor }) =>
          audited(ctx, "intune_person_devices", person, async (mark) => {
            const p = person.trim();
            if (!GUID.test(p) && !UPN.test(p))
              throw new Denied("Give a sign-in name such as ada@contoso.com, or a user id");
            // Same refusal for a person outside the groups and one who does not exist, as
            // Entra gives, so a kept key cannot learn which names exist.
            const groups = keyGroups(ctx);
            if (groups && !(await inGroups(ctx, mark, groups, p)))
              throw new Denied("This agent may not look up that person");
            // The person's own devices, so the check above is the scope check: this listing
            // leaves userId and userPrincipalName blank (seen live), which deviceFilter would
            // read as a device with no user.
            const data = await ok(ctx, mark, {
              method: "GET",
              path: `/v1.0/users/${encodeURIComponent(p)}/managedDevices?${pageQuery(max_results, cursor)}`,
            });
            return json(listed(data, data.value));
          }),
      );

    if (allowed.has("intune_compliance_summary"))
      server.registerTool(
        "intune_compliance_summary",
        {
          description: `How many of the devices this agent may see are in each compliance state, by operating system. Reads at most ${SUMMARY_MAX_DEVICES} devices; \`complete\` says whether that was all of them. List the devices of one state with intune_list_devices.`,
          inputSchema: z.object({}),
          annotations: readOnly,
        },
        () =>
          audited(ctx, "intune_compliance_summary", null, async (mark) => {
            const members = await groupMembers(ctx, mark);
            const counts: Record<string, Record<string, number>> = {};
            let counted = 0;
            let scanned = 0;
            let cursor: string | undefined;
            let complete = true;
            do {
              const data = await ok(ctx, mark, {
                method: "GET",
                path: `${DEVICES}?$select=id,userId,operatingSystem,complianceState&${pageQuery(1000, cursor)}`,
              });
              scanned += data.value.length;
              for (const d of data.value as Device[]) {
                if (members && !(d.userId && members.has(d.userId.toLowerCase()))) continue;
                const os = String(d.operatingSystem ?? "unknown");
                const state = String(d.complianceState ?? "unknown");
                counts[os] ??= {};
                counts[os][state] = (counts[os][state] ?? 0) + 1;
                counted++;
              }
              cursor = nextCursor(data) ?? undefined;
              if (cursor && scanned >= SUMMARY_MAX_DEVICES) {
                complete = false;
                break;
              }
            } while (cursor);
            return json({ devices_counted: counted, complete, by_operating_system: counts });
          }),
      );

    if (allowed.has("intune_autopilot_status"))
      server.registerTool(
        "intune_autopilot_status",
        {
          description:
            "Where a Windows Autopilot device stands, by its exact serial number: its enrollment state (unknown, enrolled, pendingReset, failed, notContacted), assigned profile, group tag, last contact with Intune and, once Intune manages it, that device's id, compliance and primary user. Unknown to Autopilot: {serial_number, found: false}. When this key is kept to groups and no person is tied to the device, only {serial_number, found, deployment_state, managed, compliance, limited} is answered.",
          inputSchema: z.object({
            serial_number: z
              .string()
              .max(50)
              .describe(
                "The device's exact serial number: letters, digits and dashes only (one with spaces, dots or underscores is refused).",
              ),
          }),
          annotations: readOnly,
        },
        ({ serial_number }) =>
          audited(ctx, "intune_autopilot_status", serial_number, async (mark) => {
            const serial = serial_number.trim();
            if (!SERIAL.test(serial))
              throw new Denied(
                "Give the device's exact serial number: letters, digits and dashes only, up to 50. Serial numbers with spaces, dots or underscores (some virtual machines have them) are not accepted by this tool.",
              );
            const groups = keyGroups(ctx);
            const identity = await autopilotIdentity(ctx, mark, serial);
            if (!identity) return json({ serial_number: serial, found: false });

            const managedId = String(identity.managedDeviceId ?? "");
            // The identity can outlive the managed device it pointed to (404). This read is the
            // always-on `read` capability's permission, so a 403 carries that capability's hint.
            const managed: Device | null =
              GUID.test(managedId) && managedId !== NO_ID
                ? await okOrNone(ctx, mark, {
                    method: "GET",
                    path: `${DEVICES}/${managedId}`,
                  }).catch((error) => {
                    if (error instanceof VendorFailed)
                      throw new VendorFailed(error.status, error.message, "intune_get_device");
                    throw error;
                  })
                : null;
            // Under a groups limit every person tied to the device must be in the groups: the
            // managed device's primary user and a user assigned on the Autopilot record.
            if (groups) {
              const inScope = deviceFilter(ctx, mark);
              if (managed?.userId && !(await inScope(managed)))
                throw new Denied("This agent may not read that device");
              const assigned = String(identity.userPrincipalName ?? "").trim();
              if (
                assigned &&
                (!UPN.test(assigned) || !(await inGroups(ctx, mark, groups, assigned)))
              )
                throw new Denied("This agent may not read that device");
              // Nobody tied to it (a spare): see UNASSIGNED_UNDER_GROUPS.
              if (!managed?.userId && !assigned) {
                if (UNASSIGNED_UNDER_GROUPS === "refuse")
                  throw new Denied("This agent may not read that device");
                return json({
                  serial_number: String(identity.serialNumber ?? serial),
                  found: true,
                  deployment_state: identity.enrollmentState ?? null,
                  managed: Boolean(managed),
                  compliance: managed ? (managed.complianceState ?? null) : null,
                  limited:
                    "No person is tied to this device and this key is kept to groups: only its Autopilot state and compliance are answered.",
                });
              }
            }

            // Only Graph beta has the assigned profile; a device with none is expected to answer
            // 404. Any other failure leaves the profile out, with why, not the v1.0 answer.
            let profile: Device | null = null;
            let profileError: string | null = null;
            if (GUID.test(String(identity.id)))
              profile = await okOrNone(ctx, mark, {
                method: "GET",
                path: `/beta/deviceManagement/windowsAutopilotDeviceIdentities/${identity.id}/deploymentProfile`,
              }).catch((error) => {
                if (!(error instanceof VendorFailed)) throw error;
                profileError = `Microsoft Graph (beta) answered ${error.status} for the profile`;
                return null;
              });
            return json({
              serial_number: String(identity.serialNumber ?? serial),
              found: true,
              deployment_state: identity.enrollmentState ?? null,
              enrollment_profile: (profile?.displayName as string | undefined) ?? null,
              ...(profileError ? { enrollment_profile_error: profileError } : {}),
              group_tag: identity.groupTag || null,
              last_contact: identity.lastContactedDateTime ?? null,
              managed_device: managed
                ? {
                    id: managed.id,
                    compliance: managed.complianceState ?? null,
                    primary_user: managed.userId ? (managed.userPrincipalName ?? null) : null,
                  }
                : null,
            });
          }),
      );
  },
};
