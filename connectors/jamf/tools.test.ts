import assert from "node:assert/strict";
import { test } from "node:test";
import type { z } from "zod";
import { jamfTools, MAX_POLICY_READS, MAX_SCOPED_PAGES } from "./tools.ts";
import type { ToolContext } from "../../packages/connector-kit/src/common.ts";
import type { VendorRequest, VendorResponse } from "../../packages/connector-kit/src/types.ts";

// The Jamf Pro tools against an in-memory Jamf Pro: the tools' own rules (the sites limit,
// refusals before any call, what is sent and answered), not the MCP plumbing. Sites 1 and 2;
// computer 3 has no site (-1), computer 4 never enrolled in MDM (no management id).

const computer = (id: number, site: number, extra: Record<string, unknown> = {}) => ({
  id: String(id),
  general: {
    name: `Mac ${id}`,
    lastContact: `2026-10-${10 + id}T08:00:00Z`,
    reportDate: "2026-10-20T08:00:00Z",
    lastEnrolledDate: "2026-01-01T00:00:00Z",
    supervised: true,
    remoteManagement: { managed: true, managementUsername: "jamfadmin" },
    managementId: id === 4 ? undefined : `mc-${id}`,
    site: { id: String(site), name: site === -1 ? "None" : `Site ${site}` },
    extensionAttributes: [{ name: "secret-ea", values: ["x"] }],
  },
  hardware: { serialNumber: `C0${id}`, model: "MacBook Pro", macAddress: "aa:bb" },
  operatingSystem: { name: "macOS", version: "15.1", fileVault2Status: "ALL_ENCRYPTED" },
  userAndLocation: { username: `user${id}`, realname: `User ${id}`, email: `user${id}@x.com` },
  groupMemberships: [{ groupId: "5", groupName: "Pilot", smartGroup: false }],
  applications: [{ name: "Safari.app", version: "18.0", bundleId: "com.apple.Safari", path: "/A" }],
  ...extra,
});
const mobile = (id: number, site: number) => ({
  mobileDeviceId: String(id),
  deviceType: "iOS",
  general: {
    displayName: `iPad ${id}`,
    siteId: String(site),
    osVersion: "18.1",
    lastContactDate: "2026-10-21T08:00:00Z",
    lastInventoryUpdateDate: "2026-10-21T07:00:00Z",
    lastEnrolledDate: "2026-02-01T00:00:00Z",
    managed: true,
    supervised: true,
    managementId: `mm-${id}`,
  },
  hardware: { serialNumber: `M${id}`, model: "iPad Air" },
  userAndLocation: { username: `kid${id}`, realName: `Kid ${id}`, emailAddress: `kid${id}@x.com` },
  security: { dataProtected: true, passcodePresent: true, lostModeEnabled: false },
  groups: [{ groupId: "8", groupName: "Carts", smart: true }],
  applications: [{ name: "Notes", version: "1", shortVersion: "1.2", identifier: "com.x.notes" }],
});

const COMPUTERS = [computer(1, 1), computer(2, 2), computer(3, -1), computer(4, 1)];
const MOBILES = [mobile(10, 1), mobile(11, 2)];
const GROUPS: Record<string, Record<string, unknown>> = {
  "/JSSResource/computergroups/id/5": {
    computer_group: {
      id: 5,
      name: "Pilot",
      is_smart: false,
      site: { id: 1, name: "Site 1" },
      computers: [{ id: 1 }, { id: 4 }, { id: 2 }],
    },
  },
  "/JSSResource/computergroups/id/6": {
    computer_group: { id: 6, name: "Other", is_smart: false, site: { id: 2 }, computers: [] },
  },
  "/JSSResource/computergroups/id/7": {
    computer_group: { id: 7, name: "Smart", is_smart: true, site: { id: 1 }, computers: [] },
  },
  "/JSSResource/mobiledevicegroups/id/8": {
    mobile_device_group: {
      id: 8,
      name: "Carts",
      is_smart: false,
      site: { id: 1, name: "Site 1" },
      mobile_devices: [{ id: 10 }],
    },
  },
};
const policy = (id: number, site: number) => ({
  policy: {
    general: {
      id,
      name: `Policy ${id}`,
      enabled: true,
      frequency: "Once per computer",
      trigger_checkin: true,
      trigger_login: false,
      trigger_other: id === 20 ? "install-office" : "",
      category: { id: 1, name: "Apps" },
      site: { id: site, name: `Site ${site}` },
    },
    scope: {
      all_computers: false,
      computers: [{ id: 1 }, { id: 2 }],
      computer_groups: [{ id: 5, name: "Pilot" }],
      exclusions: { computer_groups: [{ id: 6, name: "Other" }] },
    },
    self_service: { use_for_self_service: true },
    package_configuration: { packages: [{ id: 1, name: "Office.pkg", action: "Install" }] },
    scripts: [{ id: 2, name: "cleanup.sh", priority: "After" }],
  },
});
const POLICIES = [20, 21, 22].map((id, i) => ({ id, name: `Policy ${id}`, site: [1, 2, -1][i] }));

type Options = {
  // Answers instead of the fake's own, by path prefix.
  answers?: Record<string, VendorResponse>;
  policies?: { id: number; name: string; site: number }[];
  errorUuids?: string[];
  forbidden?: Record<string, string>;
};

// The RSQL Jamf would run, in the parts these tools send.
function matches(filter: string | null, row: Record<string, unknown>, mobileKind: boolean) {
  if (!filter) return true;
  const inIds = /=in=\(([^)]*)\)/.exec(filter);
  const id = String(mobileKind ? row.mobileDeviceId : row.id);
  if (inIds) return inIds[1].split(",").includes(id);
  const text = /=="\*(.*?)\*"/.exec(filter)?.[1] ?? "";
  const name = String(
    mobileKind
      ? (row.general as { displayName: string }).displayName
      : (row.general as { name: string }).name,
  );
  const serial = (row.hardware as { serialNumber: string }).serialNumber;
  return name.includes(text) || serial === text;
}

function setup(limits: Record<string, unknown> = {}, options: Options = {}) {
  const requests: VendorRequest[] = [];
  const audit: { tool: string; target: string | null; decision: string; reason: string | null }[] =
    [];
  const policies = options.policies ?? POLICIES;
  const call = async (req: VendorRequest): Promise<VendorResponse> => {
    requests.push(req);
    const [path, query = ""] = req.path.split("?");
    const q = new URLSearchParams(query);
    for (const [prefix, answer] of Object.entries(options.answers ?? {}))
      if (req.path.startsWith(prefix)) return answer;
    const listing = (rows: Record<string, unknown>[], mobileKind: boolean) => {
      const all = rows.filter((r) => matches(q.get("filter"), r, mobileKind));
      const page = Number(q.get("page"));
      const size = Number(q.get("page-size"));
      return {
        status: 200,
        data: { totalCount: all.length, results: all.slice(page * size, page * size + size) },
      };
    };
    if (path === "/api/v4/computers-inventory") return listing(COMPUTERS, false);
    if (path === "/api/v2/mobile-devices/detail") return listing(MOBILES, true);
    if (path === "/api/v3/computer-groups/static-groups")
      return {
        status: 200,
        data: {
          totalCount: 2,
          results: [
            { id: "5", name: "Pilot", siteId: "1", count: 3 },
            { id: "6", name: "Other", siteId: "2", count: 0 },
          ],
        },
      };
    if (path === "/api/v2/mobile-device-groups/smart-groups")
      return {
        status: 200,
        data: {
          totalCount: 1,
          results: [{ groupId: "9", groupName: "Smart iPads", siteId: "1", count: 4 }],
        },
      };
    if (GROUPS[path] && req.method === "GET") return { status: 200, data: GROUPS[path] };
    if (GROUPS[path] && req.method === "PUT")
      return { status: 201, data: "<computer_group><id>5</id></computer_group>" };
    if (path === "/JSSResource/policies")
      return { status: 200, data: { policies: policies.map(({ id, name }) => ({ id, name })) } };
    const p = /^\/JSSResource\/policies\/id\/(\d+)/.exec(path);
    if (p) {
      const found = policies.find((x) => x.id === Number(p[1]));
      return found
        ? { status: 200, data: policy(found.id, found.site) }
        : { status: 404, data: "<html><body>Not Found</body></html>" };
    }
    if (path === "/JSSResource/computerhistory/id/1/subset/PolicyLogs")
      return {
        status: 200,
        data: {
          computer_history: {
            policy_logs: [
              {
                policy_id: 20,
                policy_name: "Old",
                username: "ada",
                date_time_epoch: 1000,
                date_time_utc: "a",
                status: "Completed",
              },
              {
                policy_id: 21,
                policy_name: "New",
                username: "",
                date_time_epoch: 3000,
                date_time_utc: "c",
                status: "Failed",
              },
              {
                policy_id: 22,
                policy_name: "Mid",
                username: "ada",
                date_time_epoch: 2000,
                date_time_utc: "b",
                status: "Completed",
              },
            ],
          },
        },
      };
    if (path === "/api/v2/mdm/commands")
      return { status: 201, data: [{ id: "cmd-1", href: "/x" }] };
    if (path === "/api/v2/mdm/blank-push")
      return { status: 200, data: { errorUuids: options.errorUuids ?? [] } };
    return {
      status: 404,
      data: { httpStatus: 404, errors: [{ code: "NOT_FOUND", description: "nope" }] },
    };
  };
  const ctx = {
    limits,
    config: {},
    call,
    forbidden: options.forbidden,
    async audit(e) {
      audit.push(e);
    },
  } as ToolContext;
  const tools = new Map<string, { schema: z.ZodType; fn: (args: unknown) => Promise<any> }>();
  const server = {
    registerTool: (name: string, def: { inputSchema: z.ZodType }, fn: never) =>
      tools.set(name, { schema: def.inputSchema, fn }),
  };
  jamfTools.register(server as never, ctx, new Set(jamfTools.names));
  const run = async (name: string, args: Record<string, unknown>) => {
    const tool = tools.get(name)!;
    const r = await tool.fn(tool.schema.parse(args));
    const text = r.content[0].text as string;
    return { isError: Boolean(r.isError), text, body: r.isError ? null : JSON.parse(text) };
  };
  const sent = (method: string) => requests.filter((r) => r.method === method);
  return { run, requests, audit, sent };
}

const sites1 = { sites: ["1"] };
const query = (req: VendorRequest) => new URLSearchParams(req.path.split("?")[1]);

test("jamf_search_devices lists computers trimmed to what an agent needs", async () => {
  const { run, requests } = setup();
  const r = await run("jamf_search_devices", { kind: "computer", max_results: 2 });
  assert.equal(r.isError, false);
  assert.deepEqual(r.body.devices[0], {
    id: "1",
    kind: "computer",
    name: "Mac 1",
    serial_number: "C01",
    model: "MacBook Pro",
    os: "macOS",
    os_version: "15.1",
    user: "user1",
    user_name: "User 1",
    email: "user1@x.com",
    last_contact: "2026-10-11T08:00:00Z",
    last_inventory: "2026-10-20T08:00:00Z",
    managed: true,
    supervised: true,
    filevault: "ALL_ENCRYPTED",
    site_id: "1",
  });
  assert.ok(!r.text.includes("secret-ea") && !r.text.includes("macAddress"));
  assert.equal(r.body.next, "1");
  const q = query(requests[0]);
  assert.equal(requests[0].path.split("?")[0], "/api/v4/computers-inventory");
  assert.deepEqual(q.getAll("section"), [
    "GENERAL",
    "HARDWARE",
    "OPERATING_SYSTEM",
    "USER_AND_LOCATION",
  ]);
  assert.equal(q.get("page"), "0");
  assert.equal(q.get("page-size"), "2");
  assert.equal(q.get("sort"), "general.name:asc");
  // The next page.
  const next = await run("jamf_search_devices", { kind: "computer", max_results: 2, cursor: "1" });
  assert.deepEqual(
    next.body.devices.map((d: { id: string }) => d.id),
    ["3", "4"],
  );
  assert.equal(next.body.next, null);
});

test("jamf_search_devices searches name, serial, username and email in one RSQL filter", async () => {
  const { run, requests } = setup();
  const r = await run("jamf_search_devices", { kind: "computer", search: "C02" });
  assert.deepEqual(
    r.body.devices.map((d: { id: string }) => d.id),
    ["2"],
  );
  assert.equal(
    query(requests[0]).get("filter"),
    'general.name=="*C02*",hardware.serialNumber=="C02",userAndLocation.username=="C02",userAndLocation.email=="C02"',
  );
  await run("jamf_search_devices", { kind: "mobile", search: "Ada O'Neil" });
  assert.equal(
    query(requests[1]).get("filter"),
    `displayName=="*Ada O'Neil*",serialNumber=="Ada O'Neil",username=="Ada O'Neil",emailAddress=="Ada O'Neil"`,
  );
});

test("jamf_search_devices refuses search text that could change the filter, before any call", async () => {
  const { run, requests, audit } = setup();
  for (const bad of ['a"b', "a*", "x,id==1", "a;b", "a\\b", "a(b)", "a=b"]) {
    const r = await run("jamf_search_devices", { kind: "computer", search: bad });
    assert.equal(r.isError, true, bad);
  }
  assert.equal(requests.length, 0);
  assert.ok(audit.every((a) => a.decision === "denied"));
});

test("jamf_search_devices under a sites limit keeps only those sites, reading on a few pages", async () => {
  const { run, requests } = setup(sites1);
  const r = await run("jamf_search_devices", { kind: "computer", max_results: 1 });
  // Page 0 has computer 1 (site 1): enough.
  assert.deepEqual(
    r.body.devices.map((d: { id: string }) => d.id),
    ["1"],
  );
  assert.equal(r.body.next, "1");
  // From page 1: 2 (site 2) and 3 (no site) are skipped, 4 (site 1) is kept.
  const more = await run("jamf_search_devices", { kind: "computer", max_results: 1, cursor: "1" });
  assert.deepEqual(
    more.body.devices.map((d: { id: string }) => d.id),
    ["4"],
  );
  assert.equal(more.body.next, null);
  assert.equal(requests.length, 4);
  // Never more than MAX_SCOPED_PAGES pages in one call.
  const none = setup({ sites: ["9"] });
  const empty = await none.run("jamf_search_devices", { kind: "computer", max_results: 1 });
  assert.deepEqual(empty.body.devices, []);
  assert.ok(none.requests.length <= MAX_SCOPED_PAGES);
});

test("jamf_get_device answers the detail, its groups and, on request, its apps", async () => {
  const { run, requests } = setup();
  const r = await run("jamf_get_device", { kind: "computer", device_id: "1" });
  assert.equal(r.body.site, "Site 1");
  assert.equal(r.body.last_enrolled, "2026-01-01T00:00:00Z");
  assert.deepEqual(r.body.groups, [{ id: "5", name: "Pilot", smart: false }]);
  assert.equal(r.body.apps, undefined);
  assert.equal(query(requests[0]).get("filter"), "id=in=(1)");
  assert.ok(!query(requests[0]).getAll("section").includes("APPLICATIONS"));
  const apps = await run("jamf_get_device", {
    kind: "computer",
    device_id: "1",
    include_apps: true,
  });
  assert.deepEqual(apps.body.apps, [
    { name: "Safari.app", version: "18.0", bundle_id: "com.apple.Safari" },
  ]);
  assert.equal(apps.body.apps_total, 1);
  assert.ok(query(requests[1]).getAll("section").includes("APPLICATIONS"));
});

test("jamf_get_device reads a mobile device", async () => {
  const { run, requests } = setup();
  const r = await run("jamf_get_device", { kind: "mobile", device_id: "10", include_apps: true });
  assert.deepEqual(r.body, {
    id: "10",
    kind: "mobile",
    type: "iOS",
    name: "iPad 10",
    serial_number: "M10",
    model: "iPad Air",
    os_version: "18.1",
    user: "kid10",
    user_name: "Kid 10",
    email: "kid10@x.com",
    last_contact: "2026-10-21T08:00:00Z",
    last_inventory: "2026-10-21T07:00:00Z",
    managed: true,
    supervised: true,
    encrypted: true,
    passcode: true,
    lost_mode: false,
    site_id: "1",
    last_enrolled: "2026-02-01T00:00:00Z",
    groups: [{ id: "8", name: "Carts", smart: true }],
    apps: [{ name: "Notes", version: "1.2", bundle_id: "com.x.notes" }],
    apps_total: 1,
  });
  assert.equal(requests[0].path.split("?")[0], "/api/v2/mobile-devices/detail");
  assert.equal(query(requests[0]).get("filter"), "mobileDeviceId=in=(10)");
});

test("jamf_get_device refuses a device outside the sites as it refuses one that does not exist", async () => {
  const { run, requests, audit } = setup(sites1);
  const outside = await run("jamf_get_device", { kind: "computer", device_id: "2" });
  const noSite = await run("jamf_get_device", { kind: "computer", device_id: "3" });
  const missing = await run("jamf_get_device", { kind: "computer", device_id: "99" });
  assert.equal(outside.text, "No computer 2 that this agent may use");
  assert.equal(noSite.text, "No computer 3 that this agent may use");
  assert.equal(missing.text, "No computer 99 that this agent may use");
  assert.ok(audit.every((a) => a.decision === "denied"));
  // A malformed id never reaches Jamf.
  const n = requests.length;
  assert.equal((await run("jamf_get_device", { kind: "computer", device_id: "1)" })).isError, true);
  assert.equal(requests.length, n);
});

test("jamf_list_groups lists groups of the key's sites only", async () => {
  const { run, requests } = setup(sites1);
  const r = await run("jamf_list_groups", { kind: "computer", type: "static", search: "Pi" });
  assert.deepEqual(r.body.groups, [
    { id: "5", name: "Pilot", description: null, site_id: "1", members: 3 },
  ]);
  assert.equal(requests[0].path.split("?")[0], "/api/v3/computer-groups/static-groups");
  assert.equal(query(requests[0]).get("filter"), 'name=="*Pi*"');
  const m = await run("jamf_list_groups", { kind: "mobile", type: "smart" });
  assert.deepEqual(m.body.groups, [
    { id: "9", name: "Smart iPads", description: null, site_id: "1", members: 4 },
  ]);
  assert.equal(query(requests[1]).get("sort"), "groupName:asc");
});

test("jamf_get_group answers its members, leaving out any outside the sites", async () => {
  const { run, requests } = setup(sites1);
  const r = await run("jamf_get_group", { kind: "computer", group_id: "5" });
  assert.equal(r.body.name, "Pilot");
  assert.equal(r.body.smart, false);
  assert.equal(r.body.member_count, 3);
  assert.deepEqual(
    r.body.members.map((d: { id: string }) => d.id),
    ["1", "4"],
  );
  assert.equal(requests[0].path, "/JSSResource/computergroups/id/5");
  assert.equal(query(requests[1]).get("filter"), "id=in=(1,4,2)");
  // Paged by position in the group.
  const page = await run("jamf_get_group", {
    kind: "computer",
    group_id: "5",
    max_results: 1,
    cursor: "1",
  });
  assert.deepEqual(
    page.body.members.map((d: { id: string }) => d.id),
    ["4"],
  );
  assert.equal(page.body.next, "2");
});

test("jamf_get_group refuses a group of another site or one Jamf does not have, alike", async () => {
  const { run } = setup(sites1);
  assert.equal(
    (await run("jamf_get_group", { kind: "computer", group_id: "6" })).text,
    "No computer group 6 that this agent may use",
  );
  assert.equal(
    (await run("jamf_get_group", { kind: "computer", group_id: "60" })).text,
    "No computer group 60 that this agent may use",
  );
});

test("jamf_change_group_members adds and removes through the Classic API's additions and deletions", async () => {
  const { run, sent, audit } = setup(sites1);
  const r = await run("jamf_change_group_members", {
    kind: "computer",
    group_id: "5",
    add: ["4"],
    remove: ["1"],
  });
  assert.equal(r.isError, false);
  assert.deepEqual(r.body, { group_id: "5", name: "Pilot", added: ["4"], removed: ["1"] });
  const [put] = sent("PUT");
  assert.equal(put.path, "/JSSResource/computergroups/id/5");
  assert.equal(put.headers?.["content-type"], "application/xml");
  assert.equal(
    put.body,
    "<computer_group><computer_additions><computer><id>4</id></computer></computer_additions>" +
      "<computer_deletions><computer><id>1</id></computer></computer_deletions></computer_group>",
  );
  assert.equal(audit.at(-1)!.target, "computer group 5: added 4; removed 1");
  // Mobile groups use their own element names.
  const m = setup();
  await m.run("jamf_change_group_members", { kind: "mobile", group_id: "8", add: ["11"] });
  assert.equal(
    m.sent("PUT")[0].body,
    "<mobile_device_group><mobile_device_additions><mobile_device><id>11</id></mobile_device></mobile_device_additions></mobile_device_group>",
  );
});

test("jamf_change_group_members refuses before any change", async () => {
  const { run, sent, audit } = setup(sites1);
  const refused = async (args: Record<string, unknown>, why: RegExp) => {
    const r = await run("jamf_change_group_members", { kind: "computer", ...args });
    assert.equal(r.isError, true);
    assert.match(r.text, why);
  };
  await refused({ group_id: "5" }, /at least one device/);
  await refused({ group_id: "5", add: ["1"], remove: ["1"] }, /Both added and removed: 1/);
  await refused({ group_id: "7", add: ["1"] }, /smart group/);
  await refused({ group_id: "6", add: ["1"] }, /No computer group 6/);
  // A device of another site, one with no site and one Jamf does not have.
  await refused(
    { group_id: "5", add: ["2", "3", "99"] },
    /No computers 2, 3, 99 that this agent may use/,
  );
  await refused({ group_id: "5", add: ["x1"] }, /x1 is not a computer id/);
  assert.equal(sent("PUT").length, 0);
  assert.ok(audit.every((a) => a.decision === "denied"));
});

test("jamf_send_blank_push pushes each device's management id and names the ones that failed", async () => {
  const { run, sent } = setup({}, { errorUuids: ["mm-11"] });
  const r = await run("jamf_send_blank_push", { kind: "mobile", device_ids: ["10", "11", "10"] });
  assert.deepEqual(r.body, { sent: ["10"], failed: ["11"] });
  const [post] = sent("POST");
  assert.equal(post.path, "/api/v2/mdm/blank-push");
  assert.deepEqual(post.body, { clientManagementIds: ["mm-10", "mm-11"] });
});

test("jamf_send_blank_push refuses a device not enrolled in MDM or outside the sites", async () => {
  const { run, sent } = setup(sites1);
  assert.match(
    (await run("jamf_send_blank_push", { kind: "computer", device_ids: ["4"] })).text,
    /no management id for computer 4/,
  );
  assert.match(
    (await run("jamf_send_blank_push", { kind: "computer", device_ids: ["1", "2"] })).text,
    /No computers 2 that this agent may use/,
  );
  assert.equal(sent("POST").length, 0);
});

test("jamf_restart_device queues RESTART_DEVICE, notifyUser for computers only", async () => {
  const { run, sent } = setup();
  const r = await run("jamf_restart_device", { kind: "computer", device_id: "1" });
  assert.deepEqual(r.body, {
    device_id: "1",
    name: "Mac 1",
    command: "RESTART_DEVICE",
    command_id: "cmd-1",
    status: "queued",
    note: "Jamf Pro sent the command; the device carries it out when it next checks in.",
  });
  assert.deepEqual(sent("POST")[0].body, {
    clientData: [{ managementId: "mc-1" }],
    commandData: { commandType: "RESTART_DEVICE", notifyUser: true },
  });
  await run("jamf_restart_device", { kind: "mobile", device_id: "10" });
  assert.deepEqual(sent("POST")[1].body, {
    clientData: [{ managementId: "mm-10" }],
    commandData: { commandType: "RESTART_DEVICE" },
  });
});

test("jamf_lock_device needs a pin for a computer and none for a mobile device", async () => {
  const { run, sent } = setup();
  assert.match(
    (await run("jamf_lock_device", { kind: "computer", device_id: "1" })).text,
    /6-digit pin/,
  );
  assert.match(
    (await run("jamf_lock_device", { kind: "computer", device_id: "1", pin: "12a456" })).text,
    /6-digit pin/,
  );
  assert.match(
    (await run("jamf_lock_device", { kind: "mobile", device_id: "10", pin: "123456" })).text,
    /own passcode/,
  );
  assert.equal(sent("POST").length, 0);
  await run("jamf_lock_device", {
    kind: "computer",
    device_id: "1",
    pin: "123456",
    message: "Call IT",
    phone_number: "555",
  });
  assert.deepEqual((sent("POST")[0].body as { commandData: unknown }).commandData, {
    commandType: "DEVICE_LOCK",
    pin: "123456",
    message: "Call IT",
    phoneNumber: "555",
  });
});

test("jamf_erase_device erases only when the serial number matches the device", async () => {
  const { run, sent, audit } = setup(sites1);
  const wrong = await run("jamf_erase_device", {
    kind: "computer",
    device_id: "1",
    confirm_serial_number: "C04",
  });
  assert.match(wrong.text, /does not match computer 1/);
  assert.equal(sent("POST").length, 0);
  // Outside the sites: refused before the serial is even compared.
  assert.match(
    (
      await run("jamf_erase_device", {
        kind: "computer",
        device_id: "2",
        confirm_serial_number: "C02",
      })
    ).text,
    /No computer 2/,
  );
  const r = await run("jamf_erase_device", {
    kind: "computer",
    device_id: "1",
    confirm_serial_number: " c01 ",
    pin: "654321",
  });
  assert.equal(r.body.command, "ERASE_DEVICE");
  assert.deepEqual((sent("POST")[0].body as { commandData: unknown }).commandData, {
    commandType: "ERASE_DEVICE",
    pin: "654321",
  });
  assert.equal(audit.at(-1)!.target, "computer 1 (serial C01)");
  assert.equal(audit.at(-1)!.decision, "allowed");
  assert.match(
    (
      await run("jamf_erase_device", {
        kind: "mobile",
        device_id: "10",
        confirm_serial_number: "M10",
        pin: "123456",
      })
    ).text,
    /for computers only/,
  );
});

test("jamf_list_policies lists by name, paged, and under a sites limit only those sites'", async () => {
  const { run, requests } = setup();
  const all = await run("jamf_list_policies", { max_results: 2 });
  assert.deepEqual(all.body, {
    policies: [
      { id: "20", name: "Policy 20" },
      { id: "21", name: "Policy 21" },
    ],
    next: "2",
  });
  assert.equal(requests.length, 1);
  assert.deepEqual((await run("jamf_list_policies", { search: "policy 22" })).body.policies, [
    { id: "22", name: "Policy 22" },
  ]);
  const scoped = setup(sites1);
  const r = await scoped.run("jamf_list_policies", {});
  assert.deepEqual(r.body.policies, [{ id: "20", name: "Policy 20" }]);
  assert.ok(scoped.requests.slice(1).every((q) => q.path.endsWith("/subset/General")));
});

test("jamf_list_policies under a sites limit reads at most MAX_POLICY_READS policies a call", async () => {
  const many = Array.from({ length: MAX_POLICY_READS + 10 }, (_, i) => ({
    id: 100 + i,
    name: `P${i}`,
    site: 2,
  }));
  const { run, requests } = setup(sites1, { policies: many });
  const r = await run("jamf_list_policies", {});
  assert.deepEqual(r.body.policies, []);
  assert.equal(r.body.next, String(MAX_POLICY_READS));
  assert.equal(requests.length, 1 + MAX_POLICY_READS);
});

test("jamf_get_policy answers its triggers, scope, packages and scripts", async () => {
  const { run, requests } = setup();
  const r = await run("jamf_get_policy", { policy_id: "20" });
  assert.deepEqual(r.body, {
    id: "20",
    name: "Policy 20",
    enabled: true,
    frequency: "Once per computer",
    triggers: ["checkin", "install-office"],
    category: "Apps",
    site: "Site 1",
    scope: {
      all_computers: false,
      computer_groups: ["Pilot"],
      computers: 2,
      excluded_computer_groups: ["Other"],
    },
    packages: ["Office.pkg"],
    scripts: ["cleanup.sh"],
    self_service: true,
  });
  assert.equal(requests[0].path, "/JSSResource/policies/id/20");
  const scoped = setup(sites1);
  assert.equal(
    (await scoped.run("jamf_get_policy", { policy_id: "21" })).text,
    "No policy 21 that this agent may use",
  );
  assert.equal(
    (await scoped.run("jamf_get_policy", { policy_id: "99" })).text,
    "No policy 99 that this agent may use",
  );
  assert.equal((await scoped.run("jamf_get_policy", { policy_id: "../1" })).isError, true);
});

test("jamf_computer_policy_logs answers a computer's policy runs, newest first", async () => {
  const { run, requests } = setup(sites1);
  const r = await run("jamf_computer_policy_logs", { computer_id: "1", max_results: 2 });
  assert.deepEqual(r.body, {
    logs: [
      { policy_id: 21, policy_name: "New", status: "Failed", completed: "c", username: null },
      { policy_id: 22, policy_name: "Mid", status: "Completed", completed: "b", username: "ada" },
    ],
    next: "2",
  });
  assert.equal(requests[1].path, "/JSSResource/computerhistory/id/1/subset/PolicyLogs");
  const outside = await run("jamf_computer_policy_logs", { computer_id: "2" });
  assert.equal(outside.text, "No computer 2 that this agent may use");
  assert.ok(!requests.some((q) => q.path.includes("computerhistory/id/2")));
});

test("Jamf errors reach the agent in plain text, a 403 with its capability's hint", async () => {
  const { run, audit } = setup(
    {},
    {
      answers: {
        "/api/v2/mdm/commands": {
          status: 403,
          data: {
            httpStatus: 403,
            errors: [{ code: "INVALID_PRIVILEGE", description: "Forbidden" }],
          },
        },
        "/JSSResource/policies": {
          status: 500,
          data: "<html><body><p>Server error</p></body></html>",
        },
      },
      forbidden: { jamf_restart_device: "ADD PRIVILEGE" },
    },
  );
  const r = await run("jamf_restart_device", { kind: "computer", device_id: "1" });
  assert.equal(r.text, "Jamf Pro answered 403: Forbidden. ADD PRIVILEGE");
  assert.equal(audit.at(-1)!.decision, "failed");
  const p = await run("jamf_list_policies", {});
  assert.equal(p.text, "Jamf Pro answered 500: Server error");
});

test("a 403 on the device read names the read capability's privileges, not the command's", async () => {
  const { run } = setup(
    {},
    {
      answers: {
        "/api/v4/computers-inventory": {
          status: 403,
          data: { errors: [{ description: "Forbidden" }] },
        },
      },
      forbidden: { jamf_get_device: "READ HINT", jamf_restart_device: "COMMAND HINT" },
    },
  );
  const r = await run("jamf_restart_device", { kind: "computer", device_id: "1" });
  assert.equal(r.text, "Jamf Pro answered 403: Forbidden. READ HINT");
});
