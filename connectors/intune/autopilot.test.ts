import assert from "node:assert/strict";
import { test } from "node:test";
import { intuneTools } from "./tools.ts";
import { loadConnectors } from "../../packages/connector-kit/src/connector.ts";
import type { ToolContext } from "../../packages/connector-kit/src/common.ts";
import type { VendorRequest, VendorResponse } from "../../packages/connector-kit/src/types.ts";

// intune_autopilot_status against a fake Graph: the tool's own rules (serial check, the
// groups limit, what is answered), not the MCP plumbing, which mcp.test.ts covers.

const GROUP = "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa";
const ADA = "11111111-0000-0000-0000-000000000001";
const BOB = "11111111-0000-0000-0000-000000000002";
const MD_ADA = "22222222-0000-0000-0000-000000000001";
const MD_BOB = "22222222-0000-0000-0000-000000000002";
const MD_SPARE = "22222222-0000-0000-0000-000000000005";
const NONE = "00000000-0000-0000-0000-000000000000";

const identity = (serial: string, extra: Record<string, unknown> = {}) => ({
  id: `33333333-0000-0000-0000-${serial.padStart(12, "0").slice(-12)}`,
  serialNumber: serial,
  enrollmentState: "notContacted",
  groupTag: "",
  lastContactedDateTime: "0001-01-01T00:00:00Z",
  userPrincipalName: "",
  managedDeviceId: NONE,
  ...extra,
});

const IDENTITIES = [
  identity("1000", {
    enrollmentState: "enrolled",
    groupTag: "Sales",
    lastContactedDateTime: "2026-10-20T10:00:00Z",
    managedDeviceId: MD_ADA,
  }),
  identity("2000", { managedDeviceId: MD_BOB }),
  identity("3000"),
  identity("30001"),
  identity("4000", { userPrincipalName: "bob@x.com" }),
  // Points to a managed device Intune no longer has (404).
  identity("6000", { managedDeviceId: "22222222-0000-0000-0000-000000000099" }),
  // A spare: enrolled and managed, nobody signed in on it, nobody assigned.
  identity("5000", {
    enrollmentState: "enrolled",
    groupTag: "Spares",
    lastContactedDateTime: "2026-10-24T09:00:00Z",
    managedDeviceId: MD_SPARE,
  }),
];
const MANAGED: Record<string, object> = {
  [MD_ADA]: {
    id: MD_ADA,
    userId: ADA,
    userPrincipalName: "ada@x.com",
    complianceState: "compliant",
  },
  [MD_BOB]: {
    id: MD_BOB,
    userId: BOB,
    userPrincipalName: "bob@x.com",
    complianceState: "noncompliant",
  },
  [MD_SPARE]: {
    id: MD_SPARE,
    deviceName: "SPARE-0042",
    userId: null,
    userPrincipalName: "",
    complianceState: "compliant",
  },
};

type Options = {
  identities?: ReturnType<typeof identity>[];
  /** The Autopilot list's page size; next pages are linked with `$skip`, not `$skiptoken`. */
  pageSize?: number;
  /** Where the next page link points (default Graph). */
  nextOrigin?: string;
  /** Answers instead of the fake's own. */
  managed?: VendorResponse;
  profile?: VendorResponse;
  forbidden?: Record<string, string>;
};

function setup(
  limits: Record<string, unknown> = {},
  config: Record<string, string> = {},
  options: Options = {},
) {
  const { identities = IDENTITIES, pageSize = 50 } = options;
  const requests: VendorRequest[] = [];
  const audit: { tool: string; decision: string; reason: string | null }[] = [];
  const call = async (req: VendorRequest): Promise<VendorResponse> => {
    requests.push(req);
    const [path, query = ""] = decodeURIComponent(req.path).split("?");
    if (path === "/v1.0/deviceManagement/windowsAutopilotDeviceIdentities") {
      const m = /contains\(serialNumber,'(.*)'\)/.exec(query);
      const all = identities.filter((i) => m && i.serialNumber.includes(m[1]));
      const skip = Number(new URLSearchParams(query).get("$skip") ?? 0);
      const more = skip + pageSize < all.length;
      const next = `${options.nextOrigin ?? "https://graph.microsoft.com"}${path}?${query.replace(/&\$skip=\d+/, "")}&$skip=${skip + pageSize}`;
      return {
        status: 200,
        data: {
          value: all.slice(skip, skip + pageSize),
          ...(more ? { "@odata.nextLink": next } : {}),
        },
      };
    }
    const md = /managedDevices\/([^/]+)$/.exec(path)?.[1];
    if (md && options.managed) return options.managed;
    if (md) return MANAGED[md] ? { status: 200, data: MANAGED[md] } : { status: 404, data: {} };
    if (path.endsWith("/deploymentProfile") && options.profile) return options.profile;
    const profile =
      /beta\/deviceManagement\/windowsAutopilotDeviceIdentities\/([^/]+)\/deploymentProfile$/.exec(
        path,
      );
    if (profile)
      return profile[1].endsWith("1000")
        ? { status: 200, data: { displayName: "Laptops pre-provisioned" } }
        : { status: 404, data: { error: { message: "no profile" } } };
    const user = /users\/([^/]+)\/checkMemberObjects$/.exec(path)?.[1];
    if (user) {
      const inGroup = [ADA, "ada@x.com"].includes(user);
      return inGroup
        ? { status: 200, data: { value: [GROUP] } }
        : user === "ghost@x.com"
          ? { status: 404, data: {} }
          : { status: 200, data: { value: [] } };
    }
    return { status: 500, data: {} };
  };
  const ctx = {
    limits,
    config,
    call,
    forbidden: options.forbidden,
    async audit(e) {
      audit.push(e);
    },
  } as ToolContext;
  const tools = new Map<string, (args: unknown) => Promise<any>>();
  const server = { registerTool: (name: string, _def: unknown, fn: never) => tools.set(name, fn) };
  const register = (allowed: string[]) =>
    intuneTools.register(server as never, ctx, new Set(allowed));
  register(["intune_autopilot_status"]);
  const run = async (serial_number: string) => {
    const r = await tools.get("intune_autopilot_status")!({ serial_number });
    return { isError: Boolean(r.isError), body: r.content[0].text as string };
  };
  return { run, requests, audit, tools, register };
}

const groups = { groups: [GROUP] };
const on = { people_groups: "groups" };

test("a found enrolled device with its user answers the whole shape", async () => {
  const { run } = setup();
  const r = await run("1000");
  assert.equal(r.isError, false);
  assert.deepEqual(JSON.parse(r.body), {
    serial_number: "1000",
    found: true,
    deployment_state: "enrolled",
    enrollment_profile: "Laptops pre-provisioned",
    group_tag: "Sales",
    last_contact: "2026-10-20T10:00:00Z",
    managed_device: { id: MD_ADA, compliance: "compliant", primary_user: "ada@x.com" },
  });
});

test("a found device that is not enrolled has no managed device", async () => {
  const { run, requests } = setup();
  assert.deepEqual(JSON.parse((await run("3000")).body), {
    serial_number: "3000",
    found: true,
    deployment_state: "notContacted",
    enrollment_profile: null,
    group_tag: null,
    last_contact: "0001-01-01T00:00:00Z",
    managed_device: null,
  });
  // The serial 30001 also contains 3000: only the exact one is answered.
  assert.ok(!requests.some((r) => r.path.includes("managedDevices")));
});

test("a serial Autopilot does not know is found: false", async () => {
  const { run } = setup();
  assert.deepEqual(JSON.parse((await run("9999")).body), { serial_number: "9999", found: false });
  // A serial that only part-matches another is not that device.
  assert.deepEqual(JSON.parse((await run("300")).body), { serial_number: "300", found: false });
});

test("an implausible serial is refused before Graph is called", async () => {
  const { run, requests, audit } = setup();
  for (const bad of ["", "a b", "x'; drop", "a".repeat(51), "1000') or ('1'='1", "AB'12", "é1"]) {
    const r = await run(bad);
    assert.equal(r.isError, true, bad);
  }
  assert.equal(requests.length, 0);
  assert.ok(audit.every((a) => a.decision === "denied"));
});

test("the serial reaches the filter as a quoted contains() literal", async () => {
  // A quote itself never gets this far: the serial check refuses it (test above).
  const { run, requests } = setup();
  await run("AB-12");
  const path = decodeURIComponent(requests[0].path);
  assert.match(path, /\$filter=contains\(serialNumber,'AB-12'\)/);
});

test("a serial with spaces, dots or underscores is refused with why", async () => {
  const { run, requests } = setup();
  for (const vm of ["0000-0000 1234", "VMware-56.4d", "VM_1"]) {
    const r = await run(vm);
    assert.equal(r.isError, true, vm);
    assert.match(r.body, /spaces, dots or underscores/);
  }
  assert.equal(requests.length, 0);
});

test("a managed device id Intune no longer has (404) answers no managed device", async () => {
  const { run } = setup();
  const r = await run("6000");
  assert.equal(r.isError, false);
  assert.equal(JSON.parse(r.body).managed_device, null);
});

test("a 403 on the managed device read carries the read capability's hint", async () => {
  const { run, audit } = setup(
    {},
    {},
    {
      managed: { status: 403, data: { error: { message: "Forbidden" } } },
      forbidden: { intune_get_device: "READ HINT", intune_autopilot_status: "AUTOPILOT HINT" },
    },
  );
  const r = await run("1000");
  assert.equal(r.isError, true);
  assert.equal(r.body, "Microsoft Graph answered 403: Forbidden. READ HINT");
  assert.equal(audit.at(-1)!.decision, "failed");
});

for (const status of [403, 500, 503]) {
  test(`a ${status} on the beta profile read keeps the v1.0 answer, profile null with why`, async () => {
    const { run } = setup({}, {}, { profile: { status, data: { error: { message: "no" } } } });
    const r = await run("1000");
    assert.equal(r.isError, false);
    const body = JSON.parse(r.body);
    assert.equal(body.enrollment_profile, null);
    assert.equal(
      body.enrollment_profile_error,
      `Microsoft Graph (beta) answered ${status} for the profile`,
    );
    assert.equal(body.deployment_state, "enrolled");
    assert.equal(body.managed_device.id, MD_ADA);
  });
}

test("under a groups limit a device of someone in the groups is answered", async () => {
  const { run } = setup(groups, on);
  assert.equal(JSON.parse((await run("1000")).body).managed_device.primary_user, "ada@x.com");
});

test("under a groups limit a device of someone outside is refused like intune_get_device", async () => {
  const { run, audit } = setup(groups, on);
  const r = await run("2000");
  assert.equal(r.isError, true);
  assert.equal(r.body, "This agent may not read that device");
  assert.equal(audit.at(-1)!.decision, "denied");
  // Assigned on the Autopilot record to someone outside, not yet enrolled: the same refusal.
  assert.equal((await run("4000")).body, "This agent may not read that device");
});

const LIMITED =
  "No person is tied to this device and this key is kept to groups: only its Autopilot state and compliance are answered.";

test("under a groups limit a device nobody is tied to answers only the minimum", async () => {
  const { run, requests } = setup(groups, on);
  // Managed, compliant, no user: no managed device id, name, group tag, contact or profile.
  assert.deepEqual(JSON.parse((await run("5000")).body), {
    serial_number: "5000",
    found: true,
    deployment_state: "enrolled",
    managed: true,
    compliance: "compliant",
    limited: LIMITED,
  });
  // Not yet managed.
  assert.deepEqual(JSON.parse((await run("3000")).body), {
    serial_number: "3000",
    found: true,
    deployment_state: "notContacted",
    managed: false,
    compliance: null,
    limited: LIMITED,
  });
  assert.ok(!requests.some((r) => r.path.includes("deploymentProfile")));
  assert.ok(!requests.some((r) => r.path.includes("checkMemberObjects")));
});

test("without a groups limit a device nobody is tied to answers the whole shape", async () => {
  const { run } = setup();
  const r = JSON.parse((await run("5000")).body);
  assert.equal(r.group_tag, "Spares");
  assert.deepEqual(r.managed_device, { id: MD_SPARE, compliance: "compliant", primary_user: null });
});

test("a match on a later page is found by following Graph's next link as given", async () => {
  // Graph links the next page with $skip here, not $skiptoken.
  // Six serials that contain 70006 and the exact one last.
  const many = [...Array.from({ length: 6 }, (_, n) => identity(`70006${n}`)), identity("70006")];
  const { run, requests } = setup({}, {}, { identities: many, pageSize: 4 });
  const r = JSON.parse((await run("70006")).body);
  assert.equal(r.found, true);
  assert.equal(r.serial_number, "70006");
  const lists = requests.filter((q) => q.path.includes("windowsAutopilotDeviceIdentities?"));
  assert.equal(lists.length, 2);
  assert.match(decodeURIComponent(lists[1].path), /^\/v1\.0\/.*\$skip=4/);
  assert.match(decodeURIComponent(lists[1].path), /contains\(serialNumber,'70006'\)/);
  // Not in any page: found: false after the last one.
  assert.deepEqual(JSON.parse((await run("7000")).body), { serial_number: "7000", found: false });
});

test("more matching pages than the cap is refused after the cap", async () => {
  const many = Array.from({ length: 30 }, (_, n) => identity(`80${String(n).padStart(2, "0")}9`));
  const { run, requests } = setup({}, {}, { identities: many, pageSize: 2 });
  const r = await run("80");
  assert.equal(r.isError, true);
  assert.equal(
    r.body,
    "Too many Autopilot devices match that serial number. Give the full serial number.",
  );
  assert.equal(requests.length, 5);
});

test("a next page link off Graph is never followed", async () => {
  const many = [identity("910020"), identity("91002")];
  const { run, requests } = setup(
    {},
    {},
    {
      identities: many,
      pageSize: 1,
      nextOrigin: "https://graph.evil.example",
    },
  );
  const r = await run("91002");
  assert.equal(r.isError, true);
  assert.equal(requests.length, 1);
});

test("a groups limit Intune is not set up to check refuses before any call", async () => {
  const { run, requests } = setup(groups, {});
  assert.equal((await run("3000")).isError, true);
  assert.equal(requests.length, 0);
});

test("a key without the autopilot capability is never offered the tool", async () => {
  const { tools, register } = setup();
  tools.clear();
  register(["intune_list_devices"]);
  assert.ok(!tools.has("intune_autopilot_status"));
});

test("the connector grants the tool through its own read-only capability", () => {
  const intune = loadConnectors().find((c) => c.id === "intune")!;
  const cap = intune.capabilities.find((c) => c.tools.includes("intune_autopilot_status"))!;
  assert.equal(cap.id, "autopilot");
  assert.deepEqual(cap.scopes, ["DeviceManagementServiceConfig.Read.All"]);
  assert.ok(!cap.required);
  assert.ok(!("destructive" in cap && cap.destructive));
});
