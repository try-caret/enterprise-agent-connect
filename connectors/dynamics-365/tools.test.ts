import assert from "node:assert/strict";
import { test } from "node:test";
import { dynamicsTools, literal } from "./tools.ts";
import { loadConnectors } from "../../packages/connector-kit/src/connector.ts";
import type { ToolContext } from "../../packages/connector-kit/src/common.ts";
import type { VendorRequest, VendorResponse } from "../../packages/connector-kit/src/types.ts";

// The Dynamics 365 tools against a fake Dataverse Web API (answers shaped as Dataverse returns
// them, annotations included). Business unit EAST holds Contoso, its contact, opportunity and
// case; WEST holds Fabrikam, owned by team WEST_TEAM. The application user is in EAST.

const guid = (n: number) => `00000000-0000-0000-0000-${String(n).padStart(12, "0")}`;
const EAST = guid(1);
const WEST = guid(2);
const WEST_TEAM = guid(3);
const EAST_TEAM = guid(4);
const CONTOSO = guid(10);
const FABRIKAM = guid(11);
const ADA = guid(20);
const DEAL = guid(30);
const WEST_DEAL = guid(31);
const CASE = guid(40);
const TASK = guid(50);
const WEST_TASK = guid(51);
const NOTE = guid(60);
const PROJECT = guid(70);
const API = "/api/data/v9.2";
const F = "@OData.Community.Display.V1.FormattedValue";
const L = "@Microsoft.Dynamics.CRM.lookuplogicalname";

const owned = (unit: string, team: string | null = null) => ({
  _owningbusinessunit_value: unit,
  _owningteam_value: team,
  _ownerid_value: team ?? guid(99),
  [`_ownerid_value${F}`]: team ? "West team" : "Agents app",
  [`_ownerid_value${L}`]: team ? "team" : "systemuser",
});
const RECORDS: Record<string, Record<string, Record<string, unknown>>> = {
  accounts: {
    [CONTOSO]: {
      accountid: CONTOSO,
      name: "Contoso",
      statecode: 0,
      [`statecode${F}`]: "Active",
      revenue: 1000000,
      [`revenue${F}`]: "$1,000,000.00",
      _primarycontactid_value: ADA,
      [`_primarycontactid_value${F}`]: "Ada Lovelace",
      [`_primarycontactid_value${L}`]: "contact",
      _parentaccountid_value: null,
      description: "x".repeat(2500),
      "@odata.etag": 'W/"1"',
      ...owned(EAST),
    },
    [FABRIKAM]: { accountid: FABRIKAM, name: "Fabrikam", ...owned(WEST, WEST_TEAM) },
  },
  contacts: { [ADA]: { contactid: ADA, fullname: "Ada Lovelace", ...owned(EAST) } },
  opportunities: {
    [DEAL]: { opportunityid: DEAL, name: "Big deal", ...owned(EAST) },
    [WEST_DEAL]: { opportunityid: WEST_DEAL, name: "West deal", ...owned(WEST) },
  },
  incidents: { [CASE]: { incidentid: CASE, title: "Printer", ...owned(EAST) } },
  tasks: {
    [TASK]: {
      activityid: TASK,
      _regardingobjectid_value: CONTOSO,
      [`_regardingobjectid_value${L}`]: "account",
      ...owned(WEST),
    },
    [WEST_TASK]: {
      activityid: WEST_TASK,
      _regardingobjectid_value: FABRIKAM,
      [`_regardingobjectid_value${L}`]: "account",
      ...owned(EAST),
    },
  },
  annotations: {
    [NOTE]: { annotationid: NOTE, _objectid_value: CONTOSO, [`_objectid_value${L}`]: "account" },
  },
  teams: {
    [WEST_TEAM]: { teamid: WEST_TEAM, _businessunitid_value: WEST },
    [EAST_TEAM]: { teamid: EAST_TEAM, _businessunitid_value: EAST },
  },
  new_projects: {
    [PROJECT]: {
      new_projectid: PROJECT,
      new_name: "Apollo",
      _owningbusinessunit_value: EAST,
      _owningteam_value: null,
    },
  },
};
const DEFINITIONS: Record<string, object> = {
  new_project: {
    EntitySetName: "new_projects",
    PrimaryIdAttribute: "new_projectid",
    PrimaryNameAttribute: "new_name",
    OwnershipType: "UserOwned",
  },
  product: {
    EntitySetName: "products",
    PrimaryIdAttribute: "productid",
    PrimaryNameAttribute: "name",
    OwnershipType: "OrganizationOwned",
  },
};

type Options = { status?: Record<string, VendorResponse> };

function setup(limits: Record<string, unknown> = {}, options: Options = {}) {
  const requests: VendorRequest[] = [];
  const records = structuredClone(RECORDS);
  const audit: { tool: string; target: string | null; decision: string; reason: string | null }[] =
    [];
  const call = async (req: VendorRequest): Promise<VendorResponse> => {
    requests.push(req);
    const [path, query = ""] = req.path.split("?");
    const forced = options.status?.[`${req.method} ${path}`];
    if (forced) return forced;
    let m;
    if (path === `${API}/WhoAmI`)
      return { status: 200, data: { UserId: guid(99), BusinessUnitId: EAST } };
    if ((m = /^\/api\/data\/v9\.2\/EntityDefinitions\(LogicalName='(\w+)'\)$/.exec(path)))
      return DEFINITIONS[m[1]]
        ? { status: 200, data: DEFINITIONS[m[1]] }
        : { status: 404, data: { error: { code: "0x80060888", message: "Not found" } } };
    if ((m = /^\/api\/data\/v9\.2\/(\w+)\(([0-9a-f-]+)\)$/.exec(path))) {
      const row = records[m[1]]?.[m[2]];
      if (!row)
        return { status: 404, data: { error: { code: "0x80040217", message: "Does Not Exist" } } };
      return req.method === "GET" ? { status: 200, data: row } : { status: 204, data: null };
    }
    if ((m = /^\/api\/data\/v9\.2\/(\w+)$/.exec(path))) {
      if (req.method === "POST") {
        const id = guid(500 + requests.length);
        const key = m[1] === "annotations" ? "annotationid" : `${m[1].replace(/s$/, "")}id`;
        const idColumn = ["tasks", "phonecalls"].includes(m[1]) ? "activityid" : key;
        const row = { [idColumn]: id, ...(req.body as object) };
        (records[m[1]] ??= {})[id] = row;
        return { status: 201, data: row };
      }
      // Every row, whatever the filter: the tools must drop rows outside the key's scope.
      const rows = Object.values(records[m[1]] ?? {});
      const params = new URLSearchParams(query);
      return {
        status: 200,
        data: {
          value: rows,
          ...(params.get("$skiptoken")
            ? {}
            : {
                "@odata.nextLink": `https://contoso.crm.dynamics.com${path}?${query}&$skiptoken=%3Ccookie%20pagenumber%3D%222%22%20%2F%3E`,
              }),
        },
      };
    }
    if (/^\/api\/data\/v9\.2\/(WinOpportunity|LoseOpportunity|CloseIncident)$/.test(path))
      return { status: 204, data: null };
    return { status: 500, data: {} };
  };
  const ctx = {
    limits,
    config: { environment: "https://contoso.crm.dynamics.com" },
    call,
    forbidden: { dynamics_get_customer: "The agents' security role needs prvReadAccount." },
    async audit(e) {
      audit.push(e);
    },
  } as ToolContext;
  const tools = new Map<string, (args: unknown) => Promise<any>>();
  const server = {
    registerTool: (name: string, def: any, fn: (args: unknown) => Promise<any>) => {
      // Parsed as the MCP server would, so defaults apply.
      tools.set(name, (args: unknown) => fn(def.inputSchema.parse(args)));
    },
  };
  dynamicsTools.register(server as never, ctx, new Set(dynamicsTools.names));
  const run = async (name: string, args: object) => {
    const r = await tools.get(name)!(args);
    const text = r.content[0].text as string;
    return { isError: Boolean(r.isError), text, body: r.isError ? null : JSON.parse(text) };
  };
  const writes = () => requests.filter((r) => r.method !== "GET");
  const filterOf = (req: VendorRequest) =>
    new URLSearchParams(req.path.split("?")[1]).get("$filter");
  return { run, requests, audit, writes, filterOf };
}

const east = { business_units: [EAST] };
const OUTSIDE = "This agent may not use that record";

test("every connector capability's tools are Dynamics tools, and every Dynamics tool is in one", () => {
  const dynamics = loadConnectors().find((c) => c.id === "dynamics-365")!;
  assert.deepEqual(
    dynamics.capabilities.flatMap((c) => c.tools).sort(),
    [...dynamicsTools.names].sort(),
  );
  assert.equal(dynamics.capabilities.find((c) => c.id === "delete")!.destructive, true);
});

test("dynamics_search_customers builds a structured query and trims the answer", async () => {
  const { run, requests, filterOf } = setup();
  const r = await run("dynamics_search_customers", {
    table: "account",
    text: "O'Brien",
    filters: [
      { column: "address1_city", op: "eq", value: "Paris" },
      { column: "_primarycontactid_value", op: "eq", value: ADA },
      { column: "modifiedon", op: "ge", value: "2026-10-01" },
      { column: "statecode", op: "eq", value: 0 },
      { column: "websiteurl", op: "null" },
    ],
    top: 5,
  });
  assert.equal(r.isError, false);
  const req = requests[0];
  assert.equal(req.path.split("?")[0], `${API}/accounts`);
  assert.equal(
    filterOf(req),
    `(contains(name,'O''Brien') or contains(accountnumber,'O''Brien')) and address1_city eq 'Paris' and _primarycontactid_value eq ${ADA} and modifiedon ge 2026-10-01 and statecode eq 0 and websiteurl eq null`,
  );
  assert.match(req.headers!.Prefer, /odata\.maxpagesize=5/);
  assert.match(req.headers!.Prefer, /odata\.include-annotations="\*"/);
  const contoso = r.body.records[0];
  assert.equal(contoso.id, CONTOSO);
  assert.equal(contoso.statecode, "Active");
  assert.equal(contoso.revenue, "$1,000,000.00");
  assert.deepEqual(contoso.primarycontactid, { id: ADA, name: "Ada Lovelace", table: "contact" });
  assert.equal(contoso.parentaccountid, null);
  assert.equal(contoso.description.length, 2001);
  assert.ok(!r.text.includes("@odata") && !r.text.includes("owningbusinessunit"));
  assert.equal(r.body.more, true);
  assert.equal(r.body.next, '<cookie pagenumber="2" />');

  const again = await run("dynamics_search_customers", { table: "account", cursor: r.body.next });
  assert.equal(again.body.more, false);
  assert.equal(
    new URLSearchParams(requests[1].path.split("?")[1]).get("$skiptoken"),
    '<cookie pagenumber="2" />',
  );
});

test("filters refuse anything but plain columns and typed values, before any call", async () => {
  const { run, requests } = setup();
  for (const filters of [
    [{ column: "name) or (1 eq 1", op: "eq", value: "x" }],
    [{ column: "Name", op: "eq", value: "x" }],
    [{ column: "name", op: "contains", value: 3 }],
    [{ column: "name", op: "eq" }],
  ]) {
    const r = await run("dynamics_search_customers", { table: "contact", filters });
    assert.equal(r.isError, true, JSON.stringify(filters));
  }
  assert.equal(requests.length, 0);
  assert.equal(literal("it's"), "'it''s'");
  assert.equal(literal("2026-10-27T09:00:00Z"), "2026-10-27T09:00:00Z");
  assert.equal(literal("2026-10-27 or 1 eq 1"), "'2026-10-27 or 1 eq 1'");
  assert.equal(literal(true), "true");
});

test("a business unit or team limit filters every search and drops rows outside it", async () => {
  const units = setup(east);
  const r = await units.run("dynamics_search_customers", { table: "account" });
  assert.deepEqual(
    r.body.records.map((x: { id: string }) => x.id),
    [CONTOSO],
  );
  assert.match(
    units.filterOf(units.requests[0])!,
    new RegExp(`_owningbusinessunit_value eq ${EAST}`),
  );
  assert.match(units.requests[0].path, /_owningteam_value/);

  const teams = setup({ teams: [WEST_TEAM.toUpperCase()] });
  const t = await teams.run("dynamics_search_customers", { table: "account", text: "a" });
  assert.deepEqual(
    t.body.records.map((x: { id: string }) => x.id),
    [FABRIKAM],
  );
  assert.equal(
    teams.filterOf(teams.requests[0]),
    `(contains(name,'a') or contains(accountnumber,'a')) and (_owningteam_value eq ${WEST_TEAM})`,
  );
});

test("the read tools read one record each, and refuse one outside the key's scope", async () => {
  const { run, audit, requests } = setup(east);
  const cases: [string, object, string][] = [
    ["dynamics_get_customer", { table: "account", id: CONTOSO }, "accounts"],
    ["dynamics_get_customer", { table: "contact", id: ADA }, "contacts"],
    ["dynamics_get_opportunity", { id: DEAL }, "opportunities"],
    ["dynamics_get_case", { id: CASE }, "incidents"],
  ];
  for (const [tool, args, set] of cases) {
    const r = await run(tool, args);
    assert.equal(r.isError, false, tool);
    assert.match(requests.at(-1)!.path, new RegExp(`^${API.replaceAll(".", "\\.")}/${set}\\(`));
    assert.match(requests.at(-1)!.path, /\$select=.*_owningbusinessunit_value/);
  }
  const contoso = (await run("dynamics_get_customer", { table: "account", id: CONTOSO })).body;
  assert.equal(contoso.revenue, "$1,000,000.00");

  // Outside the scope and not found at all get the same refusal.
  for (const id of [FABRIKAM, guid(12345)]) {
    const r = await run("dynamics_get_customer", { table: "account", id });
    assert.equal(r.text, OUTSIDE);
    assert.equal(audit.at(-1)!.decision, "denied");
  }
  assert.equal((await run("dynamics_get_opportunity", { id: WEST_DEAL })).text, OUTSIDE);
  assert.equal((await run("dynamics_get_case", { id: "not-a-guid" })).isError, true);
});

test("opportunity and case searches use their own tables", async () => {
  const { run, requests } = setup();
  await run("dynamics_search_opportunities", { text: "deal" });
  assert.match(requests[0].path, /^\/api\/data\/v9\.2\/opportunities\?\$select=opportunityid,name/);
  await run("dynamics_search_cases", { filters: [{ column: "prioritycode", op: "eq", value: 1 }] });
  assert.match(
    requests[1].path,
    /^\/api\/data\/v9\.2\/incidents\?\$select=incidentid,ticketnumber/,
  );
  assert.match(decodeURIComponent(requests[1].path), /prioritycode eq 1/);
});

test("Dynamics' own errors reach the agent with the capability's hint on a 403", async () => {
  const { run, audit } = setup(
    {},
    {
      status: {
        [`GET ${API}/accounts(${CONTOSO})`]: {
          status: 403,
          data: {
            error: { code: "0x80040220", message: "Principal user is missing prvReadAccount" },
          },
        },
        [`GET ${API}/contacts`]: {
          status: 429,
          data: { error: { message: "Number of requests exceeded the limit" } },
        },
      },
    },
  );
  const r = await run("dynamics_get_customer", { table: "account", id: CONTOSO });
  assert.equal(
    r.text,
    "Dynamics 365 answered 403: Principal user is missing prvReadAccount. The agents' security role needs prvReadAccount.",
  );
  assert.equal(audit.at(-1)!.decision, "failed");
  const busy = await run("dynamics_search_customers", { table: "contact" });
  assert.equal(busy.text, "Dynamics 365 answered 429: Number of requests exceeded the limit");
  const gone = await run("dynamics_get_customer", { table: "account", id: guid(12345) });
  assert.equal(gone.text, "Dynamics 365 answered 404: Does Not Exist");
});

test("dynamics_create_customer creates as the app, or for a team inside the key's scope", async () => {
  const open = setup();
  const r = await open.run("dynamics_create_customer", {
    table: "contact",
    fields: { lastname: "Lovelace", firstname: "Ada" },
    account_id: CONTOSO,
  });
  assert.equal(r.isError, false);
  const post = open.writes()[0];
  assert.match(post.path, /^\/api\/data\/v9\.2\/contacts\?\$select=contactid,fullname/);
  assert.deepEqual(post.body, {
    lastname: "Lovelace",
    firstname: "Ada",
    "parentcustomerid_account@odata.bind": `/accounts(${CONTOSO})`,
  });
  assert.match(post.headers!.Prefer, /return=representation/);

  // The app's own business unit is in the scope: no owner is named.
  const units = setup(east);
  await units.run("dynamics_create_customer", { table: "account", fields: { name: "New" } });
  assert.deepEqual(units.writes()[0].body, { name: "New" });

  // One team: records go to it. Several: the agent names one.
  const one = setup({ teams: [WEST_TEAM] });
  await one.run("dynamics_create_customer", { table: "account", fields: { name: "New" } });
  assert.equal((one.writes()[0].body as any)["ownerid@odata.bind"], `/teams(${WEST_TEAM})`);
  const two = setup({ teams: [WEST_TEAM, EAST_TEAM] });
  const refused = await two.run("dynamics_create_customer", {
    table: "account",
    fields: { name: "New" },
  });
  assert.equal(refused.isError, true);
  assert.match(refused.text, /owner_team/);
  await two.run("dynamics_create_customer", {
    table: "account",
    fields: { name: "New" },
    owner_team: EAST_TEAM,
  });
  assert.equal((two.writes()[0].body as any)["ownerid@odata.bind"], `/teams(${EAST_TEAM})`);

  // A team outside the teams, unless its business unit is one of the key's.
  const west = await units.run("dynamics_create_customer", {
    table: "account",
    fields: { name: "New" },
    owner_team: WEST_TEAM,
  });
  assert.equal(west.text, "This agent may not give records to that team");
  const ok = await units.run("dynamics_create_customer", {
    table: "account",
    fields: { name: "New" },
    owner_team: EAST_TEAM,
  });
  assert.equal(ok.isError, false);
  assert.equal(units.writes().length, 2);
});

test("dynamics_create_customer refuses lookups, a contact's account outside the scope, and bad columns", async () => {
  const { run, writes } = setup(east);
  const attempts = [
    { table: "contact", fields: { lastname: "X" }, account_id: FABRIKAM },
    { table: "account", fields: { name: "X" }, account_id: CONTOSO },
    { table: "account", fields: { "ownerid@odata.bind": `/teams(${WEST_TEAM})` } },
    { table: "account", fields: { _ownerid_value: WEST_TEAM } },
    { table: "account", fields: {} },
  ];
  for (const args of attempts) {
    const r = await run("dynamics_create_customer", args);
    assert.equal(r.isError, true, JSON.stringify(args));
  }
  assert.equal(writes().length, 0);
});

test("dynamics_update_customer patches only an existing record inside the scope", async () => {
  const { run, writes } = setup(east);
  const r = await run("dynamics_update_customer", {
    table: "account",
    id: CONTOSO,
    fields: { telephone1: "+1 555 0100" },
  });
  assert.deepEqual(r.body, { updated: true, id: CONTOSO, columns: ["telephone1"] });
  assert.deepEqual(writes()[0], {
    method: "PATCH",
    path: `${API}/accounts(${CONTOSO})`,
    body: { telephone1: "+1 555 0100" },
    headers: { Prefer: 'odata.include-annotations="*"', "If-Match": "*" },
  });
  const outside = await run("dynamics_update_customer", {
    table: "account",
    id: FABRIKAM,
    fields: { telephone1: "x" },
  });
  assert.equal(outside.text, OUTSIDE);
  assert.equal(writes().length, 1);
});

test("dynamics_close_opportunity wins or loses an opportunity inside the scope", async () => {
  const { run, writes } = setup(east);
  const won = await run("dynamics_close_opportunity", {
    id: DEAL,
    outcome: "won",
    subject: "Signed",
    actual_revenue: 5000,
  });
  assert.deepEqual(won.body, { closed: true, id: DEAL, outcome: "won" });
  assert.deepEqual(writes()[0].path, `${API}/WinOpportunity`);
  assert.deepEqual(writes()[0].body, {
    Status: 3,
    OpportunityClose: {
      subject: "Signed",
      actualrevenue: 5000,
      "opportunityid@odata.bind": `/opportunities(${DEAL})`,
    },
  });
  await run("dynamics_close_opportunity", { id: DEAL, outcome: "lost", subject: "Price" });
  assert.equal(writes()[1].path, `${API}/LoseOpportunity`);
  assert.equal((writes()[1].body as any).Status, 4);
  const refused = [
    { id: DEAL, outcome: "lost", subject: "x", actual_revenue: 1 },
    { id: WEST_DEAL, outcome: "won", subject: "x" },
  ];
  for (const args of refused)
    assert.equal((await run("dynamics_close_opportunity", args)).isError, true);
  assert.equal(writes().length, 2);
});

test("dynamics_close_case resolves with a resolution or cancels", async () => {
  const { run, writes } = setup(east);
  await run("dynamics_close_case", { id: CASE, outcome: "resolved", resolution: "Replaced toner" });
  assert.deepEqual(writes()[0].body, {
    Status: 5,
    IncidentResolution: {
      subject: "Replaced toner",
      "incidentid@odata.bind": `/incidents(${CASE})`,
    },
  });
  await run("dynamics_close_case", { id: CASE, outcome: "cancelled" });
  assert.equal(writes()[1].method, "PATCH");
  assert.deepEqual(writes()[1].body, { statecode: 2, statuscode: 6 });
  assert.equal((await run("dynamics_close_case", { id: CASE, outcome: "resolved" })).isError, true);
  const outside = await run("dynamics_close_case", { id: guid(41), outcome: "cancelled" });
  assert.equal(outside.text, OUTSIDE);
  assert.equal(writes().length, 2);
});

test("dynamics_list_activities lists a record's activities or notes once the record is in scope", async () => {
  const { run, requests, filterOf } = setup(east);
  const r = await run("dynamics_list_activities", { table: "account", id: CONTOSO });
  assert.equal(r.isError, false);
  const list = requests.at(-1)!;
  assert.match(list.path, /^\/api\/data\/v9\.2\/activitypointers\?/);
  // The record's activities, whoever owns them: no owner filter here.
  assert.equal(filterOf(list), `_regardingobjectid_value eq ${CONTOSO}`);
  assert.equal(r.body.records.length, 0);
  await run("dynamics_list_activities", { table: "account", id: CONTOSO, kind: "notes" });
  assert.equal(filterOf(requests.at(-1)!), `_objectid_value eq ${CONTOSO}`);
  const before = requests.length;
  const outside = await run("dynamics_list_activities", { table: "account", id: FABRIKAM });
  assert.equal(outside.text, OUTSIDE);
  assert.equal(requests.length, before + 1);
});

test("dynamics_add_activity adds a note, a task or a completed phone call", async () => {
  const { run, writes } = setup(east);
  const note = await run("dynamics_add_activity", {
    table: "opportunity",
    id: DEAL,
    kind: "note",
    subject: "Call notes",
    text: "Wants a demo",
  });
  assert.equal(note.body.kind, "note");
  assert.deepEqual(writes()[0].body, {
    subject: "Call notes",
    notetext: "Wants a demo",
    "objectid_opportunity@odata.bind": `/opportunities(${DEAL})`,
  });
  await run("dynamics_add_activity", {
    table: "account",
    id: CONTOSO,
    kind: "task",
    subject: "Send quote",
    due: "2026-11-03T17:00:00Z",
  });
  assert.deepEqual(writes()[1].body, {
    subject: "Send quote",
    scheduledend: "2026-11-03T17:00:00Z",
    "regardingobjectid_account_task@odata.bind": `/accounts(${CONTOSO})`,
  });
  const call = await run("dynamics_add_activity", {
    table: "incident",
    id: CASE,
    kind: "phone_call",
    subject: "Called back",
    phone_number: "+1 555 0100",
    direction: "incoming",
  });
  assert.deepEqual(writes()[2].body, {
    subject: "Called back",
    directioncode: false,
    phonenumber: "+1 555 0100",
    "regardingobjectid_incident_phonecall@odata.bind": `/incidents(${CASE})`,
  });
  assert.deepEqual(writes()[3], {
    method: "PATCH",
    path: `${API}/phonecalls(${call.body.id})`,
    body: { statecode: 1, statuscode: 2 },
    headers: { Prefer: 'odata.include-annotations="*"', "If-Match": "*" },
  });
  await run("dynamics_add_activity", {
    table: "contact",
    id: ADA,
    kind: "phone_call",
    subject: "To call",
    completed: false,
  });
  assert.equal(writes().length, 5);
  const refused = [
    { table: "account", id: FABRIKAM, kind: "note", subject: "x" },
    { table: "account", id: CONTOSO, kind: "note", subject: "x", due: "2026-11-03" },
    { table: "account", id: CONTOSO, kind: "task", subject: "x", due: "next week" },
    { table: "account", id: CONTOSO, kind: "task", subject: "x", phone_number: "1" },
  ];
  for (const args of refused)
    assert.equal((await run("dynamics_add_activity", args)).isError, true, JSON.stringify(args));
  assert.equal(writes().length, 5);
});

test("dynamics_list_rows and dynamics_get_row read only the tables listed on the key", async () => {
  const none = setup();
  const unlisted = await none.run("dynamics_list_rows", { table: "new_project" });
  assert.match(unlisted.text, /no tables to read/);
  assert.equal(none.requests.length, 0);

  const { run, requests, filterOf } = setup({ tables: ["new_project", "product"] });
  const other = await run("dynamics_list_rows", { table: "systemuser" });
  assert.match(other.text, /may not read systemuser/);
  assert.equal(requests.length, 0);
  const r = await run("dynamics_list_rows", {
    table: "new_project",
    columns: ["new_name", "_new_customer_value"],
    filters: [{ column: "new_name", op: "startswith", value: "Ap" }],
    order_by: "new_name",
    descending: true,
    top: 10,
  });
  assert.deepEqual(r.body.records[0], { id: PROJECT, new_name: "Apollo" });
  const list = requests.at(-1)!;
  const q = new URLSearchParams(list.path.split("?")[1]);
  assert.equal(q.get("$select"), "new_projectid,new_name,_new_customer_value");
  assert.equal(q.get("$orderby"), "new_name desc");
  assert.equal(filterOf(list), "startswith(new_name,'Ap')");
  const products = await run("dynamics_list_rows", { table: "product" });
  assert.equal(products.isError, false);
  assert.equal(
    new URLSearchParams(requests.at(-1)!.path.split("?")[1]).get("$select"),
    "productid,name",
  );
  const row = await run("dynamics_get_row", { table: "new_project", id: PROJECT });
  assert.deepEqual(row.body, { id: PROJECT, new_name: "Apollo" });
  assert.equal(
    (await run("dynamics_list_rows", { table: "new_project", order_by: "x desc" })).isError,
    true,
  );
});

test("under a scope, other tables are filtered by owner, and tables without owners are refused", async () => {
  const { run, requests, filterOf } = setup({ ...east, tables: ["new_project", "product"] });
  const r = await run("dynamics_list_rows", { table: "new_project" });
  assert.equal(r.body.records.length, 1);
  assert.match(filterOf(requests.at(-1)!)!, new RegExp(`_owningbusinessunit_value eq ${EAST}`));
  const products = await run("dynamics_list_rows", { table: "product" });
  assert.match(products.text, /have no owner/);
  assert.equal((await run("dynamics_get_row", { table: "product", id: PROJECT })).isError, true);
  assert.ok(!requests.some((q) => q.path.startsWith(`${API}/products`)));
});

test("dynamics_delete_record deletes, and under a scope only activities and notes of its records", async () => {
  const open = setup();
  const r = await open.run("dynamics_delete_record", { table: "account", id: FABRIKAM });
  assert.deepEqual(r.body, { deleted: true, table: "account", id: FABRIKAM });
  assert.deepEqual(open.writes()[0], {
    method: "DELETE",
    path: `${API}/accounts(${FABRIKAM})`,
    headers: { Prefer: 'odata.include-annotations="*"' },
  });

  const { run, writes, audit } = setup(east);
  const customer = await run("dynamics_delete_record", { table: "account", id: CONTOSO });
  assert.match(customer.text, /cannot delete account records/);
  assert.equal(audit.at(-1)!.decision, "denied");
  // TASK is owned in WEST but belongs to Contoso (EAST): its record decides.
  assert.equal((await run("dynamics_delete_record", { table: "task", id: TASK })).isError, false);
  assert.equal(
    (await run("dynamics_delete_record", { table: "annotation", id: NOTE })).isError,
    false,
  );
  for (const id of [WEST_TASK, guid(52)])
    assert.equal((await run("dynamics_delete_record", { table: "task", id })).text, OUTSIDE);
  assert.deepEqual(
    writes().map((w) => w.path),
    [`${API}/tasks(${TASK})`, `${API}/annotations(${NOTE})`],
  );
});

test("dynamics_delete_record is the only destructive tool", () => {
  const destructive: string[] = [];
  dynamicsTools.register(
    {
      registerTool: (name: string, def: { annotations?: { destructiveHint?: boolean } }) => {
        if (def.annotations?.destructiveHint) destructive.push(name);
      },
    } as never,
    { limits: {} } as ToolContext,
    new Set(dynamicsTools.names),
  );
  assert.deepEqual(destructive, ["dynamics_delete_record"]);
});
