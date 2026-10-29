import assert from "node:assert/strict";
import { test } from "node:test";
import type { z } from "zod";
import type { ToolContext } from "../../packages/connector-kit/src/common.ts";
import { loadConnectors } from "../../packages/connector-kit/src/connector.ts";
import type { VendorRequest, VendorResponse } from "../../packages/connector-kit/src/types.ts";
import { htmlText, servicenowTools } from "./tools.ts";
import { normalizeInstance, servicenowVendor } from "./vendor.ts";

// The ServiceNow tools against a fake instance. Its Table API answers with
// sysparm_display_value=all records and applies the simple `field=value` conditions of the
// query (OR groups included); anything else (LIKE, ORDERBY, text search) is ignored.
// Service Desk and Acme are in the usual limits; Network and Globex are not.

const id = (n: number) => n.toString(16).padStart(32, "0");
const f = (value: string, display = value) => ({ value, display_value: display });

const incident = (n: number, group: string, company: string) => ({
  sys_id: f(id(n)),
  number: f(`INC000000${n}`),
  short_description: f(`Printer ${n}`),
  state: f("2", "In Progress"),
  priority: f("3", "3 - Moderate"),
  assignment_group: f(id(100 + n), group),
  assigned_to: f("", ""),
  caller_id: f(id(200 + n), `Caller ${n}`),
  company: f(id(300 + n), company),
  opened_at: f("2026-10-01 09:00:00", "01/10/2026 11:00:00"),
  sys_updated_on: f("2026-10-02 09:00:00"),
  description: f("It jams"),
  work_notes: f("", "2026-10-02 - Agent (Work notes)\nChecked the tray"),
  comments: f("", ""),
});

const TABLES: Record<string, Record<string, unknown>[]> = {
  incident: [incident(1, "Service Desk", "Acme"), incident(2, "Network", "Globex")],
  change_request: [
    {
      sys_id: f(id(11)),
      number: f("CHG0030001"),
      assignment_group: f(id(111), "Service Desk"),
      company: f(id(311), "Acme"),
      type: f("normal", "Normal"),
      backout_plan: f("Roll back"),
    },
    {
      sys_id: f(id(12)),
      number: f("CHG0030002"),
      assignment_group: f(id(112), "Network"),
      company: f(id(312), "Acme"),
    },
  ],
  cmdb_ci: [
    {
      sys_id: f(id(21)),
      name: f("LAPTOP-1"),
      serial_number: f("SN1"),
      company: f(id(321), "Acme"),
      sys_class_name: f("cmdb_ci_computer", "Computer"),
    },
  ],
  kb_knowledge: [
    {
      sys_id: f(id(31)),
      number: f("KB0010001"),
      short_description: f("VPN"),
      workflow_state: f("published"),
      text: f("<p>Restart&nbsp;the <b>VPN</b></p><script>x()</script><p>Then sign in</p>"),
    },
    {
      sys_id: f(id(32)),
      number: f("KB0010002"),
      short_description: f("Draft"),
      workflow_state: f("draft"),
      text: f("<p>Not yet</p>"),
    },
  ],
  sys_user: [
    {
      sys_id: f(id(41)),
      name: f("Ada"),
      email: f("ada@acme.example"),
      user_name: f("ada"),
      active: f("true"),
      company: f(id(341), "Acme"),
    },
    {
      sys_id: f(id(42)),
      name: f("Bob"),
      email: f("bob@globex.example"),
      user_name: f("bob"),
      active: f("true"),
      company: f(id(342), "Globex"),
    },
  ],
  sys_user_group: [
    { sys_id: f(id(51)), name: f("Service Desk"), active: f("true") },
    { sys_id: f(id(52)), name: f("Network"), active: f("true") },
  ],
  sc_req_item: [
    {
      sys_id: f(id(61)),
      number: f("RITM0010001"),
      request: f(id(62), "REQ0010001"),
      "requested_for.company": f(id(341), "Acme"),
      stage: f("fulfillment", "Fulfillment"),
    },
  ],
};

// A record's value for a query field; `x.name` (or any `x.<field>`) reads x's display value.
function valueOf(r: Record<string, any>, field: string): string | undefined {
  if (r[field]) return r[field].display_value;
  const ref = field.replace(/\.[^.]+$/, "");
  return ref === field ? undefined : r[ref]?.display_value;
}

function matches(r: Record<string, any>, query: string): boolean {
  const groups: string[][] = [];
  for (const part of query.split("^")) {
    if (part.startsWith("OR")) groups.at(-1)!.push(part.slice(2));
    else groups.push([part]);
  }
  return groups.every((alts) =>
    alts.some((c) => {
      const m = /^([a-z_.]+)=(.*)$/.exec(c);
      if (!m) return true;
      const v = valueOf(r, m[1]);
      // A field the fake does not keep (active, say) matches.
      if (v === undefined) return !m[1].includes(".");
      return v.toLowerCase() === m[2].toLowerCase() || r[m[1]]?.value === m[2];
    }),
  );
}

type Options = {
  answer?: VendorResponse;
  ignoreQuery?: boolean;
  forbidden?: Record<string, string>;
};

function setup(limits: Record<string, unknown> = {}, options: Options = {}) {
  const sent: VendorRequest[] = [];
  const audit: Parameters<ToolContext["audit"]>[0][] = [];
  const call = async (req: VendorRequest): Promise<VendorResponse> => {
    sent.push(req);
    if (options.answer) return options.answer;
    const [path, qs = ""] = req.path.split("?");
    const params = new URLSearchParams(qs);
    const table = /^\/api\/now\/table\/([a-z_]+)(?:\/([0-9a-f]{32}))?$/.exec(path);
    if (table && req.method === "GET") {
      const rows = (TABLES[table[1]] ?? []).filter(
        (r) => options.ignoreQuery || matches(r, params.get("sysparm_query") ?? ""),
      );
      const start = Number(params.get("sysparm_offset") ?? 0);
      return {
        status: 200,
        data: { result: rows.slice(start, start + Number(params.get("sysparm_limit"))) },
      };
    }
    if (table && req.method === "POST")
      return {
        status: 201,
        data: { result: { ...incident(9, "Service Desk", "Acme"), number: f("INC0000009") } },
      };
    if (table && req.method === "PATCH") {
      const r = TABLES[table[1]].find((x: any) => x.sys_id.value === table[2]);
      return { status: 200, data: { result: { ...r, state: f("6", "Resolved") } } };
    }
    if (path === "/api/sn_sc/servicecatalog/items")
      return {
        status: 200,
        data: {
          result: [1, 2, 3].map((n) => ({
            sys_id: id(70 + n),
            name: `Laptop ${n}`,
            category: { title: "Hardware" },
            price: "$1,000.00",
            extra: "dropped",
          })),
        },
      };
    if (path === `/api/sn_sc/servicecatalog/items/${id(71)}`)
      return {
        status: 200,
        data: {
          result: {
            sys_id: id(71),
            name: "Laptop",
            description: "<p>A <b>laptop</b></p>",
            variables: [
              {
                name: "model",
                label: "Model",
                friendly_type: "select_box",
                mandatory: true,
                choices: [{ value: "pro", label: "Pro", index: 0 }],
              },
              { label: "Section", type: 19 },
            ],
          },
        },
      };
    if (path === `/api/sn_sc/servicecatalog/items/${id(71)}/order_now`)
      return { status: 200, data: { result: { request_number: "REQ0010009", sys_id: id(79) } } };
    return { status: 404, data: { error: { message: "No Record found" } } };
  };
  const ctx: ToolContext = {
    limits: limits as ToolContext["limits"],
    config: {},
    call,
    forbidden: options.forbidden,
    audit: async (e) => {
      audit.push(e);
    },
  };
  const tools = new Map<string, { schema: z.ZodType; run: (a: unknown) => Promise<any> }>();
  const server = {
    registerTool: (
      name: string,
      def: { inputSchema: z.ZodType },
      run: (a: unknown) => Promise<any>,
    ) => tools.set(name, { schema: def.inputSchema, run }),
  };
  servicenowTools.register(server as never, ctx, new Set(servicenowTools.names));
  const run = async (name: string, args: unknown) => {
    const tool = tools.get(name)!;
    const r = await tool.run(tool.schema.parse(args));
    const text = r.content[0].text as string;
    return { isError: Boolean(r.isError), text, body: r.isError ? null : JSON.parse(text) };
  };
  const queries = () =>
    sent.map((s) => new URLSearchParams(s.path.split("?")[1]).get("sysparm_query") ?? "");
  return { run, sent, audit, queries };
}

const DESK = { groups: ["Service Desk"] };
const ACME = { companies: ["Acme"] };

test("every connector tool is registered and in exactly one capability", () => {
  const connector = loadConnectors().find((c) => c.id === "servicenow")!;
  assert.deepEqual(
    connector.capabilities.flatMap((c) => c.tools).sort(),
    [...servicenowTools.names].sort(),
  );
  // Nothing deletes, so no capability asks to confirm.
  assert.ok(connector.capabilities.every((c) => !c.destructive));
});

test("search_incidents answers trimmed incidents and pages with a cursor", async () => {
  const { run, sent } = setup();
  const r = await run("servicenow_search_incidents", { max_results: 1 });
  assert.deepEqual(r.body.incidents, [
    {
      number: "INC0000001",
      short_description: "Printer 1",
      state: "In Progress",
      priority: "3 - Moderate",
      assignment_group: "Service Desk",
      assigned_to: null,
      caller: "Caller 1",
      company: "Acme",
      opened: "2026-10-01T09:00:00Z",
      updated: "2026-10-02T09:00:00Z",
    },
  ]);
  assert.equal(r.body.next, "1");
  const params = new URLSearchParams(sent[0].path.split("?")[1]);
  assert.equal(params.get("sysparm_limit"), "2");
  assert.equal(params.get("sysparm_display_value"), "all");
  assert.ok(params.get("sysparm_fields")!.split(",").includes("short_description"));
  const second = await run("servicenow_search_incidents", { max_results: 1, cursor: "1" });
  assert.equal(second.body.incidents[0].number, "INC0000002");
  assert.equal(second.body.next, null);
  assert.equal((await run("servicenow_search_incidents", { cursor: "1^x" })).isError, true);
});

test("search_incidents builds its query from parameters and the key's limits", async () => {
  const { run, queries } = setup({ ...DESK, ...ACME });
  const r = await run("servicenow_search_incidents", {
    text: "jam",
    priority: 2,
    caller: "ada",
    assignment_group: "service desk",
  });
  assert.equal(r.isError, false);
  assert.equal(
    queries()[0],
    "active=true^short_descriptionLIKEjam^ORdescriptionLIKEjam^priority=2" +
      "^assignment_group.name=service desk^caller_id.email=ada^ORcaller_id.user_name=ada" +
      "^assignment_group.name=Service Desk^company.name=Acme^ORDERBYDESCsys_updated_on",
  );
});

test("search_incidents keeps only incidents within the limits, whatever ServiceNow answers", async () => {
  const { run } = setup(DESK, { ignoreQuery: true });
  const r = await run("servicenow_search_incidents", {});
  assert.deepEqual(
    r.body.incidents.map((i: { number: string }) => i.number),
    ["INC0000001"],
  );
  const companies = setup(ACME, { ignoreQuery: true });
  const c = await companies.run("servicenow_search_incidents", {});
  assert.deepEqual(
    c.body.incidents.map((i: { company: string }) => i.company),
    ["Acme"],
  );
});

test("a value that could change the query is refused before ServiceNow is called", async () => {
  const { run, sent, audit } = setup(DESK);
  for (const args of [
    { text: "x^NQactive=true" },
    { text: "x\nactive=true" },
    { caller: "javascript:gs.getUserID()" },
    { assigned_to: "a^ORsys_id!=0" },
  ]) {
    const r = await run("servicenow_search_incidents", args);
    assert.equal(r.isError, true, JSON.stringify(args));
  }
  assert.equal(sent.length, 0);
  assert.ok(audit.every((a) => a.decision === "denied"));
  const groupOut = await run("servicenow_search_incidents", { assignment_group: "Network" });
  assert.match(groupOut.text, /may not use the assignment group Network/);
  assert.equal(sent.length, 0);
});

test("get_incident reads an incident with its notes; one outside the limits is not found", async () => {
  const { run, sent } = setup(ACME);
  const r = await run("servicenow_get_incident", { number: "INC0000001" });
  assert.equal(r.body.description, "It jams");
  assert.match(r.body.work_notes, /Checked the tray/);
  assert.equal(r.body.comments, null);
  const out = await run("servicenow_get_incident", { number: "INC0000002" });
  assert.equal(out.isError, true);
  assert.match(out.text, /was not found, or this agent may not use it/);
  const missing = await run("servicenow_get_incident", { number: "INC0000404" });
  assert.equal(missing.text, "INC0000404 was not found, or this agent may not use it");
  const bad = await run("servicenow_get_incident", { number: "INC1^ORactive=true" });
  assert.match(bad.text, /Give the incident number, such as INC0010001/);
  assert.equal(sent.length, 3);
});

test("create_incident resolves the group and caller and sets the caller's company", async () => {
  const { run, sent } = setup({ ...DESK, ...ACME });
  const r = await run("servicenow_create_incident", {
    short_description: "Printer jam",
    caller: "ada@acme.example",
    assignment_group: "Service Desk",
    impact: 2,
  });
  assert.equal(r.body.number, "INC0000009");
  const post = sent.find((s) => s.method === "POST")!;
  assert.ok(post.path.startsWith("/api/now/table/incident?"));
  assert.deepEqual(post.body, {
    short_description: "Printer jam",
    assignment_group: id(51),
    caller_id: id(41),
    company: id(341),
    impact: "2",
  });
});

test("create_incident under limits needs a group and caller inside them", async () => {
  const groups = setup(DESK);
  assert.match(
    (await groups.run("servicenow_create_incident", { short_description: "x" })).text,
    /Give an assignment_group/,
  );
  assert.match(
    (
      await groups.run("servicenow_create_incident", {
        short_description: "x",
        assignment_group: "Network",
      })
    ).text,
    /may not use the assignment group Network/,
  );
  const companies = setup(ACME);
  assert.match(
    (await companies.run("servicenow_create_incident", { short_description: "x" })).text,
    /Give a caller/,
  );
  const outside = await companies.run("servicenow_create_incident", {
    short_description: "x",
    caller: "bob",
  });
  assert.match(outside.text, /No single active person bob that this agent may act for/);
  for (const s of [groups, companies]) assert.ok(!s.sent.some((x) => x.method === "POST"));
  // Without limits nothing more is needed.
  const open = setup();
  assert.equal(
    (await open.run("servicenow_create_incident", { short_description: "x" })).isError,
    false,
  );
  assert.deepEqual(open.sent.at(-1)!.body, { short_description: "x" });
});

test("update_incident changes fields by sys_id after reading the incident in scope", async () => {
  const { run, sent } = setup(DESK);
  const r = await run("servicenow_update_incident", {
    number: "INC0000001",
    state: "on_hold",
    on_hold_reason: "awaiting_vendor",
    assigned_to: "bob",
    urgency: 1,
  });
  assert.equal(r.isError, false);
  const patch = sent.find((s) => s.method === "PATCH")!;
  assert.ok(patch.path.startsWith(`/api/now/table/incident/${id(1)}?`));
  assert.deepEqual(patch.body, { urgency: "1", state: "3", hold_reason: "4", assigned_to: id(42) });
});

test("update_incident refuses incidents and groups outside the limits, and bad holds", async () => {
  const { run, sent } = setup(DESK);
  assert.match(
    (await run("servicenow_update_incident", { number: "INC0000002", urgency: 1 })).text,
    /not found/,
  );
  assert.match(
    (await run("servicenow_update_incident", { number: "INC0000001", assignment_group: "Network" }))
      .text,
    /may not use the assignment group Network/,
  );
  assert.match(
    (await run("servicenow_update_incident", { number: "INC0000001", state: "on_hold" })).text,
    /on_hold_reason/,
  );
  assert.match(
    (await run("servicenow_update_incident", { number: "INC0000001" })).text,
    /at least one field/,
  );
  assert.ok(!sent.some((s) => s.method === "PATCH"));
});

test("add_incident_note writes a work note unless the customer should see it", async () => {
  const { run, sent, audit } = setup();
  const internal = await run("servicenow_add_incident_note", {
    number: "INC0000001",
    text: "Checked",
  });
  assert.deepEqual(internal.body, { number: "INC0000001", added: "work note" });
  await run("servicenow_add_incident_note", {
    number: "INC0000001",
    text: "Hi",
    visibility: "customer",
  });
  const patches = sent.filter((s) => s.method === "PATCH").map((s) => s.body);
  assert.deepEqual(patches, [{ work_notes: "Checked" }, { comments: "Hi" }]);
  assert.equal(audit.at(-1)!.target, "INC0000001 (customer)");
});

test("resolve_incident sets the resolved state with code and notes", async () => {
  const { run, sent } = setup(ACME);
  const r = await run("servicenow_resolve_incident", {
    number: "INC0000001",
    close_code: "Solution provided",
    close_notes: "Cleared",
  });
  assert.equal(r.body.state, "Resolved");
  assert.deepEqual(sent.at(-1)!.body, {
    state: "6",
    close_code: "Solution provided",
    close_notes: "Cleared",
  });
  const out = await run("servicenow_resolve_incident", {
    number: "INC0000002",
    close_code: "x",
    close_notes: "x",
  });
  assert.equal(out.isError, true);
  assert.equal(sent.filter((s) => s.method === "PATCH").length, 1);
});

test("changes are searched and read within the groups limit", async () => {
  const { run, queries } = setup(DESK);
  const list = await run("servicenow_search_changes", { status: "any" });
  assert.deepEqual(
    list.body.changes.map((c: { number: string }) => c.number),
    ["CHG0030001"],
  );
  assert.equal(queries()[0], "assignment_group.name=Service Desk^ORDERBYDESCsys_updated_on");
  const one = await run("servicenow_get_change", { number: "CHG0030001" });
  assert.equal(one.body.backout_plan, "Roll back");
  assert.equal(one.body.type, "Normal");
  assert.equal((await run("servicenow_get_change", { number: "CHG0030002" })).isError, true);
});

test("find_configuration_items needs a name or serial and keeps to the companies", async () => {
  const { run, queries } = setup(ACME);
  assert.match((await run("servicenow_find_configuration_items", {})).text, /name or a serial/);
  const r = await run("servicenow_find_configuration_items", { serial_number: "SN1" });
  assert.deepEqual(r.body.items[0].class, "Computer");
  assert.equal(r.body.more, false);
  assert.equal(queries()[0], "serial_number=SN1^company.name=Acme^ORDERBYname");
  const other = setup({ companies: ["Globex"] });
  assert.deepEqual(
    (await other.run("servicenow_find_configuration_items", { name: "LAPTOP" })).body.items,
    [],
  );
});

test("knowledge search and articles are published only, as text", async () => {
  const { run, queries } = setup(DESK);
  const list = await run("servicenow_search_knowledge", { text: "vpn" });
  assert.deepEqual(
    list.body.articles.map((a: { number: string }) => a.number),
    ["KB0010001"],
  );
  assert.equal(queries()[0], "workflow_state=published^123TEXTQUERY321=vpn");
  const article = await run("servicenow_get_article", { number: "KB0010001" });
  assert.equal(article.body.text, "Restart the VPN\nThen sign in");
  assert.match(
    (await run("servicenow_get_article", { number: "KB0010002" })).text,
    /not published/,
  );
  assert.equal(htmlText("a &amp;lt; b"), "a &lt; b");
});

test("catalog items are searched, read with their variables, and ordered", async () => {
  const { run, sent } = setup();
  const list = await run("servicenow_search_catalog_items", { text: "laptop", max_results: 2 });
  assert.equal(list.body.items.length, 2);
  assert.deepEqual(list.body.items[0], {
    id: id(71),
    name: "Laptop 1",
    short_description: null,
    category: "Hardware",
    price: "$1,000.00",
  });
  assert.equal(list.body.next, "2");
  const item = await run("servicenow_get_catalog_item", { id: id(71) });
  assert.equal(item.body.description, "A laptop");
  assert.deepEqual(item.body.variables, [
    {
      name: "model",
      label: "Model",
      type: "select_box",
      mandatory: true,
      choices: [{ value: "pro", label: "Pro" }],
    },
  ]);
  const order = await run("servicenow_order_catalog_item", {
    id: id(71),
    variables: { model: "pro" },
    requested_for: "ada",
  });
  assert.deepEqual(order.body, { request_number: "REQ0010009" });
  assert.deepEqual(sent.at(-1)!.body, {
    sysparm_quantity: "1",
    variables: { model: "pro" },
    sysparm_requested_for: id(41),
  });
  assert.equal((await run("servicenow_get_catalog_item", { id: "../items" })).isError, true);
});

test("orders under a companies limit are only for people of those companies", async () => {
  const { run, sent } = setup(ACME);
  assert.match(
    (await run("servicenow_order_catalog_item", { id: id(71) })).text,
    /Give requested_for/,
  );
  assert.match(
    (await run("servicenow_order_catalog_item", { id: id(71), requested_for: "bob" })).text,
    /No single active person bob/,
  );
  assert.ok(!sent.some((s) => s.path.endsWith("order_now")));
  assert.equal(
    (await run("servicenow_order_catalog_item", { id: id(71), requested_for: "ada" })).isError,
    false,
  );
});

test("search_requests finds requested items by request number within the companies", async () => {
  const { run, queries } = setup(ACME);
  const r = await run("servicenow_search_requests", { number: "REQ0010001", status: "any" });
  assert.deepEqual(
    r.body.requests.map((x: { number: string; request: string; stage: string }) => [
      x.number,
      x.request,
      x.stage,
    ]),
    [["RITM0010001", "REQ0010001", "Fulfillment"]],
  );
  assert.equal(
    queries()[0],
    "request.number=REQ0010001^requested_for.company.name=Acme^ORDERBYDESCsys_updated_on",
  );
  const other = setup({ companies: ["Globex"] }, { ignoreQuery: true });
  assert.deepEqual((await other.run("servicenow_search_requests", {})).body.requests, []);
});

test("a ServiceNow refusal reaches the agent with its reason and the capability's hint", async () => {
  const { run, audit } = setup(
    {},
    {
      answer: {
        status: 403,
        data: {
          error: { message: "Operation Failed", detail: "ACL Exception Insufficient rights" },
          status: "failure",
        },
      },
      forbidden: {
        servicenow_search_incidents: "Give the integration user the sn_incident_read role.",
      },
    },
  );
  const r = await run("servicenow_search_incidents", {});
  assert.equal(r.isError, true);
  assert.equal(
    r.text,
    "ServiceNow answered 403: Operation Failed: ACL Exception Insufficient rights. Give the integration user the sn_incident_read role.",
  );
  assert.deepEqual([audit[0].decision, audit[0].vendorStatus], ["failed", 403]);
});

test("the instance is a bare name on service-now.com or any https address as entered", () => {
  assert.equal(normalizeInstance("acme"), "acme.service-now.com");
  assert.equal(
    normalizeInstance("https://Acme.service-now.com/now/nav/ui"),
    "acme.service-now.com",
  );
  assert.equal(normalizeInstance("https://itsm.acme.example/"), "itsm.acme.example");
  assert.equal(normalizeInstance("itsm.acme.example:8443"), "itsm.acme.example:8443");
  // Addresses are not refused for being IP literals, as for the http connector.
  assert.equal(normalizeInstance("https://10.0.0.5"), "10.0.0.5");
  assert.equal(normalizeInstance("https://[fd00::5]"), "[fd00::5]");
  for (const bad of [
    "http://acme.service-now.com",
    "ftp://itsm.acme.example",
    "https://user:pw@itsm.acme.example",
    "https://intranet",
    "a b",
    "",
  ])
    assert.throws(() => normalizeInstance(bad), /instance address/, bad);
});

test("a custom domain gets every call, the token included, and nothing else", async () => {
  const seen: string[] = [];
  const fake = (async (url: string | URL) => {
    seen.push(String(url));
    if (String(url).endsWith("/oauth_token.do"))
      return new Response(JSON.stringify({ access_token: "tok", expires_in: 1799 }));
    return new Response(JSON.stringify({ result: [] }));
  }) as typeof fetch;
  const vendor = servicenowVendor(fake);
  const conn = await vendor.connect({
    instance: "https://itsm.acme.example",
    auth: "oauth",
    client_id: "cid",
    client_secret: "cs",
  });
  assert.deepEqual(
    [conn.config.instance, conn.accountLabel],
    ["itsm.acme.example", "itsm.acme.example"],
  );
  await vendor.call(conn, { method: "GET", path: "/api/now/table/incident" });
  assert.deepEqual(seen, [
    "https://itsm.acme.example/oauth_token.do",
    "https://itsm.acme.example/api/now/table/incident",
  ]);
  for (const path of ["@evil.example/x", "//evil.example/x", "/\\evil.example/x"])
    await assert.rejects(vendor.call(conn, { method: "GET", path }), /Refused ServiceNow path/);
  assert.equal(seen.length, 2);
});

test("a redirect from the instance is refused, never followed", async () => {
  const inits: RequestInit[] = [];
  const fake = (async (_url: string, init: RequestInit) => {
    inits.push(init);
    return new Response(null, { status: 302, headers: { location: "https://evil.example/" } });
  }) as typeof fetch;
  const vendor = servicenowVendor(fake);
  await assert.rejects(
    vendor.connect({ instance: "itsm.acme.example", auth: "basic", username: "u", password: "p" }),
    /redirect \(302\), which is not followed/,
  );
  await assert.rejects(
    vendor.connect({ instance: "acme", auth: "oauth", client_id: "a", client_secret: "b" }),
    /redirect \(302\)/,
  );
  assert.ok(inits.every((i) => i.redirect === "manual"));
});

test("the vendor signs in with OAuth client credentials or basic, only at the instance", async () => {
  const seen: { url: string; init: RequestInit }[] = [];
  const fake = (async (url: string, init: RequestInit) => {
    seen.push({ url, init });
    if (url.endsWith("/oauth_token.do"))
      return new Response(JSON.stringify({ access_token: "tok", expires_in: 1799 }), {
        status: 200,
      });
    return new Response(JSON.stringify({ result: [] }), { status: 200 });
  }) as typeof fetch;
  const vendor = servicenowVendor(fake);
  const conn = await vendor.connect({
    instance: "acme",
    auth: "oauth",
    client_id: "cid",
    client_secret: "cs",
  });
  assert.deepEqual(conn.config, { instance: "acme.service-now.com", auth: "oauth" });
  assert.deepEqual(conn.secret, { client_id: "cid", client_secret: "cs" });
  await vendor.call(conn, { method: "GET", path: "/api/now/table/incident" });
  assert.equal(seen[0].url, "https://acme.service-now.com/oauth_token.do");
  assert.equal(
    String(seen[0].init.body),
    "grant_type=client_credentials&client_id=cid&client_secret=cs",
  );
  assert.equal(seen[1].url, "https://acme.service-now.com/api/now/table/incident");
  assert.equal((seen[1].init.headers as Record<string, string>).authorization, "Bearer tok");

  const basic = await vendor.connect({
    instance: "acme",
    auth: "basic",
    username: "eac",
    password: "pw",
  });
  assert.deepEqual(basic.secret, { username: "eac", password: "pw" });
  await vendor.call(basic, { method: "GET", path: "/api/now/table/incident" });
  assert.equal(
    (seen.at(-1)!.init.headers as Record<string, string>).authorization,
    `Basic ${Buffer.from("eac:pw").toString("base64")}`,
  );
  assert.ok(seen.every((s) => s.url.startsWith("https://acme.service-now.com/")));
  await assert.rejects(
    vendor.connect({ instance: "acme", auth: "basic", client_id: "x", client_secret: "y" }),
    /user name and password/,
  );
});

test("the vendor says plainly when ServiceNow refuses the sign-in", async () => {
  const refuse = (async () =>
    new Response(JSON.stringify({ error: "access_denied", error_description: "access_denied" }), {
      status: 401,
    })) as unknown as typeof fetch;
  await assert.rejects(
    servicenowVendor(refuse).connect({
      instance: "acme",
      auth: "oauth",
      client_id: "a",
      client_secret: "b",
    }),
    /ServiceNow refused the sign-in: access_denied/,
  );
  await assert.rejects(
    servicenowVendor(refuse).connect({
      instance: "acme",
      auth: "basic",
      username: "a",
      password: "b",
    }),
    /did not accept the user name and password/,
  );
});
