import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createServer, type IncomingMessage, type ServerResponse } from "node:http";
import type { AddressInfo } from "node:net";
import { after, before, test } from "node:test";
import type { McpServer } from "@modelcontextprotocol/server";
import { connectorSchema, type Connector } from "../../packages/connector-kit/src/connector.ts";
import type { ToolContext } from "../../packages/connector-kit/src/common.ts";
import { assertToolsImplemented } from "../../packages/server/src/service.ts";
import odataModule, { literal, odataTools, odataVendor } from "./index.ts";

// The generic odata module, run against a local server standing in for SAP Gateway (v2) and
// a v4 service such as Dynamics.

const example = JSON.parse(
  readFileSync(
    new URL(
      "../../examples/legacy-apps/connectors/sap-sales-orders/connector.json",
      import.meta.url,
    ),
    "utf8",
  ),
);

type Seen = { method: string; url: string; headers: IncomingMessage["headers"]; body: string };
const seen: Seen[] = [];
let respond: (req: IncomingMessage, res: ServerResponse) => void;
let origin = "";
const server = createServer(async (req, res) => {
  let body = "";
  for await (const chunk of req) body += chunk;
  seen.push({ method: req.method!, url: req.url!, headers: req.headers, body });
  respond(req, res);
});
before(async () => {
  await new Promise<void>((done) => server.listen(0, "127.0.0.1", done));
  origin = `http://127.0.0.1:${(server.address() as AddressInfo).port}`;
});
after(() => server.close());

const ok =
  (data: unknown, status = 200, headers: Record<string, string | string[]> = {}) =>
  (_req: IncomingMessage, res: ServerResponse) => {
    res.writeHead(status, { "content-type": "application/json", ...headers });
    res.end(data === undefined ? undefined : JSON.stringify(data));
  };
const ROOT = "/sap/opu/odata/sap/API_SALES_ORDER_SRV";
// What the agent's $filter came out as, decoded.
const filterOf = (url: string) => new URL(url, "http://x").searchParams.get("$filter");

function sap(odata: Record<string, unknown> = {}, extra: Record<string, unknown> = {}) {
  return connectorSchema.parse({
    ...example,
    ...extra,
    odata: { ...example.odata, serviceRoot: `${origin}${ROOT}`, allowInsecureHttp: true, ...odata },
  });
}

// A v4 service with one tool of each kind.
function dynamics(odata: Record<string, unknown> = {}) {
  return connectorSchema.parse({
    ...example,
    id: "crm",
    fields: [
      { id: "client_id", label: "Client", secret: false },
      { id: "client_secret", label: "Secret", secret: true },
    ],
    limits: [{ id: "owners", label: "Owners", help: "Owners", kind: "list", required: false }],
    verify: { method: "GET", path: "/", label: "CRM answers", show: "value[].name" },
    capabilities: [
      {
        id: "read",
        label: "Read",
        description: "Read accounts",
        required: true,
        scopes: [],
        permissions: [],
        tools: ["crm_list_accounts", "crm_get_account", "crm_score"],
      },
      {
        id: "write",
        label: "Write",
        description: "Change accounts",
        scopes: [],
        permissions: [],
        tools: ["crm_create_account", "crm_update_account", "crm_qualify"],
      },
    ],
    odata: {
      serviceRoot: `${origin}/api/data/v9.2`,
      version: "4",
      allowInsecureHttp: true,
      auth: {
        type: "oauth2",
        tokenUrl: `${origin}/token`,
        clientId: "client_id",
        clientSecret: "client_secret",
        scope: "https://org.crm.dynamics.com/.default",
      },
      tools: [
        {
          name: "crm_list_accounts",
          description: "List accounts",
          kind: "list",
          entitySet: "accounts",
          filters: {
            name: { type: "string", op: "contains", description: "Part of the name" },
            owner: {
              type: "guid",
              property: "_ownerid_value",
              description: "Owner",
              limit: "owners",
            },
            min_revenue: {
              type: "decimal",
              property: "revenue",
              op: "ge",
              description: "Revenue at least",
            },
          },
          select: ["accountid", "name", "_ownerid_value"],
          top: 2,
        },
        {
          name: "crm_get_account",
          description: "Read an account",
          kind: "get",
          entitySet: "accounts",
          key: { id: { type: "guid", property: "accountid", description: "Account id" } },
          select: ["accountid", "name", "_ownerid_value"],
          answerLimits: [{ field: "_ownerid_value", limit: "owners" }],
        },
        {
          name: "crm_create_account",
          description: "Create an account",
          kind: "create",
          entitySet: "accounts",
          fields: {
            name: { type: "string", description: "Name" },
            revenue: { type: "decimal", description: "Revenue", required: false },
            founded: {
              type: "datetime",
              property: "foundedon",
              description: "Founded",
              required: false,
            },
            owner: {
              type: "string",
              property: "ownerid@odata.bind",
              description: "Owner",
              limit: "owners",
            },
          },
          select: ["accountid", "name"],
        },
        {
          name: "crm_update_account",
          description: "Rename an account",
          kind: "update",
          entitySet: "accounts",
          key: { id: { type: "guid", property: "accountid", description: "Account id" } },
          fields: { name: { type: "string", description: "Name" } },
          answerLimits: [{ field: "_ownerid_value", limit: "owners" }],
        },
        {
          name: "crm_score",
          description: "Score an account",
          kind: "function",
          function: "ScoreAccount",
          params: { account: { type: "string", property: "Account", description: "Name" } },
        },
        {
          name: "crm_qualify",
          description: "Qualify a lead",
          kind: "action",
          function: "QualifyLead",
          params: { lead: { type: "guid", property: "LeadId", description: "Lead" } },
        },
      ],
      ...odata,
    },
  });
}

// Registers the tools on a server that only records, so a test can call them directly.
async function run(
  connector: Connector,
  tool: string,
  args: Record<string, unknown>,
  { limits = {} as ToolContext["limits"] } = {},
) {
  const vendor = odataVendor(connector);
  const handlers = new Map<string, (a: Record<string, unknown>) => Promise<unknown>>();
  const recorder = {
    registerTool: (name: string, _c: unknown, h: never) => handlers.set(name, h),
  } as unknown as McpServer;
  const audit: {
    decision: string;
    reason: string | null;
    vendorStatus: number | null;
    target: string | null;
  }[] = [];
  const conn = {
    config: { host: "x", username: "svc", client_id: "cid" },
    secret: { password: "pw", client_secret: "cs" },
  };
  const ctx = {
    limits,
    config: conn.config,
    call: (r) => vendor.call(conn, r),
    audit: async (e) => void audit.push(e),
  } as ToolContext;
  const toolset = odataTools(connector.odata!);
  toolset.register(recorder, ctx, new Set(toolset.names));
  const result = (await handlers.get(tool)!(args)) as {
    isError?: boolean;
    content: { text: string }[];
  };
  return {
    result,
    text: result.content[0].text,
    data: () => JSON.parse(result.content[0].text),
    audit,
  };
}

const order = (org: string, extra = {}) => ({
  __metadata: { uri: "x", type: "API_SALES_ORDER_SRV.A_SalesOrderType" },
  SalesOrder: "1000123",
  SalesOrganization: org,
  SoldToParty: "17100001",
  to_Item: { __deferred: { uri: "x" } },
  ...extra,
});

test("the examples load and run on the generic module", () => {
  for (const connector of [sap(), dynamics()]) {
    const app = odataModule.create({ connector } as never);
    assertToolsImplemented([{ connector, ...app }]);
  }
  const write = dynamics();
  write.odata!.tools[3].destructive = true;
  assert.throws(
    () =>
      assertToolsImplemented([
        { connector: write, ...odataModule.create({ connector: write } as never) },
      ]),
    /capability write has crm_update_account/,
  );
});

test("odata definitions are validated", () => {
  const [list, get] = example.odata.tools;
  const tools = (...t: unknown[]) => ({ tools: t });
  assert.throws(() => sap(tools({ ...get, key: {} })), /get a key/);
  assert.throws(
    () => sap(tools({ ...list, key: get.key })),
    /A list names an entitySet and filters/,
  );
  assert.throws(() => sap(tools({ ...get, function: "F" })), /function and action a function/);
  assert.throws(() => sap(tools({ ...get, orderBy: "x desc" })), /orderBy is for lists/);
  assert.throws(
    () =>
      sap(
        tools({
          ...list,
          filters: { ...list.filters, sales_org: { ...list.filters.sales_org, op: "ne" } },
        }),
      ),
    /Only an eq filter can have a limit/,
  );
  assert.throws(
    () =>
      sap(
        tools({ ...list, filters: { n: { type: "integer", op: "contains", description: "n" } } }),
      ),
    /contains and startswith filter strings/,
  );
  assert.throws(
    () => sap(tools({ ...list, select: ["SalesOrder"] })),
    /limited filter's property must be selected/,
  );
  assert.throws(
    () => sap(tools({ ...list, filters: { cursor: { type: "string", description: "c" } } })),
    /cursor and max_results are taken/,
  );
  assert.throws(() => sap(tools({ ...list, orderBy: "CreationDate desc; drop" })));
  assert.throws(() => sap(tools({ ...list, entitySet: "A_SalesOrder?$expand=x" })));
  assert.throws(() => sap({ paging: "server" }), /Server paging is for OData v4/);
  assert.throws(() => sap({ headers: { Authorization: "x" } }), /Set by the service/);
  assert.throws(() => sap({ serviceRoot: "https://s4.example.com/x?sap-client=100" }), /No query/);
  assert.throws(() => sap({ allowInsecureHttp: false }), /must be https/);
  assert.throws(
    () =>
      dynamics(
        tools({ ...dynamics().odata!.tools[2], answerLimits: [{ field: "x", limit: "owners" }] }),
      ),
    /answerLimits are for reads and updates/,
  );
});

test("a v2 list writes a safe $filter from typed values, with SAP's headers", async () => {
  respond = ok({ d: { results: [order("1010")] } });
  seen.length = 0;
  const { data, audit } = await run(sap(), "sap_list_sales_orders", {
    sales_org: "1010",
    customer: "O'Brien') or (1 eq 1",
    created_from: "2026-01-02",
  });
  const sent = seen[0];
  assert.equal(sent.method, "GET");
  assert.equal(
    filterOf(sent.url),
    "SalesOrganization eq '1010' and SoldToParty eq 'O''Brien'') or (1 eq 1' and CreationDate ge datetime'2026-01-02T00:00:00'",
  );
  const q = new URL(sent.url, "http://x").searchParams;
  assert.equal(q.get("$orderby"), "CreationDate desc");
  assert.equal(q.get("$top"), "50");
  assert.match(q.get("$select")!, /^SalesOrder,SalesOrganization,/);
  assert.doesNotMatch(sent.url, /\+/);
  assert.equal(sent.headers["sap-client"], "100");
  assert.equal(sent.headers.dataserviceversion, "2.0");
  assert.equal(sent.headers.authorization, `Basic ${Buffer.from("svc:pw").toString("base64")}`);
  assert.deepEqual(data(), {
    items: [{ SalesOrder: "1000123", SalesOrganization: "1010", SoldToParty: "17100001" }],
    next: null,
  });
  assert.match(audit[0].target!, /^GET \/A_SalesOrder\?\$filter=/);
});

test("a limited filter: left out it means the agent's values, any other value is refused", async () => {
  const limits = { sales_orgs: ["1010", "1710"] };
  respond = ok({ d: { results: [order("1010"), order("1710")] } });
  seen.length = 0;
  await run(sap(), "sap_list_sales_orders", {}, { limits });
  assert.equal(
    filterOf(seen[0].url),
    "(SalesOrganization eq '1010' or SalesOrganization eq '1710')",
  );
  await run(sap(), "sap_list_sales_orders", { sales_org: "1710" }, { limits });
  assert.equal(filterOf(seen[1].url), "SalesOrganization eq '1710'");
  const refused = await run(sap(), "sap_list_sales_orders", { sales_org: "2000" }, { limits });
  assert.equal(refused.result.isError, true);
  assert.equal(refused.audit[0].decision, "denied");
  assert.equal(seen.length, 2);
  // A service that ignores the filter still cannot hand over another organization's order.
  respond = ok({ d: { results: [order("1010"), order("2000")] } });
  const leaked = await run(sap(), "sap_list_sales_orders", {}, { limits });
  assert.equal(leaked.result.isError, true);
  assert.doesNotMatch(leaked.text, /2000|1000123/);
  assert.equal(leaked.audit[0].decision, "denied");
});

test("paging: $skip when the page is full, the service's skiptoken when it gives one", async () => {
  const full = Array.from({ length: 50 }, () => order("1010"));
  respond = ok({ d: { results: full } });
  seen.length = 0;
  const first = await run(sap(), "sap_list_sales_orders", {});
  assert.equal(first.data().next, "s50");
  await run(sap(), "sap_list_sales_orders", { cursor: "s50", max_results: 10 });
  const q = new URL(seen[1].url, "http://x").searchParams;
  assert.equal(q.get("$skip"), "50");
  assert.equal(q.get("$top"), "10");
  respond = ok({
    d: { results: [order("1010")], __next: `${origin}${ROOT}/A_SalesOrder?$skiptoken=a%2Bb+c` },
  });
  const token = await run(sap(), "sap_list_sales_orders", {});
  assert.equal(token.data().next, "ta+b+c");
  await run(sap(), "sap_list_sales_orders", { cursor: "ta+b+c" });
  assert.match(seen.at(-1)!.url, /\$skiptoken=a%2Bb%2Bc/);
  respond = ok({ d: { results: [order("1010")] } });
  assert.equal((await run(sap(), "sap_list_sales_orders", {})).data().next, null);
  for (const cursor of ["x1", "s-1", ""]) {
    const bad = await run(sap(), "sap_list_sales_orders", { cursor });
    assert.equal(bad.result.isError, true);
    assert.equal(bad.audit[0].decision, "denied");
  }
});

test("server paging asks for the page size and follows the service's skiptoken (v4)", async () => {
  respond = (req, res) =>
    req.url === "/token"
      ? ok({ access_token: "tok", expires_in: 3600 })(req, res)
      : ok({
          "@odata.context": "x",
          value: [{ accountid: "a", name: "A" }],
          "@odata.nextLink": `${origin}/api/data/v9.2/accounts?$select=name&$skiptoken=%3Ccookie%20page%3D%222%22%2F%3E`,
        })(req, res);
  seen.length = 0;
  const { data } = await run(dynamics({ paging: "server" }), "crm_list_accounts", {
    max_results: 1,
  });
  const list = seen.find((s) => s.url.startsWith("/api"))!;
  assert.equal(list.headers.prefer, "odata.maxpagesize=1");
  assert.equal(new URL(list.url, "http://x").searchParams.get("$top"), null);
  assert.equal(list.headers["odata-version"], "4.0");
  assert.equal(list.headers.authorization, "Bearer tok");
  assert.deepEqual(data(), { items: [{ accountid: "a", name: "A" }], next: 't<cookie page="2"/>' });
});

test("v4 literals: contains, guids and decimals unquoted", async () => {
  respond = (req, res) =>
    req.url === "/token"
      ? ok({ access_token: "tok", expires_in: 3600 })(req, res)
      : ok({ value: [] })(req, res);
  seen.length = 0;
  const owner = "0f8fad5b-d9cb-469f-a165-70867728950e";
  await run(
    dynamics(),
    "crm_list_accounts",
    { name: "Ac'me", owner, min_revenue: "1000.50" },
    { limits: { owners: [owner] } },
  );
  assert.equal(
    filterOf(seen.at(-1)!.url),
    `contains(name,'Ac''me') and _ownerid_value eq ${owner} and revenue ge 1000.50`,
  );
  const bad = await run(dynamics(), "crm_list_accounts", { owner: "x' or 1 eq 1" });
  assert.equal(bad.result.isError, true);
});

test("get reads by an encoded key, trims the answer and checks it against the limit", async () => {
  respond = ok({ d: order("1010", { CreationDate: "/Date(1767312000000)/" }) });
  seen.length = 0;
  const limits = { sales_orgs: ["1010"] };
  const { data } = await run(sap(), "sap_get_sales_order", { sales_order: "10/0'1" }, { limits });
  assert.equal(seen[0].url.split("?")[0], `${ROOT}/A_SalesOrder(SalesOrder='10%2F0''1')`);
  assert.deepEqual(data(), {
    SalesOrder: "1000123",
    SalesOrganization: "1010",
    SoldToParty: "17100001",
    CreationDate: "/Date(1767312000000)/",
  });
  for (const answer of [order("2000"), { ...order("x"), SalesOrganization: undefined }]) {
    respond = ok({ d: answer });
    const { result, text, audit } = await run(
      sap(),
      "sap_get_sales_order",
      { sales_order: "1" },
      { limits },
    );
    assert.equal(result.isError, true);
    assert.doesNotMatch(text, /2000|1000123/);
    assert.equal(audit[0].decision, "denied");
  }
});

test("a v2 update reads first, fetches a CSRF token, then MERGEs with it", async () => {
  respond = (req, res) => {
    if (req.headers["x-csrf-token"] === "Fetch")
      return ok({ d: { EntitySets: [] } }, 200, {
        "x-csrf-token": "tok123",
        "set-cookie": [
          "SAP_SESSIONID=abc; path=/; HttpOnly",
          "sap-usercontext=sap-client=100; path=/",
        ],
      })(req, res);
    if (req.method === "GET") return ok({ d: order("1010") })(req, res);
    res.writeHead(204);
    res.end();
  };
  seen.length = 0;
  const limits = { sales_orgs: ["1010"] };
  const { data, result } = await run(
    sap(),
    "sap_set_customer_po",
    { sales_order: "1000123", customer_po: "PO-9" },
    { limits },
  );
  assert.equal(result.isError, undefined);
  assert.deepEqual(data(), { done: true });
  assert.equal(seen.length, 3);
  assert.equal(seen[0].url, `${ROOT}/A_SalesOrder(SalesOrder='1000123')?$select=SalesOrganization`);
  assert.equal(seen[1].url, `${ROOT}/`);
  const write = seen[2];
  assert.equal(write.method, "POST");
  assert.equal(write.headers["x-http-method"], "MERGE");
  assert.equal(write.headers["if-match"], "*");
  assert.equal(write.headers["x-csrf-token"], "tok123");
  assert.equal(write.headers.cookie, "SAP_SESSIONID=abc; sap-usercontext=sap-client=100");
  assert.deepEqual(JSON.parse(write.body), { PurchaseOrderByCustomer: "PO-9" });

  // An order of another sales organization is never changed.
  respond = ok({ d: order("2000") });
  seen.length = 0;
  const refused = await run(
    sap(),
    "sap_set_customer_po",
    { sales_order: "1", customer_po: "x" },
    { limits },
  );
  assert.equal(refused.result.isError, true);
  assert.equal(refused.audit[0].decision, "denied");
  assert.equal(seen.length, 1);
  // Nothing to change is refused before any call.
  const empty = await run(sap(), "sap_set_customer_po", { sales_order: "1" });
  assert.equal(empty.audit[0].decision, "denied");
  assert.equal(seen.length, 1);
  // A service that gives no token stops the write.
  respond = ok({ d: {} }, 200);
  const noToken = await run(sap(), "sap_set_customer_po", { sales_order: "1", customer_po: "x" });
  assert.match(noToken.text, /gave no CSRF token/);
  assert.equal(noToken.audit[0].decision, "failed");
});

test("v4 create and update send typed JSON, and limited fields are checked", async () => {
  const token = (req: IncomingMessage, res: ServerResponse) =>
    ok({ access_token: "tok", expires_in: 3600 })(req, res);
  respond = (req, res) =>
    req.url === "/token"
      ? token(req, res)
      : req.method === "GET"
        ? ok({ "@odata.etag": "W/1", accountid: "a1", _ownerid_value: "u1" })(req, res)
        : ok(
            { "@odata.context": "x", accountid: "a1", name: "Acme", secret: "s" },
            req.method === "POST" ? 201 : 200,
          )(req, res);
  seen.length = 0;
  const limits = { owners: ["/systemusers(u1)"] };
  const created = await run(
    dynamics(),
    "crm_create_account",
    {
      name: "Acme",
      revenue: "1000.50",
      founded: "2020-05-01T10:00:00+02:00",
      owner: "/systemusers(u1)",
    },
    { limits },
  );
  assert.deepEqual(created.data(), { accountid: "a1", name: "Acme" });
  const post = seen.find((s) => s.method === "POST" && s.url.startsWith("/api"))!;
  assert.equal(post.url, "/api/data/v9.2/accounts");
  assert.equal(post.headers.prefer, "return=representation");
  assert.deepEqual(JSON.parse(post.body), {
    name: "Acme",
    revenue: 1000.5,
    foundedon: "2020-05-01T08:00:00.000Z",
    "ownerid@odata.bind": "/systemusers(u1)",
  });
  for (const args of [{ name: "X" }, { name: "X", owner: "/systemusers(u2)" }]) {
    const { result, audit } = await run(dynamics(), "crm_create_account", args, { limits });
    assert.equal(result.isError, true);
    assert.equal(audit[0].decision, "denied");
  }

  seen.length = 0;
  const id = "0f8fad5b-d9cb-469f-a165-70867728950e";
  const updated = await run(
    dynamics(),
    "crm_update_account",
    { id, name: "Acme 2" },
    { limits: { owners: ["u1"] } },
  );
  assert.equal(updated.result.isError, undefined);
  const patch = seen.find((s) => s.method === "PATCH")!;
  assert.equal(patch.url, `/api/data/v9.2/accounts(accountid=${id})`);
  assert.equal(patch.headers["if-match"], "*");
  assert.equal(patch.headers["x-csrf-token"], undefined);
  const refused = await run(
    dynamics(),
    "crm_update_account",
    { id, name: "x" },
    { limits: { owners: ["u9"] } },
  );
  assert.equal(refused.audit[0].decision, "denied");
});

test("functions and actions in both versions", async () => {
  respond = ok({ d: { ScoreAccount: 7 } });
  seen.length = 0;
  const v2 = sap(
    {
      tools: [
        {
          name: "sap_score",
          description: "Score",
          kind: "function",
          function: "ScoreAccount",
          params: { account: { type: "string", property: "Account", description: "A" } },
        },
        {
          name: "sap_release",
          description: "Release",
          kind: "action",
          function: "ReleaseOrder",
          params: { order: { type: "string", property: "SalesOrder", description: "O" } },
        },
      ],
    },
    {
      capabilities: [
        { ...example.capabilities[0], tools: ["sap_score"] },
        { ...example.capabilities[1], tools: ["sap_release"] },
      ],
    },
  );
  assert.deepEqual((await run(v2, "sap_score", { account: "A'1" })).data(), { ScoreAccount: 7 });
  assert.equal(decodeURIComponent(seen[0].url), `${ROOT}/ScoreAccount?Account='A''1'`);
  assert.equal(seen[0].method, "GET");
  respond = (req, res) =>
    req.headers["x-csrf-token"] === "Fetch"
      ? ok({}, 200, { "x-csrf-token": "t" })(req, res)
      : ok({ d: { Released: true } })(req, res);
  await run(v2, "sap_release", { order: "1" });
  assert.equal(seen.at(-1)!.method, "POST");
  assert.equal(decodeURIComponent(seen.at(-1)!.url), `${ROOT}/ReleaseOrder?SalesOrder='1'`);

  respond = (req, res) =>
    req.url === "/token"
      ? ok({ access_token: "tok", expires_in: 3600 })(req, res)
      : ok({ "@odata.context": "x", value: 7 })(req, res);
  seen.length = 0;
  assert.equal((await run(dynamics(), "crm_score", { account: "A" })).data(), 7);
  assert.equal(seen.at(-1)!.url, "/api/data/v9.2/ScoreAccount(Account='A')");
  const lead = "0f8fad5b-d9cb-469f-a165-70867728950e";
  await run(dynamics(), "crm_qualify", { lead });
  assert.equal(seen.at(-1)!.method, "POST");
  assert.equal(seen.at(-1)!.url, "/api/data/v9.2/QualifyLead");
  assert.deepEqual(JSON.parse(seen.at(-1)!.body), { LeadId: lead });
});

test("errors: OData's own text, the token minted again once after a 401", async () => {
  respond = ok(
    { error: { code: "SY/530", message: { lang: "en", value: "Sales order 1 does not exist" } } },
    404,
  );
  const v2 = await run(sap(), "sap_get_sales_order", { sales_order: "1" });
  assert.equal(v2.text, "The service answered 404: SY/530: Sales order 1 does not exist");
  assert.equal(v2.audit[0].decision, "failed");

  let first = true;
  respond = (req, res) => {
    if (req.url === "/token")
      return ok({ access_token: first ? "old" : "new", expires_in: 3600 })(req, res);
    if (req.headers.authorization === "Bearer old") {
      first = false;
      return ok({ error: { code: "0x80040220", message: "Expired" } }, 401)(req, res);
    }
    return ok({ accountid: "a" })(req, res);
  };
  seen.length = 0;
  const v4 = await run(dynamics(), "crm_get_account", {
    id: "0f8fad5b-d9cb-469f-a165-70867728950e",
  });
  assert.equal(v4.result.isError, undefined);
  assert.equal(seen.filter((s) => s.url === "/token").length, 2);
  respond = (req, res) =>
    req.url === "/token"
      ? ok({ access_token: "new", expires_in: 3600 })(req, res)
      : ok({ error: { code: "0x1", message: "No access" } }, 403)(req, res);
  assert.equal(
    (await run(dynamics(), "crm_get_account", { id: "0f8fad5b-d9cb-469f-a165-70867728950e" })).text,
    "The service answered 403: 0x1: No access",
  );
});

test("connect proves an OAuth client and the verify probe reads the service document", async () => {
  respond = (req, res) =>
    req.url === "/token"
      ? ok({ access_token: "t", expires_in: 3600 })(req, res)
      : ok({ value: [{ name: "accounts" }] })(req, res);
  seen.length = 0;
  const vendor = odataVendor(dynamics());
  const conn = await vendor.connect({ client_id: "cid", client_secret: "cs" });
  assert.equal(seen[0].url, "/token");
  const form = new URLSearchParams(seen[0].body);
  assert.equal(form.get("scope"), "https://org.crm.dynamics.com/.default");
  assert.equal(form.get("client_secret"), "cs");
  const doc = await vendor.call(conn, { method: "GET", path: "/" });
  assert.deepEqual(doc.data, { value: [{ name: "accounts" }] });
  await assert.rejects(
    vendor.call(conn, { method: "GET", path: "/../../other" }),
    /leaves this app's address/,
  );
});

test("literals of each type, refusing a value of another type", () => {
  const cases: [Parameters<typeof literal>[0], unknown, string, string][] = [
    ["string", "it's", "'it''s'", "'it''s'"],
    ["integer", 42, "42", "42"],
    ["number", 1.5, "1.5d", "1.5"],
    ["decimal", "-10.25", "-10.25M", "-10.25"],
    ["boolean", false, "false", "false"],
    [
      "guid",
      "0f8fad5b-d9cb-469f-a165-70867728950e",
      "guid'0f8fad5b-d9cb-469f-a165-70867728950e'",
      "0f8fad5b-d9cb-469f-a165-70867728950e",
    ],
    [
      "datetime",
      "2026-01-02T03:04:05+01:00",
      "datetime'2026-01-02T02:04:05'",
      "2026-01-02T02:04:05.000Z",
    ],
    [
      "datetimeoffset",
      "2026-01-02",
      "datetimeoffset'2026-01-02T00:00:00.000Z'",
      "2026-01-02T00:00:00.000Z",
    ],
  ];
  for (const [type, value, v2, v4] of cases) {
    assert.equal(literal(type, value, "2"), v2);
    assert.equal(literal(type, value, "4"), v4);
  }
  for (const [type, value] of [
    ["integer", "1 or 1 eq 1"],
    ["decimal", "1M or true"],
    ["guid", "x"],
    ["datetime", "2026-13-45"],
    ["boolean", "true"],
  ] as const)
    assert.throws(() => literal(type, value, "4"), /is not a/);
});

test("a query key is appended without rewriting the encoded $filter", async () => {
  respond = ok({ d: { results: [] } });
  seen.length = 0;
  const keyed = sap(
    { auth: { type: "query", param: "apikey", value: "password" } },
    { fields: [{ id: "password", label: "API key", secret: true }] },
  );
  await run(keyed, "sap_list_sales_orders", { customer: "A B" });
  assert.match(seen[0].url, /\$filter=SoldToParty%20eq%20%27A%20B%27&.*&apikey=pw$/);
  assert.equal(seen[0].headers.authorization, undefined);
});
