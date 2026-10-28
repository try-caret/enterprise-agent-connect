import assert from "node:assert/strict";
import { test } from "node:test";
import { day, sapTools } from "./tools.ts";
import { normalizeHost, sapVendor } from "./vendor.ts";
import { loadConnectors } from "../../packages/connector-kit/src/connector.ts";
import type { ToolContext } from "../../packages/connector-kit/src/common.ts";
import type { VendorConnection, VendorRequest } from "../../packages/connector-kit/src/types.ts";
import { assertToolsImplemented } from "../../packages/server/src/service.ts";

// The SAP tools against a fake S/4HANA speaking OData v2 JSON (d.results, d), reached through
// the real vendor client, so sign-in, sap-client and the CSRF handshake are exercised too. The
// fake ignores $filter, as a lax server might, so answers must also be filtered here.

const HOST = "my123456-api.s4hana.cloud.sap";
const BASIC = `Basic ${Buffer.from("AGENT:pw").toString("base64")}`;
const date = (d: string) => `/Date(${Date.parse(`${d}T00:00:00Z`)})/`;

const SALES_ORDERS = [
  { SalesOrder: "1", SalesOrganization: "1010", plant: "1010" },
  { SalesOrder: "2", SalesOrganization: "2020", plant: "2020" },
  { SalesOrder: "3", SalesOrganization: "1010", plant: "2020" },
].map((o) => ({
  __metadata: { uri: "x", type: "API_SALES_ORDER_SRV.A_SalesOrderType" },
  SalesOrder: o.SalesOrder,
  SalesOrderType: "OR",
  SalesOrganization: o.SalesOrganization,
  DistributionChannel: "10",
  OrganizationDivision: "00",
  SoldToParty: "17100001",
  PurchaseOrderByCustomer: "",
  CreationDate: date("2026-10-01"),
  RequestedDeliveryDate: date("2026-11-01"),
  TotalNetAmount: "1250.50",
  TransactionCurrency: "EUR",
  OverallSDProcessStatus: "A",
  OverallTotalDeliveryStatus: "",
  to_Item: {
    results: [
      {
        SalesOrderItem: "10",
        Material: "TG11",
        SalesOrderItemText: "Trading good",
        RequestedQuantity: "5",
        RequestedQuantityUnit: "PC",
        NetAmount: "1250.50",
        ProductionPlant: o.plant,
        SDProcessStatus: "A",
        DeliveryStatus: "A",
      },
    ],
  },
}));

const PURCHASE_ORDERS = [
  { PurchaseOrder: "4500000001", CompanyCode: "1010", PurchasingOrganization: "1010" },
  { PurchaseOrder: "4500000002", CompanyCode: "2020", PurchasingOrganization: "1010" },
  { PurchaseOrder: "4500000003", CompanyCode: "1010", PurchasingOrganization: "2020" },
].map((o) => ({
  ...o,
  PurchaseOrderType: "NB",
  PurchasingGroup: "001",
  Supplier: "17300001",
  PurchaseOrderDate: date("2026-10-02"),
  DocumentCurrency: "EUR",
  PurchasingProcessingStatus: "05",
  PurchasingDocumentDeletionCode: "",
  to_PurchaseOrderItem: {
    results: [
      {
        PurchaseOrderItem: "10",
        Material: "RM12",
        PurchaseOrderItemText: "Raw material",
        Plant: "1010",
        OrderQuantity: "100",
        PurchaseOrderQuantityUnit: "KG",
        NetPriceAmount: "2.50",
        NetPriceQuantity: "1",
        IsCompletelyDelivered: true,
        IsFinallyInvoiced: false,
        PurchasingDocumentDeletionCode: "",
      },
    ],
  },
}));

const PARTNERS = [
  ["17100001", "Inlandskunde DE 1", "17100001", ""],
  ["17300001", "Inlandslieferant DE 1", "", "17300001"],
  ["9980000001", "Ada Lovelace", "", ""],
].map(([id, name, customer, supplier], i) => ({
  BusinessPartner: id,
  BusinessPartnerFullName: name,
  BusinessPartnerCategory: i === 2 ? "1" : "2",
  BusinessPartnerIsBlocked: i === 1,
  Customer: customer,
  Supplier: supplier,
  to_BusinessPartnerAddress: { results: [{ CityName: "Berlin", Country: "DE" }] },
}));

const STOCK = [
  ["1010", "101A", "01", "", "40"],
  ["1010", "101B", "01", "", "2"],
  ["1010", "101A", "02", "", "5"],
  ["1010", "101A", "01", "K", "7"],
  ["2020", "201A", "01", "", "99"],
].map(([plant, sloc, type, special, qty]) => ({
  Material: "TG11",
  Plant: plant,
  StorageLocation: sloc,
  Batch: "",
  InventoryStockType: type,
  InventorySpecialStockType: special,
  MatlWrhsStkQtyInMatlBaseUnit: qty,
  MaterialBaseUnit: "PC",
}));

const INVOICES = ["1010", "2020"].map((cc, i) => ({
  SupplierInvoice: `510000000${i}`,
  FiscalYear: "2026",
  CompanyCode: cc,
  InvoicingParty: "17300001",
  SupplierInvoiceIDByInvcgParty: `INV-${i}`,
  DocumentDate: date("2026-10-03"),
  PostingDate: date("2026-10-04"),
  InvoiceGrossAmount: "119.00",
  DocumentCurrency: "EUR",
  SupplierInvoiceStatus: "5",
  PaymentBlockingReason: i ? "R" : "",
}));

const ACCOUNT_ITEMS = ["1010", "2020"].map((cc, i) => ({
  CompanyCode: cc,
  FiscalYear: "2026",
  AccountingDocument: `190000000${i}`,
  AccountingDocumentItem: "1",
  AccountingDocumentType: "DR",
  Customer: "17100001",
  DocumentReferenceID: "",
  PostingDate: date("2026-10-04"),
  NetDueDate: date("2026-11-03"),
  AmountInTransactionCurrency: "500.00",
  TransactionCurrency: "EUR",
  IsCleared: i === 1,
  ClearingAccountingDocument: i ? "1400000000" : "",
  ClearingDate: i ? date("2026-10-20") : null,
}));

type Sent = { url: URL; method: string; headers: Headers; body: any };
type Reply = { status?: number; body?: unknown; headers?: Record<string, string> };

function fakeSap(
  options: {
    oauth?: boolean;
    overrides?: Record<string, Reply>;
    // Status of the CSRF token fetch.
    csrfStatus?: number;
  } = {},
) {
  const sent: Sent[] = [];
  let minted = 0;
  const valid = new Set<string>();
  const fetchImpl = (async (input: string | URL, init: RequestInit) => {
    const url = new URL(input);
    const headers = new Headers(init.headers);
    const raw = init.body;
    const body =
      typeof raw === "string" && headers.get("content-type") === "application/json"
        ? JSON.parse(raw)
        : raw instanceof URLSearchParams
          ? raw.toString()
          : raw;
    sent.push({ url, method: init.method ?? "GET", headers, body });
    const reply = route(url, init.method ?? "GET", headers, body);
    const out = new Headers({ "content-type": "application/json" });
    for (const [k, v] of Object.entries(reply.headers ?? {}))
      if (k === "set-cookie") v.split(", ").forEach((c) => out.append(k, c));
      else out.set(k, v);
    return new Response(
      typeof reply.body === "string" ? reply.body : JSON.stringify(reply.body ?? {}),
      { status: reply.status ?? 200, headers: out },
    );
  }) as unknown as typeof fetch;

  function route(url: URL, method: string, headers: Headers, body: any): Reply {
    const path = decodeURIComponent(url.pathname);
    if (path === "/sap/bc/sec/oauth2/token") {
      if (headers.get("authorization") !== BASIC || body !== "grant_type=client_credentials")
        return { status: 401, body: { error: "invalid_client", error_description: "Bad client" } };
      minted += 1;
      valid.add(`Bearer tok${minted}`);
      return { body: { access_token: `tok${minted}`, token_type: "Bearer", expires_in: 3600 } };
    }
    const auth = headers.get("authorization");
    const good = options.oauth ? valid.has(auth ?? "") : auth === BASIC;
    if (!good)
      return {
        status: 401,
        body: "<html>Logon failed</html>",
        headers: { "content-type": "text/html" },
      };
    const override = options.overrides?.[path];
    if (override) return override;
    if (headers.get("x-csrf-token") === "Fetch") {
      if (options.csrfStatus)
        return {
          status: options.csrfStatus,
          body: { error: { message: { value: "No authorization" } } },
        };
      return {
        body: { d: { EntitySets: ["A_SalesOrder"] } },
        headers: {
          "x-csrf-token": "T1",
          "set-cookie":
            "SAP_SESSIONID_X=abc; path=/; HttpOnly, sap-usercontext=sap-client=100; path=/",
        },
      };
    }
    if (method === "POST") {
      if (
        headers.get("x-csrf-token") !== "T1" ||
        !headers.get("cookie")?.includes("SAP_SESSIONID_X=abc")
      )
        return {
          status: 403,
          body: "CSRF token validation failed",
          headers: { "x-csrf-token": "Required", "content-type": "text/plain" },
        };
      if (path.endsWith("/A_SalesOrder"))
        return {
          status: 201,
          body: { d: { ...SALES_ORDERS[0], SalesOrder: "123", TotalNetAmount: "99.00" } },
        };
      if (path.endsWith("/A_PurchaseOrder"))
        return {
          status: 201,
          body: {
            d: {
              ...PURCHASE_ORDERS[0],
              PurchaseOrder: "4500000099",
              PurchasingProcessingStatus: "03",
            },
          },
        };
    }
    const top = Number(url.searchParams.get("$top") ?? 1000);
    const skip = Number(url.searchParams.get("$skip") ?? 0);
    const list = (rows: object[]) => ({ body: { d: { results: rows.slice(skip, skip + top) } } });
    const one = (rows: Row[], key: string) => {
      const id = /\('([^']*)'\)$/.exec(path)?.[1];
      const row = rows.find((r) => r[key] === id);
      return row
        ? { body: { d: row } }
        : {
            status: 404,
            body: {
              error: {
                code: "/IWBEP/CM_MGW_RT/020",
                message: { lang: "en", value: "Resource not found" },
              },
            },
          };
    };
    if (/\/API_[A-Z_]+\/$/.test(path)) return { body: { d: { EntitySets: [] } } };
    if (path.endsWith("/A_BusinessPartner")) return list(PARTNERS);
    if (path.endsWith("/A_SalesOrder")) return list(SALES_ORDERS);
    if (path.includes("/A_SalesOrder(")) return one(SALES_ORDERS, "SalesOrder");
    if (path.endsWith("/A_PurchaseOrder")) return list(PURCHASE_ORDERS);
    if (path.includes("/A_PurchaseOrder(")) return one(PURCHASE_ORDERS, "PurchaseOrder");
    if (path.endsWith("/A_ProductDescription"))
      return list([{ Product: "TG11", Language: "EN", ProductDescription: "Trading good" }]);
    if (path.endsWith("/A_MatlStkInAcctMod")) return list(STOCK);
    if (path.endsWith("/A_SupplierInvoice")) return list(INVOICES);
    if (path.endsWith("/A_OperationalAcctgDocItemCube")) return list(ACCOUNT_ITEMS);
    return { status: 404, body: { error: { message: { value: "No service found" } } } };
  }

  return { sent, fetchImpl, revoke: () => valid.clear() };
}
type Row = Record<string, any>;

const cloud: VendorConnection = {
  config: { deployment: "cloud", auth: "basic", host: HOST, user: "AGENT", system: HOST },
  secret: { password: "pw" },
};

function setup(limits: Record<string, unknown> = {}, options: Parameters<typeof fakeSap>[0] = {}) {
  const sap = fakeSap(options);
  const vendor = sapVendor(sap.fetchImpl);
  const audit: {
    tool: string;
    target: string | null;
    decision: string;
    vendorStatus: number | null;
  }[] = [];
  const ctx = {
    limits,
    config: {},
    call: (req: VendorRequest) => vendor.call(cloud, req),
    async audit(e) {
      audit.push(e);
    },
  } as ToolContext;
  const tools = new Map<string, (args: unknown) => Promise<any>>();
  const schemas = new Map<string, any>();
  const server = {
    registerTool: (name: string, def: { inputSchema: any }, fn: never) => {
      tools.set(name, fn);
      schemas.set(name, def.inputSchema);
    },
  };
  sapTools.register(server as never, ctx, new Set(sapTools.names));
  const run = async (name: string, args: object) => {
    // As the MCP server does: defaults filled by the schema before the tool runs.
    const r = await tools.get(name)!(schemas.get(name).parse(args));
    const text = r.content[0].text as string;
    return { isError: Boolean(r.isError), text, body: r.isError ? null : JSON.parse(text) };
  };
  const reads = () => sap.sent.filter((s) => !s.url.pathname.endsWith("/token"));
  return { run, sent: sap.sent, reads, audit, vendor, sap };
}

const filterOf = (s: Sent) => s.url.searchParams.get("$filter");

test("the connector's capabilities and tools line up; only creating purchase orders is destructive", () => {
  const connector = loadConnectors().find((c) => c.id === "sap-s4hana")!;
  assertToolsImplemented([{ connector, toolset: sapTools, vendor: sapVendor() }]);
  const destructive = connector.capabilities.filter((c) => c.destructive).map((c) => c.id);
  assert.deepEqual(destructive, ["create_purchasing"]);
  assert.ok(connector.capabilities.every((c) => !c.tools.some((t) => t.includes("approve"))));
});

test("connect signs in with basic auth and names the system", async () => {
  const { vendor, sent } = setup();
  const conn = await vendor.connect({
    deployment: "cloud",
    auth: "basic",
    host: ` https://${HOST.toUpperCase()}/ui `,
    sap_client: "",
    user: "AGENT",
    password: "pw",
  });
  assert.deepEqual(conn, { ...cloud, accountLabel: HOST });
  assert.equal(sent[0].url.href, `https://${HOST}/sap/opu/odata/sap/API_BUSINESS_PARTNER/`);
  assert.equal(sent[0].headers.get("authorization"), BASIC);
  await assert.rejects(
    vendor.connect({ deployment: "cloud", host: HOST, user: "AGENT", password: "wrong" }),
    /did not accept the user name and password for AGENT/,
  );
  // Fields only asked on premise are ignored on Cloud.
  const ignored = await vendor.connect({
    deployment: "cloud",
    host: HOST,
    gateway_host: "sap-gw.example.internal:44300",
    sap_client: "100",
    user: "AGENT",
    password: "pw",
  });
  assert.equal(ignored.accountLabel, HOST);
  assert.equal(ignored.config.sapClient, undefined);
  assert.throws(() => normalizeHost("http://sap.example:8000"), /over https/);
  assert.throws(() => normalizeHost("sap example"), /API host name/);
});

test("on premise, every call carries the client and the system names it", async () => {
  const { vendor, sent } = setup();
  const conn = await vendor.connect({
    deployment: "on_premise",
    // Ignored on premise: basic auth for a technical user.
    auth: "oauth",
    host: HOST,
    gateway_host: "sap-gw.example.internal:44300",
    sap_client: "100",
    user: "AGENT",
    password: "pw",
  });
  assert.deepEqual(conn.config, {
    deployment: "on_premise",
    auth: "basic",
    host: "sap-gw.example.internal:44300",
    sapClient: "100",
    user: "AGENT",
    system: "sap-gw.example.internal:44300 client 100",
  });
  await vendor.call(conn, {
    method: "GET",
    path: "/sap/opu/odata/sap/API_SALES_ORDER_SRV/A_SalesOrder?$top=1",
  });
  assert.equal(
    sent.at(-1)!.url.href,
    "https://sap-gw.example.internal:44300/sap/opu/odata/sap/API_SALES_ORDER_SRV/A_SalesOrder?$top=1&sap-client=100",
  );
  await assert.rejects(
    vendor.connect({
      deployment: "on_premise",
      gateway_host: "sap",
      sap_client: "1000",
      user: "A",
      password: "p",
    }),
    /three digits/,
  );
});

test("OAuth signs the communication user in for a token and mints a new one after a 401", async () => {
  const { vendor, sap, sent } = setup({}, { oauth: true });
  const conn = await vendor.connect({
    deployment: "cloud",
    auth: "oauth",
    host: HOST,
    user: "AGENT",
    password: "pw",
  });
  assert.equal(conn.config.auth, "oauth");
  assert.equal(sent[0].url.pathname, "/sap/bc/sec/oauth2/token");
  assert.equal(sent[0].headers.get("authorization"), BASIC);
  assert.equal(sent[1].headers.get("authorization"), "Bearer tok1");
  // SAP drops the token (a restart, say): the next call mints a new one once and succeeds.
  sap.revoke();
  const res = await vendor.call(conn, {
    method: "GET",
    path: "/sap/opu/odata/sap/API_BUSINESS_PARTNER/A_BusinessPartner",
  });
  assert.equal(res.status, 200);
  assert.deepEqual(
    sent.slice(2).map((s) => [s.url.pathname.split("/").pop(), s.headers.get("authorization")]),
    [
      ["A_BusinessPartner", "Bearer tok1"],
      ["token", BASIC],
      ["A_BusinessPartner", "Bearer tok2"],
    ],
  );
  await assert.rejects(
    vendor.connect({
      deployment: "cloud",
      auth: "oauth",
      host: HOST,
      user: "AGENT",
      password: "nope",
    }),
    /SAP refused the sign-in \(401\): Bad client/,
  );
});

test("the services check lists the services the user can open and names the rest", async () => {
  const { vendor } = setup(
    {},
    {
      overrides: {
        "/sap/opu/odata/sap/API_SALES_ORDER_SRV/": { status: 403, body: {} },
        "/sap/opu/odata/sap/API_PRODUCT_SRV/": { status: 404, body: {} },
      },
    },
  );
  const result = await vendor.checks!.services(cloud, {
    expected: ["API_BUSINESS_PARTNER", "API_SALES_ORDER_SRV", "API_PRODUCT_SRV"],
  });
  assert.deepEqual(result.values, ["API_BUSINESS_PARTNER"]);
  assert.match(
    result.warning!,
    /API_SALES_ORDER_SRV \(not allowed for this user\), API_PRODUCT_SRV \(not active\)/,
  );
});

test("sap_find_business_partners trims partners and pages", async () => {
  const { run, reads } = setup();
  const r = await run("sap_find_business_partners", { name: "Inland", role: "customer", limit: 2 });
  assert.equal(r.isError, false);
  assert.deepEqual(r.body.partners[0], {
    partner: "17100001",
    name: "Inlandskunde DE 1",
    kind: "organization",
    customer: "17100001",
    supplier: null,
    blocked: false,
    city: "Berlin",
    country: "DE",
  });
  assert.equal(r.body.next, "2");
  const sent = reads()[0];
  assert.equal(filterOf(sent), "substringof('Inland',BusinessPartnerFullName) and Customer ne ''");
  assert.equal(sent.url.searchParams.get("$top"), "3");
  const second = await run("sap_find_business_partners", { cursor: "2", limit: 2 });
  assert.equal(second.body.partners[0].kind, "person");
  assert.equal(second.body.next, null);
  // Quotes are doubled, so a name cannot end the string.
  await run("sap_find_business_partners", { name: "O'Brien' or 1 eq 1" });
  assert.equal(
    filterOf(reads().at(-1)!),
    "substringof('O''Brien'' or 1 eq 1',BusinessPartnerFullName)",
  );
});

test("sap_list_sales_orders keeps to the agent's sales organizations", async () => {
  const { run, reads } = setup({ sales_organizations: ["1010"] });
  const r = await run("sap_list_sales_orders", { created_from: "2026-10-01" });
  assert.deepEqual(
    r.body.sales_orders.map((o: { sales_order: string }) => o.sales_order),
    ["1", "3"],
  );
  assert.deepEqual(r.body.sales_orders[0], {
    sales_order: "1",
    type: "OR",
    sales_organization: "1010",
    distribution_channel: "10",
    division: "00",
    sold_to: "17100001",
    customer_reference: null,
    created: "2026-10-01",
    requested_delivery: "2026-11-01",
    net_amount: 1250.5,
    currency: "EUR",
    status: "open",
    delivery_status: null,
  });
  assert.equal(
    filterOf(reads()[0]),
    "(SalesOrganization eq '1010') and CreationDate ge datetime'2026-10-01T00:00:00'",
  );

  const refused = await run("sap_list_sales_orders", { sales_organization: "2020" });
  assert.equal(refused.isError, true);
  assert.match(refused.text, /not allowed to use sales organization 2020/);
  assert.equal(reads().length, 1);

  const bad = await run("sap_list_sales_orders", { created_to: "2026-13-45" });
  assert.match(bad.text, /not a date/);
});

test("sap_get_sales_order reads lines and refuses orders outside the limits", async () => {
  const { run, audit, reads } = setup({ sales_organizations: ["1010"], plants: ["1010"] });
  const r = await run("sap_get_sales_order", { sales_order: "1" });
  assert.equal(r.isError, false);
  assert.deepEqual(r.body.lines, [
    {
      line: "10",
      material: "TG11",
      description: "Trading good",
      quantity: 5,
      unit: "PC",
      net_amount: 1250.5,
      plant: "1010",
      status: "open",
      delivery_status: "not delivered",
    },
  ]);
  assert.equal(reads()[0].url.pathname, "/sap/opu/odata/sap/API_SALES_ORDER_SRV/A_SalesOrder('1')");
  assert.equal(reads()[0].url.searchParams.get("$expand"), "to_Item");

  const otherOrg = await run("sap_get_sales_order", { sales_order: "2" });
  assert.match(otherOrg.text, /sales organization this agent is not allowed to use/);
  assert.deepEqual([audit.at(-1)!.decision, audit.at(-1)!.vendorStatus], ["denied", 200]);
  const otherPlant = await run("sap_get_sales_order", { sales_order: "3" });
  assert.match(otherPlant.text, /lines at a plant this agent is not allowed to use/);

  const missing = await run("sap_get_sales_order", { sales_order: "404" });
  assert.equal(missing.text, "SAP answered 404: Resource not found");
  assert.deepEqual([audit.at(-1)!.decision, audit.at(-1)!.vendorStatus], ["failed", 404]);
});

test("sap_create_sales_order fetches a CSRF token and posts the order with its session", async () => {
  const { run, sent, audit } = setup({ sales_organizations: ["1010"], plants: ["1010"] });
  const r = await run("sap_create_sales_order", {
    sales_organization: "1010",
    distribution_channel: "10",
    division: "00",
    sold_to: "17100001",
    customer_reference: "PO-77",
    requested_delivery_date: "2026-11-15",
    lines: [{ material: "TG11", quantity: 2.5, unit: "PC", plant: "1010" }],
  });
  assert.equal(r.isError, false);
  assert.deepEqual(r.body, { sales_order: "123", net_amount: 99, currency: "EUR", status: "open" });
  const [fetchToken, post] = sent;
  assert.equal(fetchToken.method, "GET");
  assert.equal(fetchToken.url.pathname, "/sap/opu/odata/sap/API_SALES_ORDER_SRV/");
  assert.equal(fetchToken.headers.get("x-csrf-token"), "Fetch");
  assert.equal(post.headers.get("x-csrf-token"), "T1");
  assert.equal(post.headers.get("cookie"), "SAP_SESSIONID_X=abc; sap-usercontext=sap-client=100");
  assert.deepEqual(post.body, {
    SalesOrderType: "OR",
    SalesOrganization: "1010",
    DistributionChannel: "10",
    OrganizationDivision: "00",
    SoldToParty: "17100001",
    PurchaseOrderByCustomer: "PO-77",
    RequestedDeliveryDate: date("2026-11-15"),
    to_Item: [
      {
        SalesOrderItem: "10",
        Material: "TG11",
        RequestedQuantity: "2.5",
        RequestedQuantityUnit: "PC",
        ProductionPlant: "1010",
      },
    ],
  });
  assert.equal(audit.at(-1)!.target, "1010/17100001 sales order 123");

  const base = { distribution_channel: "10", division: "00", sold_to: "17100001" };
  const wrongOrg = await run("sap_create_sales_order", {
    ...base,
    sales_organization: "2020",
    lines: [{ material: "TG11", quantity: 1, plant: "1010" }],
  });
  assert.match(wrongOrg.text, /sales organization 2020/);
  const noPlant = await run("sap_create_sales_order", {
    ...base,
    sales_organization: "1010",
    lines: [{ material: "TG11", quantity: 1 }],
  });
  assert.match(noPlant.text, /Give a plant on every line/);
  const wrongPlant = await run("sap_create_sales_order", {
    ...base,
    sales_organization: "1010",
    lines: [{ material: "TG11", quantity: 1, plant: "2020" }],
  });
  assert.match(wrongPlant.text, /plant 2020/);
  assert.equal(sent.length, 2, "refused writes reach SAP not at all");
});

test("a refused CSRF token fetch fails the write without posting", async () => {
  const { run, sent } = setup({}, { csrfStatus: 403 });
  const r = await run("sap_create_sales_order", {
    sales_organization: "1010",
    distribution_channel: "10",
    division: "00",
    sold_to: "17100001",
    lines: [{ material: "TG11", quantity: 1 }],
  });
  assert.equal(r.text, "SAP answered 403: No authorization");
  assert.ok(!sent.some((s) => s.method === "POST"));
});

test("purchase order reads keep to company codes and purchasing organizations", async () => {
  const { run, reads } = setup({ company_codes: ["1010"], purchasing_organizations: ["1010"] });
  const list = await run("sap_list_purchase_orders", { supplier: "17300001" });
  assert.deepEqual(
    list.body.purchase_orders.map((o: { purchase_order: string }) => o.purchase_order),
    ["4500000001"],
  );
  assert.equal(list.body.purchase_orders[0].status, "approved");
  assert.equal(
    filterOf(reads()[0]),
    "(CompanyCode eq '1010') and (PurchasingOrganization eq '1010') and Supplier eq '17300001'",
  );
  const one = await run("sap_get_purchase_order", { purchase_order: "4500000001" });
  assert.deepEqual(one.body.lines[0], {
    line: "10",
    material: "RM12",
    description: "Raw material",
    plant: "1010",
    quantity: 100,
    unit: "KG",
    net_price: 2.5,
    price_per: 1,
    delivered: true,
    invoiced: false,
    deleted: false,
  });
  for (const po of ["4500000002", "4500000003"]) {
    const refused = await run("sap_get_purchase_order", { purchase_order: po });
    assert.match(refused.text, /not allowed to use/);
  }
  const askedOther = await run("sap_list_purchase_orders", { company_code: "2020" });
  assert.match(askedOther.text, /company code 2020/);
});

test("sap_create_purchase_order checks every limit before posting and cannot approve", async () => {
  const limits = { company_codes: ["1010"], purchasing_organizations: ["1010"], plants: ["1010"] };
  const { run, sent } = setup(limits);
  const base = {
    company_code: "1010",
    purchasing_organization: "1010",
    purchasing_group: "001",
    supplier: "17300001",
  };
  const r = await run("sap_create_purchase_order", {
    ...base,
    lines: [
      { material: "RM12", plant: "1010", quantity: 100, net_price: 2.5 },
      {
        description: "Consulting",
        material_group: "SERVICES",
        plant: "1010",
        quantity: 1,
        unit: "EA",
      },
    ],
  });
  assert.deepEqual(r.body, { purchase_order: "4500000099", status: "in approval" });
  const post = sent.find((s) => s.method === "POST")!;
  assert.deepEqual(post.body.to_PurchaseOrderItem, [
    {
      PurchaseOrderItem: "10",
      Plant: "1010",
      OrderQuantity: "100",
      Material: "RM12",
      NetPriceAmount: "2.5",
    },
    {
      PurchaseOrderItem: "20",
      Plant: "1010",
      OrderQuantity: "1",
      PurchaseOrderItemText: "Consulting",
      MaterialGroup: "SERVICES",
      PurchaseOrderQuantityUnit: "EA",
    },
  ]);
  assert.equal(post.body.PurchaseOrderType, "NB");
  const before = sent.length;
  for (const [args, why] of [
    [
      { ...base, company_code: "2020", lines: [{ material: "RM12", plant: "1010", quantity: 1 }] },
      /company code 2020/,
    ],
    [
      {
        ...base,
        purchasing_organization: "2020",
        lines: [{ material: "RM12", plant: "1010", quantity: 1 }],
      },
      /purchasing organization 2020/,
    ],
    [{ ...base, lines: [{ material: "RM12", plant: "2020", quantity: 1 }] }, /plant 2020/],
    [
      { ...base, lines: [{ description: "No group", plant: "1010", quantity: 1 }] },
      /a material, or a description and a material group/,
    ],
  ] as const) {
    const refused = await run("sap_create_purchase_order", args);
    assert.match(refused.text, why);
  }
  assert.equal(sent.length, before);
});

test("sap_get_stock sums unrestricted stock per allowed plant", async () => {
  const { run, reads } = setup({ plants: ["1010"] });
  const r = await run("sap_get_stock", { material: "TG11" });
  assert.deepEqual(r.body.unrestricted_by_plant, { "1010": 42 });
  assert.equal(r.body.unit, "PC");
  assert.equal(r.body.stock.length, 4);
  assert.ok(r.body.stock.every((s: { plant: string }) => s.plant === "1010"));
  assert.deepEqual(r.body.stock[2], {
    plant: "1010",
    storage_location: "101A",
    batch: null,
    stock_type: "quality inspection",
    special_stock: null,
    quantity: 5,
  });
  assert.equal(filterOf(reads()[0]), "Material eq 'TG11' and (Plant eq '1010')");
  const refused = await run("sap_get_stock", { material: "TG11", plant: "2020" });
  assert.match(refused.text, /plant 2020/);
});

test("sap_find_products searches English descriptions", async () => {
  const { run, reads } = setup();
  const r = await run("sap_find_products", { description: "Trading" });
  assert.deepEqual(r.body, {
    products: [{ product: "TG11", description: "Trading good" }],
    next: null,
  });
  assert.equal(
    filterOf(reads()[0]),
    "Language eq 'EN' and substringof('Trading',ProductDescription)",
  );
});

test("invoices and accounting items keep to company codes", async () => {
  const { run, reads } = setup({ company_codes: ["1010"] });
  const invoices = await run("sap_list_supplier_invoices", { supplier: "17300001" });
  assert.deepEqual(invoices.body.invoices, [
    {
      invoice: "5100000000",
      fiscal_year: "2026",
      company_code: "1010",
      supplier: "17300001",
      supplier_reference: "INV-0",
      document_date: "2026-10-03",
      posted: "2026-10-04",
      gross_amount: 119,
      currency: "EUR",
      status: "posted",
      payment_block: null,
    },
  ]);
  const items = await run("sap_list_account_items", {
    side: "customer",
    account: "17100001",
    status: "all",
  });
  assert.equal(items.body.items.length, 1);
  assert.equal(items.body.items[0].due, "2026-11-03");
  assert.equal(
    filterOf(reads().at(-1)!),
    "FinancialAccountType eq 'D' and (CompanyCode eq '1010') and Customer eq '17100001'",
  );
  await run("sap_list_account_items", { side: "supplier" });
  assert.match(
    filterOf(reads().at(-1)!)!,
    /^FinancialAccountType eq 'K' .* and IsCleared eq false$/,
  );
  const refused = await run("sap_list_supplier_invoices", { company_code: "2020" });
  assert.match(refused.text, /company code 2020/);
});

test("an unreachable SAP host is a failed call, not an internal error", async () => {
  const vendor = sapVendor((async () => {
    throw new TypeError("fetch failed");
  }) as unknown as typeof fetch);
  const audit: { decision: string; vendorStatus: number | null }[] = [];
  const ctx = {
    limits: {},
    config: {},
    call: (req: VendorRequest) => vendor.call(cloud, req),
    async audit(e) {
      audit.push(e);
    },
  } as ToolContext;
  let tool: (a: unknown) => Promise<any> = async () => null;
  sapTools.register(
    {
      registerTool: (n: string, _d: unknown, fn: never) => n === "sap_find_products" && (tool = fn),
    } as never,
    ctx,
    new Set(["sap_find_products"]),
  );
  const r = await tool({ limit: 5 });
  assert.equal(r.content[0].text, `${HOST} could not be reached from this service`);
  assert.deepEqual(audit[0], { ...audit[0], decision: "failed", vendorStatus: 502 });
});

test("OData dates become days", () => {
  const ms = Date.parse("2026-10-29T00:00:00Z");
  assert.equal(day(`/Date(${ms})/`), "2026-10-29");
  assert.equal(day(`/Date(${ms}+0000)/`), "2026-10-29");
  assert.equal(day(null), null);
});
