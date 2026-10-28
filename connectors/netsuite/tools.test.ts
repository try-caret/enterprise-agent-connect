import assert from "node:assert/strict";
import { constants, createHmac, generateKeyPairSync, verify } from "node:crypto";
import { test } from "node:test";
import { netsuiteTools } from "./tools.ts";
import { netsuiteVendor, normalizeAccount, tbaHeader } from "./vendor.ts";
import { loadConnectors } from "../../packages/connector-kit/src/connector.ts";
import type { ToolContext } from "../../packages/connector-kit/src/common.ts";
import type { VendorConnection, VendorRequest } from "../../packages/connector-kit/src/types.ts";
import { assertToolsImplemented } from "../../packages/server/src/service.ts";

// The NetSuite tools against a fake NetSuite, reached through the real vendor client, so the
// sign-ins, hosts and headers are exercised too. Subsidiaries 1 and 2, locations 10 and 20.

const ORIGIN = "https://1234567-sb1.suitetalk.api.netsuite.com";
const RESTLETS = "https://1234567-sb1.restlets.api.netsuite.com";
const KEYS = generateKeyPairSync("rsa", { modulusLength: 3072 });
const PRIVATE_KEY = KEYS.privateKey.export({ type: "pkcs8", format: "pem" }).toString();

type Reply = { status?: number; body?: unknown; headers?: Record<string, string> };
type Sent = { url: URL; method: string; headers: Headers; body: any };

function fakeNetSuite(
  handlers: {
    suiteql?: (q: string, url: URL) => Reply;
    record?: (path: string, body: any) => Reply;
    restlet?: (body: any) => Reply;
    token?: () => Reply;
  } = {},
) {
  const sent: Sent[] = [];
  let tokens = 0;
  const fetchImpl = (async (input: string, init: RequestInit) => {
    const url = new URL(input);
    const headers = new Headers(init.headers);
    const raw = init.body;
    const body =
      raw instanceof URLSearchParams
        ? Object.fromEntries(raw)
        : typeof raw === "string"
          ? JSON.parse(raw)
          : undefined;
    sent.push({ url, method: init.method!, headers, body });
    let reply: Reply;
    if (url.pathname === "/services/rest/auth/oauth2/v1/token")
      reply = handlers.token?.() ?? {
        body: {
          access_token: `token-${++tokens}-${body.client_assertion.length}`,
          expires_in: "3600",
        },
      };
    else if (url.pathname === "/services/rest/query/v1/suiteql")
      reply = handlers.suiteql?.(body.q, url) ?? { body: { items: [{ ok: 1 }], hasMore: false } };
    else if (url.pathname === "/app/site/hosting/restlet.nl")
      reply = handlers.restlet?.(body) ?? { status: 404 };
    else reply = handlers.record?.(url.pathname, body) ?? { status: 404 };
    return new Response(reply.body === undefined ? null : JSON.stringify(reply.body), {
      status: reply.status ?? 200,
      headers: reply.headers,
    });
  }) as unknown as typeof fetch;
  return { sent, fetchImpl };
}

const oauth: VendorConnection = {
  config: { accountId: "1234567_SB1", auth: "oauth2" },
  secret: { client_id: "client-1", certificate_id: "cert-1", private_key: PRIVATE_KEY },
};

function setup(
  limits: Record<string, unknown> = {},
  handlers: Parameters<typeof fakeNetSuite>[0] = {},
) {
  const ns = fakeNetSuite(handlers);
  const vendor = netsuiteVendor(ns.fetchImpl);
  const audit: { tool: string; target: string | null; decision: string; reason: string | null }[] =
    [];
  const ctx = {
    limits,
    config: {},
    call: (req: VendorRequest) => vendor.call(oauth, req),
    forbidden: { netsuite_find_customers: "The role needs Customers (View)." },
    async audit(e) {
      audit.push(e);
    },
  } as ToolContext;
  const tools = new Map<string, (args: unknown) => Promise<any>>();
  const server = { registerTool: (name: string, _def: unknown, fn: never) => tools.set(name, fn) };
  netsuiteTools.register(server as never, ctx, new Set(netsuiteTools.names));
  const run = async (name: string, args: object) => {
    const r = await tools.get(name)!(args);
    const text = r.content[0].text as string;
    return { isError: Boolean(r.isError), text, body: r.isError ? null : JSON.parse(text) };
  };
  const queries = () =>
    ns.sent.filter((s) => s.url.pathname.endsWith("/suiteql")).map((s) => s.body.q as string);
  const apiCalls = () => ns.sent.filter((s) => !s.url.pathname.endsWith("/token"));
  return { run, sent: ns.sent, queries, apiCalls, audit, vendor };
}

const decode = (part: string) => JSON.parse(Buffer.from(part, "base64url").toString());

test("the connector's capabilities, tools and checks line up; nothing is destructive", () => {
  const connector = loadConnectors().find((c) => c.id === "netsuite")!;
  assertToolsImplemented([{ connector, toolset: netsuiteTools, vendor: netsuiteVendor() }]);
  assert.deepEqual(
    connector.capabilities.filter((c) => c.destructive),
    [],
  );
  assert.equal(connector.capabilities.find((c) => c.required)?.id, "read");
  const asked = (auth: string) =>
    connector.fields.filter((f) => !f.requires || f.requires.value === auth).map((f) => f.id);
  assert.deepEqual(asked("oauth2"), [
    "account_id",
    "auth",
    "client_id",
    "certificate_id",
    "private_key",
  ]);
  assert.deepEqual(asked("tba"), [
    "account_id",
    "auth",
    "client_id",
    "client_secret",
    "token_id",
    "token_secret",
  ]);
});

test("OAuth 2.0 signs a PS256 assertion with the certificate id and proves it with SuiteQL", async () => {
  const { vendor, sent } = setup();
  const conn = await vendor.connect({
    account_id: " 1234567-sb1 ",
    auth: "oauth2",
    client_id: "client-1",
    certificate_id: "cert-1",
    private_key: PRIVATE_KEY.replaceAll("\n", " "),
  });
  assert.deepEqual(conn.config, { accountId: "1234567_SB1", auth: "oauth2" });
  assert.equal(conn.accountLabel, "NetSuite 1234567_SB1");
  assert.deepEqual(Object.keys(conn.secret).sort(), ["certificate_id", "client_id", "private_key"]);

  const [token, query] = sent;
  assert.equal(token.url.href, `${ORIGIN}/services/rest/auth/oauth2/v1/token`);
  assert.equal(token.body.grant_type, "client_credentials");
  assert.equal(
    token.body.client_assertion_type,
    "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
  );
  const [h, p, s] = token.body.client_assertion.split(".");
  assert.deepEqual(decode(h), { alg: "PS256", typ: "JWT", kid: "cert-1" });
  const claims = decode(p);
  assert.equal(claims.iss, "client-1");
  assert.deepEqual(claims.scope, ["rest_webservices"]);
  assert.equal(claims.aud, token.url.href);
  assert.ok(claims.exp - claims.iat <= 3600);
  assert.ok(
    verify(
      "sha256",
      Buffer.from(`${h}.${p}`),
      {
        key: KEYS.publicKey,
        padding: constants.RSA_PKCS1_PSS_PADDING,
        saltLength: constants.RSA_PSS_SALTLEN_DIGEST,
      },
      Buffer.from(s, "base64url"),
    ),
  );

  assert.equal(query.url.origin, ORIGIN);
  assert.match(query.headers.get("authorization")!, /^Bearer token-1-/);
  assert.equal(query.headers.get("prefer"), "transient");
  assert.equal(query.body.q, "SELECT 1 AS ok FROM DUAL");
});

test("connect refuses a short key, missing fields, a bad account and a refused sign-in", async () => {
  const { vendor } = setup();
  const weak = generateKeyPairSync("rsa", { modulusLength: 2048 })
    .privateKey.export({ type: "pkcs8", format: "pem" })
    .toString();
  const base = { account_id: "1234567", auth: "oauth2", client_id: "c", certificate_id: "k" };
  await assert.rejects(vendor.connect({ ...base, private_key: weak }), /at least 3072 bits/);
  await assert.rejects(vendor.connect({ ...base, private_key: "nope" }), /whole private key/);
  await assert.rejects(
    vendor.connect({ ...base, private_key: "" }),
    /certificate ID and the private key/,
  );
  await assert.rejects(vendor.connect({ ...base, auth: "tba" }), /token ID and secret/);
  await assert.rejects(vendor.connect({ ...base, account_id: "acme corp" }), /account ID/);
  assert.equal(normalizeAccount("https://1234567-sb1.app.netsuite.com"), "1234567_SB1");

  const refused = setup({}, { token: () => ({ status: 400, body: { error: "invalid_grant" } }) });
  await assert.rejects(
    refused.vendor.connect({ ...base, private_key: PRIVATE_KEY }),
    /NetSuite refused the sign-in: invalid_grant/,
  );
});

test("token-based authentication signs each call with HMAC-SHA256 and the account as realm", async () => {
  const header = tbaHeader({
    method: "POST",
    url: `${ORIGIN}/services/rest/query/v1/suiteql?limit=5&offset=0`,
    realm: "1234567_SB1",
    consumerKey: "ck",
    consumerSecret: "cs",
    tokenId: "tk",
    tokenSecret: "ts",
    nonce: "n0nce",
    timestamp: 1700000000,
  });
  const base =
    "POST&https%3A%2F%2F1234567-sb1.suitetalk.api.netsuite.com%2Fservices%2Frest%2Fquery%2Fv1%2Fsuiteql&" +
    encodeURIComponent(
      "limit=5&oauth_consumer_key=ck&oauth_nonce=n0nce&oauth_signature_method=HMAC-SHA256" +
        "&oauth_timestamp=1700000000&oauth_token=tk&oauth_version=1.0&offset=0",
    );
  const signature = createHmac("sha256", "cs&ts").update(base).digest("base64");
  assert.ok(header.startsWith('OAuth realm="1234567_SB1", oauth_consumer_key="ck"'));
  assert.ok(header.includes(`oauth_signature="${encodeURIComponent(signature)}"`));

  const { vendor, sent } = setup();
  const conn = await vendor.connect({
    account_id: "1234567_sb1",
    auth: "tba",
    client_id: "ck",
    client_secret: "cs",
    token_id: "tk",
    token_secret: "ts",
  });
  assert.deepEqual(conn.config, { accountId: "1234567_SB1", auth: "tba" });
  assert.equal(sent.length, 1, "no token call");
  assert.match(sent[0].headers.get("authorization")!, /^OAuth realm="1234567_SB1", /);
});

test("the token is reused, minted again once on a 401, and saved searches get their own", async () => {
  let unauthorized = 1;
  const { vendor, sent } = setup(
    {},
    {
      suiteql: () => (unauthorized-- > 0 ? { status: 401, body: {} } : { body: { items: [] } }),
      restlet: () => ({ body: { columns: [], rows: [], pages: 0, total: 0 } }),
    },
  );
  const q = { method: "POST" as const, path: "/services/rest/query/v1/suiteql", body: { q: "x" } };
  assert.equal((await vendor.call(oauth, q)).status, 200);
  assert.equal((await vendor.call(oauth, q)).status, 200);
  await vendor.call(oauth, {
    method: "POST",
    path: "/app/site/hosting/restlet.nl?script=1&deploy=1",
    body: {},
  });
  const tokens = sent.filter((s) => s.url.pathname.endsWith("/token"));
  assert.equal(tokens.length, 3, "first, after the 401, and one for restlets");
  assert.deepEqual(decode(tokens[2].body.client_assertion.split(".")[1]).scope, ["restlets"]);
  assert.equal(sent.at(-1)!.url.origin, RESTLETS);
});

const CUSTOMERS = [
  { links: [], id: "5", number: "C5", name: "Acme", subsidiary_id: "1" },
  { links: [], id: "6", number: "C6", name: "Other", subsidiary_id: "3" },
];

test("finding customers keeps to the key's subsidiaries, quotes text and pages", async () => {
  const { run, queries } = setup(
    { subsidiaries: ["1", "2"] },
    {
      suiteql: (_q, url) => ({
        body: { items: CUSTOMERS, hasMore: url.searchParams.get("offset") === "0" },
      }),
    },
  );
  const first = await run("netsuite_find_customers", { name: "O'Brien", max_results: 2 });
  assert.equal(first.isError, false);
  // The row outside the key's subsidiaries is dropped even when NetSuite returns it.
  assert.deepEqual(first.body.customers, [
    { id: "5", number: "C5", name: "Acme", subsidiary_id: "1" },
  ]);
  assert.equal(first.body.next, "2:2");
  const [q] = queries();
  assert.match(q, /FROM customer e WHERE/);
  assert.match(q, /LIKE '%O''BRIEN%'/);
  assert.match(q, /e\.isinactive = 'F' AND e\.subsidiary IN \(1, 2\)/);

  const second = await run("netsuite_find_customers", { cursor: first.body.next, max_results: 25 });
  assert.equal(second.body.next, null);
  const bad = await run("netsuite_find_customers", { cursor: "3:2", max_results: 25 });
  assert.match(bad.text, /cursor exactly/);
  const badId = await run("netsuite_find_vendors", { id: "1 OR 1=1", max_results: 25 });
  assert.match(badId.text, /internal id/);
  assert.equal(queries().length, 2, "refusals never reach NetSuite");
});

test("a record type the role cannot read gets the permission hint", async () => {
  const { run, audit } = setup(
    {},
    {
      suiteql: () => ({
        status: 400,
        body: {
          title: "Bad Request",
          "o:errorDetails": [
            {
              detail:
                "Invalid search query. Search error occurred: Record 'customer' was not found.",
            },
          ],
        },
      }),
    },
  );
  const r = await run("netsuite_find_customers", { max_results: 25 });
  assert.equal(r.isError, true);
  assert.match(
    r.text,
    /NetSuite answered 400: Invalid search query.*The role needs Customers \(View\)\./,
  );
  assert.equal(audit[0].decision, "failed");
});

test("sales orders keep to subsidiaries and locations, with lines when asked", async () => {
  const orders = [
    { id: "100", number: "SO1", subsidiary_id: "1", location_id: "10" },
    { id: "101", number: "SO2", subsidiary_id: "1", location_id: "30" },
  ];
  const lines = [
    { transaction_id: "100", line: "1", item_id: "7", location_id: "10" },
    { transaction_id: "100", line: "2", item_id: "8", location_id: "20" },
    { transaction_id: "100", line: "3", item_id: "9", location_id: null },
  ];
  const { run, queries } = setup(
    { subsidiaries: ["1"], locations: ["10"] },
    {
      suiteql: (q) => ({
        body: { items: q.includes("mainline = 'F'") ? lines : orders, hasMore: false },
      }),
    },
  );
  const r = await run("netsuite_find_sales_orders", {
    customer: "5",
    from: "2026-01-01",
    open_only: true,
    include_lines: true,
    max_results: 50,
  });
  assert.equal(r.isError, false);
  assert.deepEqual(
    r.body.transactions.map((t: { id: string }) => t.id),
    ["100"],
  );
  assert.deepEqual(
    r.body.transactions[0].lines.map((l: { line: string }) => l.line),
    ["1", "3"],
  );
  assert.equal(r.body.lines_hidden, 1);
  const [header, detail] = queries();
  assert.match(header, /t\.type = 'SalesOrd'/);
  assert.match(header, /t\.entity = 5/);
  assert.match(header, /t\.trandate >= TO_DATE\('2026-01-01', 'YYYY-MM-DD'\)/);
  assert.match(header, /t\.status NOT IN \('C', 'G', 'H'\)/);
  assert.match(header, /tl\.subsidiary IN \(1\) AND tl\.location IN \(10\)/);
  assert.match(detail, /tl\.transaction IN \(100\)/);
  assert.match(detail, /-tl\.quantity AS quantity/);

  const bad = await run("netsuite_find_sales_orders", { from: "2026-02-31x", max_results: 5 });
  assert.match(bad.text, /YYYY-MM-DD/);
});

test("invoices answer amounts due and can be limited to overdue ones", async () => {
  const { run, queries } = setup({}, { suiteql: () => ({ body: { items: [], hasMore: false } }) });
  const r = await run("netsuite_find_invoices", {
    overdue_only: true,
    include_lines: false,
    max_results: 5,
  });
  assert.equal(r.isError, false);
  const [q] = queries();
  assert.match(q, /foreignamountremaining AS amount_due/);
  assert.match(q, /t\.type = 'CustInvc'.*t\.status = 'A' AND t\.duedate < SYSDATE/);
  const po = await run("netsuite_find_purchase_orders", { vendor: "9", max_results: 5 });
  assert.equal(po.isError, false);
  assert.match(queries()[1], /t\.type = 'PurchOrd'.*t\.entity = 9/);
  assert.doesNotMatch(queries()[1], /amount_due/);
});

test("item availability refuses a location outside the key and filters to its own", async () => {
  const rows = [
    { item_id: "7", location_id: "10", subsidiary_id: "1", available: "4" },
    { item_id: "7", location_id: "20", subsidiary_id: "1", available: "9" },
  ];
  const { run, queries } = setup(
    { locations: ["10"] },
    {
      suiteql: () => ({ body: { items: rows, hasMore: false } }),
    },
  );
  const refused = await run("netsuite_item_availability", {
    item: "7",
    location: "20",
    max_results: 50,
  });
  assert.match(refused.text, /may not use location 20/);
  assert.equal(queries().length, 0);
  const r = await run("netsuite_item_availability", { item: "WIDGET", max_results: 50 });
  assert.deepEqual(r.body.availability, [rows[0]]);
  assert.match(queries()[0], /FROM aggregateitemlocation a/);
  assert.match(queries()[0], /UPPER\(i\.itemid\) LIKE '%WIDGET%'.*a\.location IN \(10\)/);
});

test("creating a customer needs one of the key's subsidiaries and answers the new id", async () => {
  const { run, apiCalls } = setup(
    { subsidiaries: ["1"] },
    {
      record: () => ({
        status: 204,
        headers: { location: `${ORIGIN}/services/rest/record/v1/customer/812` },
      }),
    },
  );
  assert.match(
    (await run("netsuite_create_customer", { company_name: "Acme" })).text,
    /must name a subsidiary/,
  );
  assert.match(
    (await run("netsuite_create_customer", { company_name: "Acme", subsidiary: "2" })).text,
    /must name a subsidiary: one of 1/,
  );
  assert.equal(apiCalls().length, 0);
  const r = await run("netsuite_create_customer", {
    company_name: "Acme",
    subsidiary: "1",
    email: "a@acme.example",
  });
  assert.deepEqual(r.body, { id: "812" });
  const [post] = apiCalls();
  assert.equal(post.url.pathname, "/services/rest/record/v1/customer");
  assert.deepEqual(post.body, {
    isPerson: false,
    companyName: "Acme",
    email: "a@acme.example",
    subsidiary: { id: "1" },
  });
});

test("creating a sales order checks the customer's subsidiary and the location first", async () => {
  const { run, apiCalls, audit } = setup(
    { subsidiaries: ["1"], locations: ["10"] },
    {
      suiteql: (q) => ({
        body: {
          items: q.includes("e.id = 5")
            ? [{ id: "5", subsidiary_id: "1" }]
            : q.includes("e.id = 6")
              ? [{ id: "6", subsidiary_id: "3" }]
              : [],
        },
      }),
      record: () => ({
        status: 204,
        headers: { location: `${ORIGIN}/services/rest/record/v1/salesOrder/9001` },
      }),
    },
  );
  const lines = [{ item: "7", quantity: 2 }];
  assert.match(
    (await run("netsuite_create_sales_order", { customer: "5", lines })).text,
    /must name a location/,
  );
  assert.match(
    (await run("netsuite_create_sales_order", { customer: "5", lines, location: "20" })).text,
    /may not use location 20/,
  );
  // Outside the key's subsidiaries and not found read the same.
  for (const customer of ["6", "404"])
    assert.match(
      (await run("netsuite_create_sales_order", { customer, lines, location: "10" })).text,
      new RegExp(`may not use customer ${customer}`),
    );
  assert.ok(!apiCalls().some((s) => s.method === "POST" && s.url.pathname.includes("/record/")));

  const r = await run("netsuite_create_sales_order", {
    customer: "5",
    lines: [{ item: "7", quantity: 2, rate: 9.5 }],
    location: "10",
    po_number: "PO-77",
  });
  assert.deepEqual(r.body, { id: "9001" });
  const post = apiCalls().at(-1)!;
  assert.equal(post.url.pathname, "/services/rest/record/v1/salesOrder");
  assert.deepEqual(post.body, {
    entity: { id: "5" },
    subsidiary: { id: "1" },
    location: { id: "10" },
    otherRefNum: "PO-77",
    item: { items: [{ item: { id: "7" }, quantity: 2, rate: 9.5, location: { id: "10" } }] },
  });
  assert.equal(audit.at(-1)!.target, "salesOrder/9001 for customer 5");
});

const SEARCH = {
  columns: [
    { name: "tranid", join: null, label: "Number" },
    { name: "subsidiary", join: null, label: null },
    { name: "subsidiary", join: "customer", label: null },
  ],
  rows: [
    [
      { value: "SO1", text: null },
      { value: "1", text: "US" },
      { value: "1", text: "US" },
    ],
    [
      { value: "SO2", text: null },
      { value: "3", text: "UK" },
      { value: "3", text: "UK" },
    ],
  ],
  pages: 2,
  total: 4,
};

test("saved searches run only when listed on the key, filtered to its subsidiaries", async () => {
  const bodies: any[] = [];
  const { run, apiCalls } = setup(
    { saved_searches: ["customsearch_open"], subsidiaries: ["1"] },
    { restlet: (b) => (bodies.push(b), { body: SEARCH }) },
  );
  assert.match(
    (await run("netsuite_run_saved_search", { search: "customsearch_all", max_results: 100 })).text,
    /only these saved searches: customsearch_open/,
  );
  assert.equal(apiCalls().length, 0);
  const r = await run("netsuite_run_saved_search", {
    search: "customsearch_open",
    max_results: 50,
  });
  assert.equal(r.isError, false);
  assert.deepEqual(r.body, {
    total: 4,
    rows: [{ Number: "SO1", subsidiary: "US", "customer.subsidiary": "US" }],
    rows_hidden: 1,
    next: "50:50",
  });
  const call = apiCalls()[0];
  assert.equal(call.url.origin, RESTLETS);
  assert.equal(
    call.url.search,
    "?script=customscript_agent_saved_search&deploy=customdeploy_agent_saved_search",
  );
  assert.deepEqual(bodies[0], { search: "customsearch_open", page: 0, size: 50 });
  const last = await run("netsuite_run_saved_search", {
    search: "customsearch_open",
    cursor: "50:50",
    max_results: 100,
  });
  assert.equal(last.body.next, null);
  assert.equal(bodies[1].page, 1);
});

test("a saved search without the column a limit needs is refused, and no list means none", async () => {
  const noColumn = {
    ...SEARCH,
    columns: SEARCH.columns.slice(0, 1),
    rows: [[{ value: "SO1", text: null }]],
  };
  const limited = setup(
    { saved_searches: ["customsearch_open"], locations: ["10"] },
    { restlet: () => ({ body: noColumn }) },
  );
  const r = await limited.run("netsuite_run_saved_search", {
    search: "customsearch_open",
    max_results: 100,
  });
  assert.match(r.text, /no Location column/);
  assert.equal(limited.audit[0].decision, "denied");
  const none = setup();
  assert.match(
    (await none.run("netsuite_run_saved_search", { search: "customsearch_open", max_results: 100 }))
      .text,
    /no saved searches/,
  );
  const broken = setup(
    { saved_searches: ["customsearch_open"] },
    {
      restlet: () => ({
        status: 400,
        body: { error: { code: "SSS_MISSING_REQD_ARGUMENT", message: "id missing" } },
      }),
    },
  );
  assert.match(
    (
      await broken.run("netsuite_run_saved_search", {
        search: "customsearch_open",
        max_results: 100,
      })
    ).text,
    /NetSuite answered 400: id missing/,
  );
});

test("the account check proves the connection and discover lists subsidiaries", async () => {
  const { vendor, sent } = setup(
    {},
    {
      suiteql: (q) => ({
        body: { items: q.includes("subsidiary") ? [{ id: "1" }, { id: "2" }] : [{ ok: 1 }] },
      }),
    },
  );
  assert.deepEqual(await vendor.checks!.account(oauth, { expected: [] }), {
    values: ["1234567_SB1"],
  });
  const connector = loadConnectors().find((c) => c.id === "netsuite")!;
  const probe = connector.discover as { method: "POST"; path: string; body: object };
  const res = await vendor.call(oauth, {
    method: probe.method,
    path: probe.path,
    body: probe.body,
  });
  assert.deepEqual(res.data, { items: [{ id: "1" }, { id: "2" }] });
  assert.equal(sent.at(-1)!.headers.get("prefer"), "transient");
});
