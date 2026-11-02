import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { readFileSync } from "node:fs";
import { createServer, type IncomingMessage, type ServerResponse } from "node:http";
import type { AddressInfo } from "node:net";
import { fileURLToPath } from "node:url";
import { after, before, test } from "node:test";
import type { McpServer } from "@modelcontextprotocol/server";
import type { ZodType } from "zod";
import { connectorSchema, type Connector } from "../../connector-kit/src/connector.ts";
import type { ToolContext } from "../../connector-kit/src/common.ts";
import { assertToolsImplemented } from "../src/service.ts";
import {
  httpTools,
  httpVendor,
  MAX_RESPONSE_BYTES,
  MAX_TARGET_BYTES,
} from "../../../connectors/http/index.ts";

// The generic http module, run against a local server standing in for an internal app.

const example = JSON.parse(
  readFileSync(
    new URL("../../../examples/internal-app/connectors/orders/connector.json", import.meta.url),
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

const ok = (data: unknown) => (_req: IncomingMessage, res: ServerResponse) => {
  res.writeHead(200, { "content-type": "application/json" });
  res.end(JSON.stringify(data));
};

function orders(http: Record<string, unknown> = {}, extra: Record<string, unknown> = {}) {
  return connectorSchema.parse({
    ...example,
    ...extra,
    http: { ...example.http, baseUrl: `${origin}/api`, ...http },
  });
}

// Registers the tools on a server that only records, so a test can call them directly.
async function run(
  connector: Connector,
  tool: string,
  args: Record<string, unknown>,
  {
    limits = {},
    secret = { api_key: "k1" } as Record<string, string>,
    config = {} as Record<string, string>,
  } = {},
) {
  const vendor = httpVendor(connector);
  const handlers = new Map<string, (a: Record<string, unknown>) => Promise<unknown>>();
  // Parses the arguments as the MCP server does before calling the tool.
  const recorder = {
    registerTool: (name: string, c: { inputSchema: ZodType }, h: (a: unknown) => unknown) =>
      handlers.set(name, async (a) => h(c.inputSchema.parse(a))),
  } as unknown as McpServer;
  const audit: { decision: string; reason: string | null; vendorStatus: number | null }[] = [];
  const conn = { config: { host: "x", ...config }, secret };
  const ctx = {
    limits,
    config: conn.config,
    call: (r) => vendor.call(conn, r),
    audit: async (e) => void audit.push(e),
  } as ToolContext;
  const toolset = httpTools(connector.http!);
  toolset.register(recorder, ctx, new Set(toolset.names));
  const result = (await handlers.get(tool)!(args)) as {
    isError?: boolean;
    content: { text: string }[];
  };
  return { result, text: result.content[0].text, audit };
}

test("the example connector loads and runs on the generic module", () => {
  const connector = orders();
  assertToolsImplemented([
    { connector, toolset: httpTools(connector.http!), vendor: httpVendor(connector) },
  ]);
});

test("a DELETE tool needs a destructive capability", () => {
  const connector = orders();
  const soft = {
    ...connector,
    capabilities: connector.capabilities.map((c) => ({ ...c, destructive: false })),
  };
  assert.throws(
    () =>
      assertToolsImplemented([
        { connector: soft, toolset: httpTools(soft.http!), vendor: httpVendor(soft) },
      ]),
    /capability cancel has orders_cancel/,
  );
});

test("http definitions are validated", () => {
  assert.throws(() => orders({ baseUrl: `${origin}/api?x=1` }), /No query/);
  assert.throws(() => orders({ baseUrl: "ftp://host/api" }));
  // The API key must be a secret field.
  assert.throws(
    () => orders({}, { fields: [{ id: "api_key", label: "Key", secret: false }] }),
    /auth names its fields/,
  );
  assert.throws(
    () =>
      orders(
        {},
        {
          fields: [
            { id: "api_key", label: "Key", secret: true },
            { id: "api_key", label: "Key again", secret: false },
          ],
        },
      ),
    /Field ids must be unique/,
  );
  const tool = example.http.tools[1];
  assert.throws(
    () => orders({ tools: [{ ...tool, params: {} }] }),
    /Each \{param\} in a tool's path/,
  );
  assert.throws(
    () =>
      orders({
        tools: [
          { ...tool, params: { id: { ...tool.params.id, limit: "nope" } } },
          ...example.http.tools.slice(2),
        ],
      }),
    /limit names a list limit/,
  );
});

test("path params are encoded and cannot climb out of the path", async () => {
  respond = ok({ id: "x" });
  seen.length = 0;
  await run(orders(), "orders_get", { id: "a/b?c=1#d" });
  assert.equal(seen[0].url, "/api/orders/a%2Fb%3Fc%3D1%23d");
  for (const id of ["..", "."]) {
    const { result, audit } = await run(orders(), "orders_get", { id });
    assert.equal(result.isError, true);
    assert.equal(audit[0].decision, "denied");
  }
  assert.equal(seen.length, 1);
});

test("a limited param must be one of the agent's values, and may not be left out", async () => {
  respond = ok([]);
  seen.length = 0;
  const limits = { customers: ["acme"] };
  assert.equal(
    (await run(orders(), "orders_list", { customer: "acme" }, { limits })).result.isError,
    undefined,
  );
  assert.equal(seen[0].url, "/api/orders?customer=acme");
  for (const args of [{ customer: "globex" }, {}]) {
    const { result, audit } = await run(orders(), "orders_list", args, { limits });
    assert.equal(result.isError, true);
    assert.equal(audit[0].decision, "denied");
  }
  assert.equal(seen.length, 1);
  // Without the limit, leaving it out lists everything.
  await run(orders(), "orders_list", {});
  assert.equal(seen[1].url, "/api/orders");
});

test("an answer limit refuses an answer outside the agent's values, failing closed", async () => {
  const limits = { customers: ["acme"] };
  const get = async (data: unknown) => {
    respond = ok(data);
    return run(orders(), "orders_get", { id: "1" }, { limits });
  };
  const allowed = await get({ id: "1001", customer: "acme" });
  assert.equal(allowed.result.isError, undefined);
  assert.match(allowed.text, /acme/);
  for (const data of [
    { id: "1002", customer: "globex" },
    { id: "1003" },
    { id: "1004", customer: { name: "acme" } },
    [{ customer: "acme" }, { customer: "globex" }],
    "acme",
  ]) {
    const { result, text, audit } = await get(data);
    assert.equal(result.isError, true);
    assert.doesNotMatch(text, /globex|1003|1004/);
    assert.equal(audit[0].decision, "denied");
    assert.equal(audit[0].vendorStatus, 200);
  }
  // Without values for the limit, the answer is passed on as is.
  respond = ok({ id: "1002", customer: "globex" });
  assert.equal((await run(orders(), "orders_get", { id: "1" })).result.isError, undefined);
});

test("plain http needs allowInsecureHttp, and answer limits are for reads only", () => {
  assert.throws(() => orders({ allowInsecureHttp: false }), /must be https/);
  orders({ baseUrl: "https://orders.example.com/api", allowInsecureHttp: false });
  assert.throws(
    () =>
      orders({
        allowInsecureHttp: false,
        baseUrl: "https://orders.example.com/api",
        auth: { type: "oauth2", tokenUrl: "http://idp/token", clientId: "id", clientSecret: "s" },
      }),
    /must be https/,
  );
  const note = example.http.tools[2];
  assert.throws(
    () =>
      orders({
        tools: [
          ...example.http.tools.slice(0, 2),
          { ...note, answerLimits: [{ field: "customer", limit: "customers" }] },
          example.http.tools[3],
        ],
      }),
    /answerLimits are for GET tools only/,
  );
});

test("body params are sent as JSON and the agent's headers are never forwarded", async () => {
  respond = ok({ ok: true });
  seen.length = 0;
  await run(orders(), "orders_add_note", { id: "1001", text: "call back" });
  assert.equal(seen[0].method, "POST");
  assert.deepEqual(JSON.parse(seen[0].body), { text: "call back" });
  assert.equal(seen[0].headers["x-api-key"], "k1");
  assert.equal(seen[0].headers.cookie, undefined);
});

test("each auth type signs the call", async () => {
  respond = ok({});
  seen.length = 0;
  const fields = (ids: [string, boolean][]) => ({
    fields: ids.map(([id, secret]) => ({ id, label: id, secret })),
  });
  await run(
    orders({ auth: { type: "bearer", token: "tok" } }, fields([["tok", true]])),
    "orders_get",
    { id: "1" },
    { secret: { tok: "t1" } },
  );
  assert.equal(seen[0].headers.authorization, "Bearer t1");
  await run(
    orders(
      { auth: { type: "basic", username: "user", password: "pass" } },
      fields([
        ["user", false],
        ["pass", true],
      ]),
    ),
    "orders_get",
    { id: "1" },
    { secret: { pass: "p" }, config: { user: "u" } },
  );
  assert.equal(seen[1].headers.authorization, `Basic ${Buffer.from("u:p").toString("base64")}`);

  // OAuth client credentials: a token from tokenUrl, then the call with it. The schema takes
  // only an https tokenUrl, so this one is built past it.
  respond = (req, res) =>
    req.url === "/token"
      ? ok({ access_token: "minted", expires_in: 3600 })(req, res)
      : ok({ id: "1" })(req, res);
  const auth = {
    type: "oauth2",
    tokenUrl: "https://login.example/token",
    clientId: "cid",
    clientSecret: "csecret",
  };
  const oauth = orders(
    { auth },
    fields([
      ["cid", false],
      ["csecret", true],
    ]),
  );
  oauth.http!.auth = {
    type: "oauth2",
    tokenUrl: `${origin}/token`,
    clientId: "cid",
    clientSecret: "csecret",
  };
  await run(oauth, "orders_get", { id: "1" }, { secret: { csecret: "s" }, config: { cid: "c" } });
  const token = seen.find((s) => s.url === "/token")!;
  assert.equal(new URLSearchParams(token.body).get("client_secret"), "s");
  assert.equal(seen.at(-1)!.headers.authorization, "Bearer minted");
});

test("redirects, large answers and binary answers are refused", async () => {
  const cases: [(req: IncomingMessage, res: ServerResponse) => void, RegExp][] = [
    [
      (_q, res) => (res.writeHead(302, { location: "http://169.254.169.254/" }), res.end()),
      /redirect/,
    ],
    [
      (_q, res) => (
        res.writeHead(200, { "content-type": "text/plain" }),
        res.end("x".repeat(MAX_RESPONSE_BYTES + 1))
      ),
      /larger than/,
    ],
    [
      (_q, res) => (res.writeHead(200, { "content-type": "image/png" }), res.end("png")),
      /only JSON or text/,
    ],
  ];
  for (const [handler, why] of cases) {
    respond = handler;
    const { result, text, audit } = await run(orders(), "orders_get", { id: "1" });
    assert.equal(result.isError, true);
    assert.match(text, why);
    assert.equal(audit[0].decision, "failed");
  }
});

test("a refusal from the app reaches the agent without the credentials", async () => {
  respond = (_q, res) => (
    res.writeHead(403, { "content-type": "application/json" }),
    res.end('{"error":"no access"}')
  );
  const { text, audit } = await run(
    orders(),
    "orders_get",
    { id: "1" },
    { secret: { api_key: "secret-value" } },
  );
  assert.match(text, /403: no access/);
  assert.doesNotMatch(text + JSON.stringify(audit), /secret-value/);
});

test("an empty success answers the agent, and an over-long address is refused", async () => {
  respond = (_q, res) => (res.writeHead(204), res.end());
  assert.equal((await run(orders(), "orders_cancel", { id: "1002" })).result.isError, undefined);
  seen.length = 0;
  const { result, audit } = await run(orders(), "orders_list", {
    customer: "x".repeat(MAX_TARGET_BYTES),
  });
  assert.equal(result.isError, true);
  assert.equal(audit[0].decision, "denied");
  assert.equal(seen.length, 0);
});

// The example's tools plus these, so a test can try a feature on a tool of its own.
const withTools = (tools: unknown[], http: Record<string, unknown> = {}) =>
  orders({ tools: [...example.http.tools, ...tools], ...http });
const search = (extra: Record<string, unknown> = {}) => ({
  name: "orders_search",
  description: "Search orders.",
  method: "GET",
  path: "/orders",
  ...extra,
});

test("fixed headers are sent with every call, and cannot carry credentials", async () => {
  respond = ok({});
  seen.length = 0;
  const connector = orders({
    headers: { "X-Api-Version": "2022-11-28", Accept: "application/vnd.example+json" },
  });
  await run(connector, "orders_get", { id: "1" });
  assert.equal(seen[0].headers["x-api-version"], "2022-11-28");
  assert.equal(seen[0].headers.accept, "application/vnd.example+json");
  assert.equal(seen[0].headers["x-api-key"], "k1");
  for (const headers of [
    { Authorization: "Bearer x" },
    { "X-Api-Key": "k" },
    { "x-auth-token": "t" },
    { "Content-Type": "text/plain" },
    { "X-Version": "1\r\nX-Injected: 1" },
  ])
    assert.throws(() => orders({ headers }), /headers|Invalid/);
});

test("a header auth prefix and query auth sign the call", async () => {
  respond = ok({});
  seen.length = 0;
  await run(
    orders({
      auth: { type: "header", header: "Authorization", value: "api_key", prefix: "SSWS " },
    }),
    "orders_get",
    { id: "1" },
  );
  assert.equal(seen[0].headers.authorization, "SSWS k1");
  const query = orders({ auth: { type: "query", param: "api_key", value: "api_key" } });
  await run(query, "orders_list", { customer: "acme" });
  assert.equal(seen[1].url, "/api/orders?customer=acme&api_key=k1");
  assert.equal(seen[1].headers["x-api-key"], undefined);

  // The key stays out of what the agent and the audit see.
  respond = (_q, res) => (
    res.writeHead(400, { "content-type": "application/json" }),
    res.end('{"error":"bad"}')
  );
  const { text, audit } = await run(
    query,
    "orders_get",
    { id: "1" },
    { secret: { api_key: "secret-value" } },
  );
  assert.doesNotMatch(text + JSON.stringify(audit), /secret-value/);

  // A query key must be a secret field, and no tool param may be sent under its name.
  assert.throws(
    () =>
      orders(
        { auth: { type: "query", param: "key", value: "api_key" } },
        { fields: [{ id: "api_key", label: "Key", secret: false }] },
      ),
    /auth names its fields/,
  );
  assert.throws(
    () => orders({ auth: { type: "query", param: "customer", value: "api_key" } }),
    /query auth param/,
  );
});

test("array, enum and default params, and a param sent under another name", async () => {
  respond = ok([]);
  seen.length = 0;
  const connector = withTools([
    search({
      params: {
        status: {
          type: "array",
          items: "string",
          enum: ["open", "shipped"],
          description: "Statuses",
          in: "query",
          default: ["open"],
        },
        tags: {
          type: "array",
          items: "string",
          description: "Tags",
          in: "query",
          join: ",",
          required: false,
        },
        page_size: { type: "integer", description: "Page size", in: "query", default: 20 },
        page_token: {
          type: "string",
          description: "Cursor",
          in: "query",
          name: "pageToken",
          required: false,
        },
      },
    }),
  ]);
  await run(connector, "orders_search", {});
  assert.equal(seen[0].url, "/api/orders?status=open&page_size=20");
  await run(connector, "orders_search", {
    status: ["open", "shipped"],
    tags: ["a", "b"],
    page_size: 5,
    page_token: "c2",
  });
  assert.equal(
    seen[1].url,
    "/api/orders?status=open&status=shipped&tags=a%2Cb&page_size=5&pageToken=c2",
  );
  // Outside the enum, or the wrong type: refused before the tool runs.
  await assert.rejects(run(connector, "orders_search", { status: ["lost"] }));
  await assert.rejects(run(connector, "orders_search", { page_size: 2.5 }));
  assert.equal(seen.length, 2);

  const param = (p: Record<string, unknown>) => () =>
    withTools([search({ params: { x: { description: "X", in: "query", ...p } } })]);
  assert.throws(param({ type: "array" }), /item type/);
  assert.throws(param({ type: "string", items: "string" }), /item type/);
  assert.throws(param({ type: "integer", enum: ["a"] }), /enum and default/);
  assert.throws(param({ type: "string", enum: ["a"], default: "b" }), /enum and default/);
  assert.throws(param({ type: "array", items: "string", default: "a" }), /enum and default/);
  assert.throws(param({ type: "string", join: "," }), /join/);
  assert.throws(
    () =>
      withTools([
        {
          ...example.http.tools[1],
          name: "orders_get_default",
          params: { id: { type: "string", description: "Id", in: "path", default: "1" } },
        },
      ]),
    /no default/,
  );
  assert.throws(
    () =>
      withTools([
        search({
          params: {
            a: { type: "string", description: "A", in: "query", name: "q" },
            b: { type: "string", description: "B", in: "query", name: "q" },
          },
        }),
      ]),
    /same name/,
  );
});

test("a limited array param needs every value within the agent's values", async () => {
  respond = ok([]);
  seen.length = 0;
  const connector = withTools([
    search({
      params: {
        customer: {
          type: "array",
          items: "string",
          description: "Customers",
          in: "query",
          required: false,
          limit: "customers",
        },
      },
    }),
  ]);
  const limits = { customers: ["acme", "initech"] };
  const allowed = await run(
    connector,
    "orders_search",
    { customer: ["acme", "initech"] },
    { limits },
  );
  assert.equal(allowed.result.isError, undefined);
  for (const args of [{ customer: ["acme", "globex"] }, { customer: [] }, {}]) {
    const { result, audit } = await run(connector, "orders_search", args, { limits });
    assert.equal(result.isError, true);
    assert.equal(audit[0].decision, "denied");
  }
  assert.equal(seen.length, 1);
});

test("a result keeps only the fields named, and answer limits see the whole answer", async () => {
  const connector = withTools([
    search({
      result: { items: "data.orders", fields: ["id", "owner.name", "lines.sku"] },
      answerLimits: [{ field: "data.orders.customer", limit: "customers" }],
    }),
  ]);
  const answer = {
    data: {
      orders: [
        {
          id: "1",
          customer: "acme",
          secret_note: "x",
          owner: { name: "Ann", email: "a@x" },
          lines: [{ sku: "s1", price: 1 }, { sku: "s2" }],
        },
        { id: "2", customer: "acme" },
      ],
    },
    total: 2,
  };
  respond = ok(answer);
  const { result, text } = await run(
    connector,
    "orders_search",
    {},
    {
      limits: { customers: ["acme"] },
    },
  );
  assert.equal(result.isError, undefined);
  assert.deepEqual(JSON.parse(text), [
    { id: "1", "owner.name": "Ann", "lines.sku": ["s1", "s2"] },
    { id: "2" },
  ]);

  // customer is trimmed away, but still checked.
  respond = ok({ data: { orders: [{ id: "3", customer: "globex" }] } });
  const refused = await run(connector, "orders_search", {}, { limits: { customers: ["acme"] } });
  assert.equal(refused.result.isError, true);
  assert.doesNotMatch(refused.text, /globex|"3"/);

  // An app leaving an empty list out answers an empty list.
  respond = ok({ data: {} });
  assert.equal((await run(connector, "orders_search", {})).text, "[]");

  // Text where JSON was expected is refused rather than passed on untrimmed.
  respond = (_q, res) => (res.writeHead(200, { "content-type": "text/plain" }), res.end("hi"));
  assert.equal((await run(connector, "orders_search", {})).result.isError, true);
  assert.throws(() => withTools([search({ result: {} })]), /names items, fields or both/);
});

test("paging hands the agent the next cursor from the body, a header or the Link header", async () => {
  const cursor = { type: "string", description: "Cursor", in: "query", required: false };
  const page = { type: "integer", description: "Page", in: "query", required: false };
  const connector = withTools([
    search({
      name: "orders_page_body",
      params: { cursor },
      paging: { param: "cursor", body: "meta.next" },
      result: { items: "orders", fields: ["id"] },
    }),
    search({
      name: "orders_page_header",
      params: { cursor },
      paging: { param: "cursor", header: "X-Next-Cursor" },
    }),
    search({ name: "orders_page_link", params: { page }, paging: { param: "page", link: true } }),
  ]);
  seen.length = 0;
  respond = ok({ orders: [{ id: "1", customer: "acme" }], meta: { next: "c2" } });
  let { text } = await run(connector, "orders_page_body", { cursor: "c1" });
  assert.deepEqual(JSON.parse(text), { items: [{ id: "1" }], next: "c2" });
  assert.equal(seen[0].url, "/api/orders?cursor=c1");
  respond = ok({ orders: [], meta: {} });
  ({ text } = await run(connector, "orders_page_body", {}));
  assert.deepEqual(JSON.parse(text), { items: [], next: null });

  respond = (_q, res) => (
    res.writeHead(200, { "content-type": "application/json", "x-next-cursor": "h2" }),
    res.end("[]")
  );
  ({ text } = await run(connector, "orders_page_header", {}));
  assert.deepEqual(JSON.parse(text), { items: [], next: "h2" });

  // Only the cursor is taken from the Link address; the service never calls it.
  respond = (_q, res) => (
    res.writeHead(200, {
      "content-type": "application/json",
      link: '<https://elsewhere.example/orders?page=1>; rel="prev", <https://elsewhere.example/orders?page=3&x=1>; rel="next"',
    }),
    res.end("[]")
  );
  seen.length = 0;
  ({ text } = await run(connector, "orders_page_link", { page: 2 }));
  assert.deepEqual(JSON.parse(text), { items: [], next: 3 });
  assert.equal(seen.length, 1);
  respond = ok([]);
  ({ text } = await run(connector, "orders_page_link", {}));
  assert.deepEqual(JSON.parse(text), { items: [], next: null });

  for (const paging of [
    { param: "nope", body: "next" },
    { param: "cursor", body: "next", header: "X" },
  ])
    assert.throws(() => withTools([search({ params: { cursor }, paging })]));
});

test("the example app pages through its orders with the example connector", async () => {
  const free = createServer();
  await new Promise<void>((done) => free.listen(0, "127.0.0.1", done));
  const port = (free.address() as AddressInfo).port;
  free.close();
  const app = spawn(
    process.execPath,
    [fileURLToPath(new URL("../../../examples/internal-app/server.mjs", import.meta.url))],
    { env: { ...process.env, PORT: String(port), API_KEY: "k1" }, stdio: "ignore" },
  );
  try {
    const connector = connectorSchema.parse({
      ...example,
      http: { ...example.http, baseUrl: `http://localhost:${port}/api` },
    });
    for (let up = false; !up; await new Promise((r) => setTimeout(r, 50)))
      up = await fetch(`http://localhost:${port}/api/health`).then(
        () => true,
        () => false,
      );
    const page1 = JSON.parse((await run(connector, "orders_list", {})).text);
    assert.deepEqual(page1, {
      items: [
        { id: "1001", customer: "acme", status: "shipped" },
        { id: "1002", customer: "globex", status: "open" },
      ],
      next: "2",
    });
    const page2 = JSON.parse((await run(connector, "orders_list", { cursor: page1.next })).text);
    assert.deepEqual(page2, {
      items: [{ id: "1003", customer: "acme", status: "open" }],
      next: null,
    });
  } finally {
    app.kill();
  }
});
