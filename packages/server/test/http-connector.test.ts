import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createServer, type IncomingMessage, type ServerResponse } from "node:http";
import type { AddressInfo } from "node:net";
import { after, before, test } from "node:test";
import type { McpServer } from "@modelcontextprotocol/server";
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
  const recorder = {
    registerTool: (name: string, _c: unknown, h: never) => handlers.set(name, h),
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
