import assert from "node:assert/strict";
import { createServer, type IncomingMessage } from "node:http";
import type { AddressInfo } from "node:net";
import { after, before, test } from "node:test";
import type { McpServer } from "@modelcontextprotocol/server";
import type { ZodType } from "zod";
import {
  connectorSchema,
  loadConnectors,
  pickLabels,
  type Connector,
} from "../../packages/connector-kit/src/connector.ts";
import type { ToolContext } from "../../packages/connector-kit/src/common.ts";
import { httpTools, httpVendor } from "./index.ts";
import { builtinConnectors } from "../../packages/server/src/builtin.ts";
import { assertToolsImplemented } from "../../packages/server/src/service.ts";

// The built-in apps described only by a connector.json `http` block, each loaded as the server
// loads it and one of its tools run against a local server standing in for the app.

const httpApps = loadConnectors(undefined, (id) =>
  builtinConnectors.some((m) => m.id === id),
).filter((c) => c.http);

const seen: { method: string; url: string; headers: IncomingMessage["headers"]; body: string }[] =
  [];
let answer: unknown = {};
let answerType = "application/json";
let origin = "";
const server = createServer(async (req, res) => {
  let body = "";
  for await (const chunk of req) body += chunk;
  seen.push({ method: req.method!, url: req.url!, headers: req.headers, body });
  const token = req.url === "/token";
  res.writeHead(200, { "content-type": token ? "application/json" : answerType });
  res.end(
    token
      ? JSON.stringify({ access_token: "tok", expires_in: 3600 })
      : typeof answer === "string"
        ? answer
        : JSON.stringify(answer),
  );
});
before(async () => {
  await new Promise<void>((done) => server.listen(0, "127.0.0.1", done));
  origin = `http://127.0.0.1:${(server.address() as AddressInfo).port}`;
});
after(() => server.close());

const app = (id: string) => httpApps.find((c) => c.id === id)!;

// The app with its address (and token service) moved to the local server, keeping the path
// under it.
function local(c: Connector) {
  const path = c.http!.baseUrl.replace(/^https:\/\/[^/]*/, "").replace(/\/$/, "");
  const auth =
    c.http!.auth.type === "oauth2" ? { ...c.http!.auth, tokenUrl: `${origin}/token` } : c.http!.auth;
  return connectorSchema.parse({
    ...c,
    http: {
      ...c.http,
      baseUrl: `${origin}${path}`,
      allowedHosts: undefined,
      auth,
      allowInsecureHttp: true,
    },
  });
}

async function run(
  connector: Connector,
  tool: string,
  args: Record<string, unknown>,
  secret: Record<string, string>,
  config: Record<string, string> = {},
) {
  const vendor = httpVendor(connector);
  const handlers = new Map<string, (a: Record<string, unknown>) => Promise<unknown>>();
  const recorder = {
    registerTool: (name: string, c: { inputSchema: ZodType }, h: (a: unknown) => unknown) =>
      handlers.set(name, async (a) => h(c.inputSchema.parse(a))),
  } as unknown as McpServer;
  const conn = { config: { host: "x", ...config }, secret };
  const ctx = {
    limits: {},
    config: conn.config,
    call: (r) => vendor.call(conn, r),
    audit: async () => {},
  } as ToolContext;
  const toolset = httpTools(connector.http!);
  toolset.register(recorder, ctx, new Set(toolset.names));
  seen.length = 0;
  const result = (await handlers.get(tool)!(args)) as { content: { text: string }[] };
  const text = result.content[0].text;
  return { data: answerType === "application/json" ? JSON.parse(text) : text, call: seen.at(-1)! };
}

test("every built-in http app loads and runs on the generic module", () => {
  for (const id of ["automox", "qualys", "druva", "drata", "safebase", "adobe-admin-console"])
    assert.ok(httpApps.some((c) => c.id === id), `${id} is a built-in http app`);
  assertToolsImplemented(
    httpApps.map((connector) => ({
      connector,
      toolset: httpTools(connector.http!),
      vendor: httpVendor(connector),
    })),
  );
});

test("a probe reads an answer that is itself a list", () => {
  assert.deepEqual(pickLabels([{ id: 7, name: "IT" }], "[].id", "[].name"), { "7": "IT" });
});

test("Automox lists an organization's devices with the API key", async () => {
  answer = [{ id: 1, name: "laptop-1", pending_patches: 3, detail: { big: true } }];
  const { data, call } = await run(
    local(app("automox")),
    "automox_list_devices",
    { org_id: 42, pending: 1 },
    { api_key: "ak" },
  );
  assert.equal(call.method, "GET");
  assert.equal(call.url, "/api/servers?o=42&pending=1&limit=50&page=0");
  assert.equal(call.headers.authorization, "Bearer ak");
  assert.equal(data[0].pending_patches, 3);
  assert.equal(data[0].detail, undefined);
});

test("Drata lists a workspace's evidence a page at a time", async () => {
  answer = {
    data: [{ id: 5, name: "Access review", status: "READY" }],
    pagination: { cursor: "c2" },
  };
  const { data, call } = await run(
    local(app("drata")),
    "drata_list_evidence",
    { workspace_id: 3, statuses: ["EXPIRED", "NEEDS_ARTIFACT"] },
    { api_key: "dk" },
  );
  assert.equal(call.method, "GET");
  assert.equal(
    call.url,
    "/public/v2/workspaces/3/evidence?evidenceStatuses%5B%5D=EXPIRED&evidenceStatuses%5B%5D=NEEDS_ARTIFACT&size=50",
  );
  assert.equal(call.headers.authorization, "Bearer dk");
  assert.deepEqual(data.next, "c2");
  assert.equal(data.items[0].name, "Access review");
});

test("SafeBase searches the knowledge base with its key header", async () => {
  answer = { ok: true, result: { count: 1, items: [{ id: "k1", content: "Yes, AES-256." }] } };
  const { data, call } = await run(
    local(app("safebase")),
    "safebase_search_knowledge_base",
    { query: "encryption at rest", products: ["default", "api"] },
    { api_key: "sk" },
  );
  assert.equal(call.method, "GET");
  assert.equal(
    call.url,
    "/api/ext/v1/rest/kb/search?query=encryption+at+rest&products=default%2Capi&pageNumber=1&pageSize=20",
  );
  assert.equal(call.headers["x-sb-api-key"], "sk");
  assert.equal(data[0].content, "Yes, AES-256.");
});

test("Qualys lists detections with its sign-in and passes the XML on as text", async () => {
  answer = "<HOST_LIST_VM_DETECTION_OUTPUT/>";
  answerType = "text/xml";
  const { data, call } = await run(
    local(app("qualys")),
    "qualys_list_detections",
    { severities: "4-5" },
    { password: "pw" },
    { username: "agents" },
  );
  answerType = "application/json";
  assert.equal(call.method, "GET");
  assert.equal(
    call.url,
    "/api/2.0/fo/asset/host/vm/detection/?action=list&severities=4-5&show_results=0&truncation_limit=50",
  );
  assert.equal(call.headers.authorization, `Basic ${Buffer.from("agents:pw").toString("base64")}`);
  assert.equal(call.headers["x-requested-with"], "Enterprise Agent Connect");
  assert.equal(data, "<HOST_LIST_VM_DETECTION_OUTPUT/>");
});

test("Druva signs in with Basic client credentials and pages through devices", async () => {
  answer = { devices: [{ deviceID: 9, deviceName: "mac-9", userID: 4 }], nextPageToken: "" };
  const { data, call } = await run(
    local(app("druva")),
    "druva_list_devices",
    { platform: "darwin" },
    { secret_key: "s" },
    { client_id: "c" },
  );
  const token = seen.find((s) => s.url === "/token")!;
  assert.equal(token.headers.authorization, `Basic ${Buffer.from("c:s").toString("base64")}`);
  assert.match(token.body, /scope=read/);
  assert.equal(call.url, "/insync/endpoints/v2/devices?platformOS=darwin");
  assert.equal(call.headers.authorization, "Bearer tok");
  assert.deepEqual(data, { items: [{ deviceID: 9, deviceName: "mac-9", userID: 4 }], next: null });
});

test("Adobe reads a user with its token and the client id as API key", async () => {
  answer = { result: "success", user: { email: "ann@example.com", groups: ["Acrobat Pro"] } };
  const { data, call } = await run(
    local(app("adobe-admin-console")),
    "adobe_get_user",
    { org_id: "ABC123@AdobeOrg", email: "ann@example.com" },
    { client_secret: "s" },
    { client_id: "cid", organization: "ABC123@AdobeOrg" },
  );
  assert.equal(call.method, "GET");
  assert.equal(
    call.url,
    "/v2/usermanagement/organizations/ABC123%40AdobeOrg/users/ann%40example.com",
  );
  assert.equal(call.headers.authorization, "Bearer tok");
  assert.equal(call.headers["x-api-key"], "cid");
  assert.deepEqual(data.user.groups, ["Acrobat Pro"]);
});
