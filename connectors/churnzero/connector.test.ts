import assert from "node:assert/strict";
import { createServer, type IncomingMessage } from "node:http";
import type { AddressInfo } from "node:net";
import { after, before, test } from "node:test";
import type { McpServer } from "@modelcontextprotocol/server";
import { connectorSchema, loadConnectors } from "../../packages/connector-kit/src/connector.ts";
import type { ToolContext } from "../../packages/connector-kit/src/common.ts";
import { assertToolsImplemented } from "../../packages/server/src/service.ts";
import odataModule, { odataTools, odataVendor } from "../odata/index.ts";

// ChurnZero's connector.json, loaded as the server loads it and run against a local fake.

const seen: { method: string; url: string; headers: IncomingMessage["headers"] }[] = [];
let origin = "";
const server = createServer((req, res) => {
  seen.push({ method: req.method!, url: req.url!, headers: req.headers });
  res.writeHead(200, { "content-type": "application/json" });
  res.end(JSON.stringify({ value: [{ Id: 9, AccountId: 42, EventDate: "2026-10-01T00:00:00Z" }] }));
});
before(async () => {
  await new Promise<void>((done) => server.listen(0, "127.0.0.1", done));
  origin = `http://127.0.0.1:${(server.address() as AddressInfo).port}`;
});
after(() => server.close());

test("churnzero loads and lists an account's usage events with basic auth", async () => {
  const [loaded] = loadConnectors(undefined, (id) => id === "churnzero");
  assertToolsImplemented([{ connector: loaded, ...odataModule.create({ connector: loaded } as never) }]);
  const connector = connectorSchema.parse({
    ...loaded,
    odata: { ...loaded.odata, serviceRoot: `${origin}/public/v1`, allowInsecureHttp: true },
  });
  const vendor = odataVendor(connector);
  const conn = { config: { host: "x", instance: "acme", username: "svc" }, secret: { api_key: "k1" } };
  const handlers = new Map<string, (a: Record<string, unknown>) => Promise<unknown>>();
  const recorder = {
    registerTool: (name: string, _c: unknown, h: never) => handlers.set(name, h),
  } as unknown as McpServer;
  const toolset = odataTools(connector.odata!);
  toolset.register(
    recorder,
    {
      limits: { accounts: ["42"] },
      config: conn.config,
      call: (r) => vendor.call(conn, r),
      audit: async () => {},
    } as ToolContext,
    new Set(toolset.names),
  );
  await handlers.get("churnzero_list_events")!({ from: "2026-09-01" });
  const url = new URL(seen[0].url, "http://x");
  assert.equal(seen[0].method, "GET");
  assert.equal(url.pathname, "/public/v1/Event");
  // Left out, the limited account filter becomes the agent's own accounts.
  assert.equal(url.searchParams.get("$filter"), "(AccountId eq 42) and EventDate ge 2026-09-01T00:00:00.000Z");
  assert.equal(url.searchParams.get("$orderby"), "EventDate desc");
  assert.equal(seen[0].headers.authorization, `Basic ${Buffer.from("svc:k1").toString("base64")}`);
});
