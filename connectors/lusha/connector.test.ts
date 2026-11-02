import assert from "node:assert/strict";
import { createServer, type IncomingMessage } from "node:http";
import type { AddressInfo } from "node:net";
import { after, before, test } from "node:test";
import type { McpServer } from "@modelcontextprotocol/server";
import type { ZodType } from "zod";
import { connectorSchema, loadConnectors } from "../../packages/connector-kit/src/connector.ts";
import type { ToolContext } from "../../packages/connector-kit/src/common.ts";
import { assertToolsImplemented } from "../../packages/server/src/service.ts";
import { httpTools, httpVendor } from "../http/index.ts";

// Lusha's connector.json, loaded as the server loads it and run against a local fake.

const seen: { method: string; url: string; headers: IncomingMessage["headers"] }[] = [];
let origin = "";
const server = createServer((req, res) => {
  seen.push({ method: req.method!, url: req.url!, headers: req.headers });
  res.writeHead(200, { "content-type": "application/json" });
  res.end(
    JSON.stringify({
      contact: {
        error: null,
        isCreditCharged: true,
        data: { personId: 7, fullName: "Ann Lee", linkedinSkills: ["x"] },
      },
    }),
  );
});
before(async () => {
  await new Promise<void>((done) => server.listen(0, "127.0.0.1", done));
  origin = `http://127.0.0.1:${(server.address() as AddressInfo).port}`;
});
after(() => server.close());

test("lusha loads and finds a contact with its key header", async () => {
  const [loaded] = loadConnectors(undefined, (id) => id === "lusha");
  assertToolsImplemented([
    { connector: loaded, toolset: httpTools(loaded.http!), vendor: httpVendor(loaded) },
  ]);
  const connector = connectorSchema.parse({
    ...loaded,
    http: { ...loaded.http, baseUrl: origin, allowInsecureHttp: true },
  });
  const vendor = httpVendor(connector);
  const conn = { config: { host: "x" }, secret: { api_key: "k1" } };
  const handlers = new Map<string, (a: unknown) => Promise<unknown>>();
  const recorder = {
    registerTool: (name: string, c: { inputSchema: ZodType }, h: (a: unknown) => unknown) =>
      handlers.set(name, async (a) => h(c.inputSchema.parse(a))),
  } as unknown as McpServer;
  const toolset = httpTools(connector.http!);
  toolset.register(
    recorder,
    {
      limits: {},
      config: conn.config,
      call: (r) => vendor.call(conn, r),
      audit: async () => {},
    } as ToolContext,
    new Set(toolset.names),
  );
  const result = (await handlers.get("lusha_find_contact")!({
    first_name: "Ann",
    last_name: "Lee",
    company_domain: "acme.com",
  })) as { content: { text: string }[] };
  assert.equal(seen[0].method, "GET");
  assert.equal(seen[0].url, "/v2/person?firstName=Ann&lastName=Lee&companyDomain=acme.com");
  assert.equal(seen[0].headers.api_key, "k1");
  const data = JSON.parse(result.content[0].text);
  assert.equal(data["contact.data.fullName"], "Ann Lee");
  assert.equal(data["contact.data.linkedinSkills"], undefined);
});
