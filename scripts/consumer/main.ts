// Run by scripts/consumer-test.sh against the packed packages: embeds the service the ways the
// README shows, then drives it as an admin and as an agent. DATABASE_URL: a scratch Postgres.
import assert from "node:assert/strict";
import { createServer, type Server } from "node:http";
import type { AddressInfo } from "node:net";
import { getRequestListener, serve } from "@hono/node-server";
import express from "express";
import { Hono } from "hono";
import postgres from "postgres";
import { builtinConnectors, createAgentConnect } from "@arbium/enterprise-agent-connect";
import { createClient } from "@arbium/enterprise-agent-connect-client";
import { inventory, orders } from "./apps.js";

const url = (server: Server) => `http://127.0.0.1:${(server.address() as AddressInfo).port}`;
const listening = (server: Server) =>
  new Promise<Server>((resolve) => server.once("listening", () => resolve(server)));

// The company's own API the two apps call.
const stock: Record<string, number> = { A1: 7, B2: 3 };
const ordersById: Record<string, { id: string; customer: string }> = {
  "1001": { id: "1001", customer: "acme" },
  "1002": { id: "1002", customer: "globex" },
};
const api = createServer((req, res) => {
  const send = (status: number, data: unknown) =>
    res.writeHead(status, { "content-type": "application/json" }).end(JSON.stringify(data));
  if (req.headers["x-api-key"] !== "example-key") return send(401, { error: "bad key" });
  const [, kind, id] = req.url!.split("/").slice(1);
  if (kind === "health") return send(200, { name: "Example" });
  const found = kind === "orders" ? ordersById[id] : kind === "stock" ? { sku: id, count: stock[id] } : null;
  return found ? send(200, found) : send(404, { error: "not found" });
}).listen(0, "127.0.0.1");
await listening(api);
const apiUrl = `${url(api)}/api`;

const TOKEN = "t".repeat(40);
const schema = `eac_consumer_${process.pid}`;
const sql = postgres(process.env.DATABASE_URL!, { onnotice: () => {} });

// README: embedding.
const eac = await createAgentConnect({
  database: sql,
  schema,
  encryptionKey: Buffer.alloc(32, 7).toString("base64"),
  adminToken: TOKEN,
  connectors: [...builtinConnectors, orders(apiUrl), inventory(apiUrl)],
});
eac.start();

// README: Hono.
const app = new Hono();
app.all("/agent-connect/*", (c) => eac.fetch(c.req.raw));
const hono = serve({ fetch: app.fetch, port: 0, hostname: "127.0.0.1" }) as Server;
await listening(hono);

// README: Express.
const ex = express();
ex.all("/agent-connect/*splat", getRequestListener(eac.fetch));
const viaExpress = ex.listen(0, "127.0.0.1");
await listening(viaExpress);

// README: a Next.js route handler (app/agent-connect/[...path]/route.ts).
const handler = (req: Request) => eac.fetch(req);

for (const base of [url(hono), url(viaExpress)])
  assert.equal((await fetch(`${base}/agent-connect/healthz`)).status, 200, base);
assert.equal((await handler(new Request("http://next/agent-connect/healthz"))).status, 200);

try {
  // README: the typed admin client.
  const client = createClient({ baseUrl: url(hono), token: TOKEN });
  const org = await client.organizations.create("Example Org");
  for (const app of ["orders", "inventory"])
    await client.connect(org.id, app, {
      capabilities: ["read"],
      fields: { api_key: "example-key" },
      connectedBy: "Example Admin",
    });
  const { agent, key } = await client.createAgent(org.id, { name: "demo", createdBy: "Example Admin" });
  await client.setGrant(org.id, agent.id, "orders", ["read"], { customers: ["acme"] }, "Example Admin");
  await client.setGrant(org.id, agent.id, "inventory", ["read"], { skus: ["A1"] }, "Example Admin");

  // README: any MCP client, here over Express.
  const call = async (name: string, args: Record<string, string>) => {
    const res = await fetch(`${url(viaExpress)}/agent-connect/mcp`, {
      method: "POST",
      headers: {
        authorization: `Bearer ${key}`,
        "content-type": "application/json",
        accept: "application/json, text/event-stream",
      },
      body: JSON.stringify({ jsonrpc: "2.0", id: 1, method: "tools/call", params: { name, arguments: args } }),
    });
    const text = await res.text();
    const body = JSON.parse(text.startsWith("{") ? text : text.match(/^data: (.*)$/m)![1]);
    return body.result as { isError?: boolean; content: { text: string }[] };
  };
  assert.match((await call("orders_get", { id: "1001" })).content[0].text, /acme/);
  assert.equal((await call("orders_get", { id: "1002" })).isError, true);
  assert.match((await call("inventory_get_stock", { sku: "A1" })).content[0].text, /"count": 7/);
  assert.equal((await call("inventory_get_stock", { sku: "B2" })).isError, true);

  // README: the client in the same process, through eac.fetch with no network hop.
  const inProcess = createClient({
    baseUrl: "http://agent-connect",
    token: TOKEN,
    fetch: (input, init) => eac.fetch(new Request(input, init)),
  });
  const audit = await inProcess.audit(org.id, { agent: agent.id });
  const seen = audit.map((r) => `${r.tool} ${r.decision}`).sort();
  assert.deepEqual(seen, [
    "inventory_get_stock allowed",
    "inventory_get_stock denied",
    "orders_get allowed",
    "orders_get denied",
  ]);
  console.log("consumer test passed");
} finally {
  await eac.stop();
  for (const s of [api, hono, viaExpress]) s.close();
  await sql.unsafe(`drop schema if exists ${schema} cascade`);
  await sql.end();
}
