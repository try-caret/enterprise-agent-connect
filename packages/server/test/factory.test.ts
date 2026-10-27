import assert from "node:assert/strict";
import { after, test } from "node:test";
import { createAgentConnect } from "../src/index.ts";
import { connectDatabase } from "../src/store.ts";

// The whole service as an embedder builds it. Needs TEST_DATABASE_URL (see store.test.ts).
const url = process.env.TEST_DATABASE_URL;
const SCHEMA = "eac_factory_test";
const TOKEN = "t".repeat(40);
const sql = url ? connectDatabase(url) : null;
after(async () => {
  await sql?.unsafe(`drop schema if exists ${SCHEMA} cascade`);
  await sql?.end();
});

test("createAgentConnect migrates, brands and serves", { skip: !sql }, async () => {
  await sql!.unsafe(`drop schema if exists ${SCHEMA} cascade`);
  const service = await createAgentConnect({
    database: sql!,
    schema: SCHEMA,
    encryptionKey: Buffer.alloc(32, 1).toString("base64"),
    adminToken: TOKEN,
    publicBaseUrl: "https://agents.example/",
    branding: { productName: "Acme Agents", resourcePrefix: "acme" },
    log: { log() {}, error() {} },
  });
  const get = (path: string) =>
    service.fetch(
      new Request(`http://svc/agent-connect/admin/v1${path}`, {
        headers: { authorization: `Bearer ${TOKEN}` },
      }),
    );
  assert.equal(
    ((await (await get("/branding")).json()) as { productName: string }).productName,
    "Acme Agents",
  );
  const connectors = JSON.stringify(await (await get("/connectors")).json());
  assert.doesNotMatch(connectors, /\{brand\./);
  assert.match(connectors, /Acme Agents/);
  const [row] = await sql!.unsafe(`select count(*)::int as n from ${SCHEMA}.schema_migrations`);
  assert.ok(row.n > 0);
  const refused = await service.fetch(
    new Request("http://svc/agent-connect/mcp", { method: "POST" }),
  );
  assert.equal(refused.status, 401);
  assert.match(refused.headers.get("www-authenticate") ?? "", /realm="acme-agent-connect"/);
  await service.stop();
  // A database it was handed stays open.
  await sql!`select 1`;
});

test("createAgentConnect refuses a connector without code", { skip: !sql }, async () => {
  await assert.rejects(
    createAgentConnect({
      database: sql!,
      schema: SCHEMA,
      migrate: false,
      encryptionKey: Buffer.alloc(32, 1).toString("base64"),
      adminToken: TOKEN,
      branding: { productName: "Acme Agents", resourcePrefix: "acme" },
      connectors: [],
    }),
    /has no code/,
  );
});

test("createAgentConnect refuses a connector folder with no connector", { skip: !sql }, async () => {
  await assert.rejects(
    createAgentConnect({
      database: sql!,
      schema: SCHEMA,
      migrate: false,
      encryptionKey: Buffer.alloc(32, 1).toString("base64"),
      adminToken: TOKEN,
      branding: { productName: "Acme Agents", resourcePrefix: "acme" },
      extraConnectorDirs: [new URL("../../../examples/internal-app", import.meta.url).pathname],
    }),
    /No connector\.json/,
  );
});

test("an app described only in its connector.json runs without code", { skip: !sql }, async () => {
  const service = await createAgentConnect({
    database: sql!,
    schema: SCHEMA,
    encryptionKey: Buffer.alloc(32, 1).toString("base64"),
    adminToken: TOKEN,
    branding: { productName: "Acme Agents", resourcePrefix: "acme" },
    extraConnectorDirs: [
      new URL("../../../examples/internal-app/connectors", import.meta.url).pathname,
    ],
    log: { log() {}, error() {} },
  });
  assert.ok(service.connectors.includes("orders"));
  await service.stop();
});
