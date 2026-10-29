import assert from "node:assert/strict";
import { after, test } from "node:test";
import { readFileSync } from "node:fs";
import { z } from "zod";
import { builtinConnectors, createAgentConnect, type ConnectorJson } from "../src/index.ts";
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

const base = {
  schema: SCHEMA,
  migrate: false,
  encryptionKey: Buffer.alloc(32, 1).toString("base64"),
  adminToken: TOKEN,
  branding: { productName: "Acme Agents", resourcePrefix: "acme" },
  log: { log() {}, error() {} },
};
const orders: ConnectorJson = JSON.parse(
  readFileSync(
    new URL("../../../examples/internal-app/connectors/orders/connector.json", import.meta.url),
    "utf8",
  ),
);
const { http: _, ...withoutHttp } = orders;
const notes: ConnectorJson = {
  ...withoutHttp,
  id: "notes",
  name: "Notes",
  capabilities: [
    {
      id: "read",
      label: "Read notes",
      description: "Read a note.",
      scopes: [],
      permissions: [],
      tools: ["notes_get"],
    },
  ],
  limits: [],
};
const notesModule = {
  id: "notes",
  connector: notes,
  create: () => ({
    vendor: {
      connect: async () => ({ config: {}, secret: {}, accountLabel: "Notes" }),
      call: async () => ({ status: 200, data: {} }),
    },
    toolset: {
      names: ["notes_get"],
      register(server: { registerTool: Function }) {
        server.registerTool("notes_get", { inputSchema: z.object({}) }, () => ({
          content: [],
        }));
      },
    },
  }),
};

test("connectors may be a subset of the built-in ones", { skip: !sql }, async () => {
  const service = await createAgentConnect({
    ...base,
    database: sql!,
    connectors: builtinConnectors.filter((m) => ["jira", "fedex"].includes(m.id)),
  });
  assert.deepEqual(service.connectors.sort(), ["fedex", "jira"]);
  await service.stop();
});

test(
  "apps given inline: a definition alone, or a module carrying its own",
  { skip: !sql },
  async () => {
    const service = await createAgentConnect({
      ...base,
      database: sql!,
      connectors: [...builtinConnectors, orders, notesModule],
    });
    assert.ok(service.connectors.includes("orders"));
    assert.ok(service.connectors.includes("notes"));
    assert.ok(service.connectors.includes("teams"));
    await service.stop();
  },
);

test("createAgentConnect refuses connectors it cannot run", { skip: !sql }, async () => {
  const refuses = (
    connectors: Parameters<typeof createAgentConnect>[0]["connectors"],
    why: RegExp,
  ) => assert.rejects(createAgentConnect({ ...base, database: sql!, connectors }), why);
  await refuses([], /No connectors/);
  await refuses([withoutHttp], /orders has no code/);
  await refuses([{ ...notesModule, connector: undefined }], /notes has no definition/);
  await refuses([{ ...notesModule, id: "other" }], /other carries the definition of notes/);
  await refuses([notesModule, notesModule], /Two connector modules have the id notes/);
  await refuses([orders, orders], /Two connectors have the id orders/);
  await refuses([{ ...orders, logo: "orders.svg" }], /Connector orders: [\s\S]*logo/);
});

test(
  "createAgentConnect refuses a connector folder with no connector",
  { skip: !sql },
  async () => {
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
  },
);

test("an app described only in its connector.json runs without code", { skip: !sql }, async () => {
  const service = await createAgentConnect({
    database: sql!,
    schema: SCHEMA,
    encryptionKey: Buffer.alloc(32, 1).toString("base64"),
    adminToken: TOKEN,
    branding: { productName: "Acme Agents", resourcePrefix: "acme" },
    extraConnectorDirs: [
      new URL("../../../examples/internal-app/connectors", import.meta.url).pathname,
      new URL("../../../examples/legacy-apps/connectors", import.meta.url).pathname,
      new URL("../../../examples/legacy-orders/connectors", import.meta.url).pathname,
    ],
    log: { log() {}, error() {} },
  });
  // http, soap, odata and sql blocks each run on their generic module.
  for (const id of ["orders", "hr-soap", "sap-sales-orders", "erp"])
    assert.ok(service.connectors.includes(id));
  await service.stop();
});
