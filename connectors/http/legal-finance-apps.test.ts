import assert from "node:assert/strict";
import { randomBytes } from "node:crypto";
import { test } from "node:test";
import type { McpServer } from "@modelcontextprotocol/server";
import type { ZodType } from "zod";
import type { ToolContext } from "../../packages/connector-kit/src/common.ts";
import { loadConnectors, type Connector } from "../../packages/connector-kit/src/connector.ts";
import type { Limits, VendorConnection } from "../../packages/connector-kit/src/types.ts";
import { httpTools, httpVendor } from "./index.ts";
import { builtinConnectors } from "../../packages/server/src/builtin.ts";
import { CredentialKey } from "../../packages/server/src/secrets.ts";
import { assertToolsImplemented, createService } from "../../packages/server/src/service.ts";
import { fakeStore, ORG } from "../../packages/server/test/helpers.ts";

// Ironclad and OneStream, described only in their connector.json: each loads, is offered, and
// runs a representative tool and its sign-in check against a fake app.

const connectors = loadConnectors();
const app = (id: string) => connectors.find((c) => c.id === id)!;

type Seen = { method: string; url: string; headers: Record<string, string>; body: string };

// Answers a token request with a token and every other call with `answer`.
function fakeApp(answer: unknown) {
  const seen: Seen[] = [];
  const fetchImpl = (async (input: string | URL, init?: RequestInit) => {
    const url = String(input);
    seen.push({
      method: init?.method ?? "GET",
      url,
      headers: Object.fromEntries(new Headers(init?.headers)),
      body: init?.body === undefined ? "" : String(init.body),
    });
    if (/\/oauth\/token$/.test(new URL(url).pathname))
      return Response.json({ access_token: "minted", expires_in: 3600 });
    return Response.json(answer);
  }) as typeof fetch;
  return { seen, fetchImpl };
}

async function run(
  connector: Connector,
  conn: VendorConnection,
  tool: string,
  args: Record<string, unknown>,
  fake: ReturnType<typeof fakeApp>,
  limits: Limits = {},
) {
  const vendor = httpVendor(connector, fake.fetchImpl);
  const handlers = new Map<string, (a: Record<string, unknown>) => Promise<unknown>>();
  const recorder = {
    registerTool: (name: string, c: { inputSchema: ZodType }, h: (a: unknown) => unknown) =>
      handlers.set(name, async (a) => h(c.inputSchema.parse(a))),
  } as unknown as McpServer;
  const ctx = {
    limits,
    config: conn.config,
    call: (r) => vendor.call(conn, r),
    audit: async () => {},
  } as ToolContext;
  const toolset = httpTools(connector.http!);
  toolset.register(recorder, ctx, new Set(toolset.names));
  const result = (await handlers.get(tool)!(args)) as {
    isError?: boolean;
    content: { text: string }[];
  };
  return { isError: result.isError ?? false, text: result.content[0].text };
}

const ironcladConn = {
  config: { host: "eu1.ironcladapp.com", client_id: "cid", user_email: "legal-ops@example.com" },
  secret: { client_secret: "cs" },
};
const onestreamConn = {
  config: {
    host: "acme.onestreamcloud.com",
    web_server_url: "https://acme.onestreamcloud.com/OneStreamWeb",
  },
  secret: { token: "pat" },
};

test("Ironclad and OneStream load, are offered and run on the generic module", () => {
  for (const id of ["ironclad", "onestream"]) {
    const connector = app(id);
    assert.ok(connector?.http, id);
    assert.ok(builtinConnectors.some((m) => m.id === id), `${id} is offered`);
    assertToolsImplemented([
      { connector, toolset: httpTools(connector.http!), vendor: httpVendor(connector) },
    ]);
  }
});

test("Ironclad: a token from the connection's region, then calls as the named user", async () => {
  let fake = fakeApp({ page: 0, pageSize: 20, count: 1, list: [{ id: "r1", type: "nda", name: "NDA", properties: {} }] });
  const out = await run(
    app("ironclad"),
    ironcladConn,
    "ironclad_list_records",
    { types: ["nda", "msa"], search: "Acme" },
    fake,
    { record_types: ["nda", "msa"] },
  );
  assert.equal(fake.seen[0].url, "https://eu1.ironcladapp.com/oauth/token");
  assert.match(fake.seen[0].body, /grant_type=client_credentials/);
  assert.equal(
    fake.seen[1].url,
    "https://eu1.ironcladapp.com/public/api/v1/records?types=nda%2Cmsa&search=Acme&sortField=agreementDate&sortDirection=DESC&page=0&pageSize=20",
  );
  assert.equal(fake.seen[1].headers.authorization, "Bearer minted");
  assert.equal(fake.seen[1].headers["x-as-user-email"], "legal-ops@example.com");
  const [record] = JSON.parse(out.text);
  assert.equal(record.name, "NDA");
  assert.equal(record.properties, undefined);

  // A record of a type outside the key's is refused after the read.
  fake = fakeApp({ id: "r2", type: "employment", name: "Offer" });
  const refused = await run(app("ironclad"), ironcladConn, "ironclad_get_record", { id: "IC-9" }, fake, {
    record_types: ["nda"],
  });
  assert.equal(refused.isError, true);

  // A signer's role name is one encoded path segment.
  fake = fakeApp({});
  await run(app("ironclad"), ironcladConn, "ironclad_remind_signer", { id: "IC-7", signer_role_name: "Counterparty Signer" }, fake);
  assert.equal(fake.seen[1].method, "POST");
  assert.equal(
    fake.seen[1].url,
    "https://eu1.ironcladapp.com/public/api/v1/workflows/IC-7/sign-status/signers/Counterparty%20Signer/remind",
  );

  // Only Ironclad's own hosts, for the token too.
  fake = fakeApp({});
  const elsewhere = await run(
    app("ironclad"),
    { ...ironcladConn, config: { ...ironcladConn.config, host: "ironclad.attacker.example" } },
    "ironclad_get_workflow",
    { id: "IC-1" },
    fake,
  );
  assert.equal(elsewhere.isError, true);
  assert.equal(fake.seen.length, 0);
});

test("OneStream: the web server address comes from the connection, never the agent", async () => {
  let fake = fakeApp({ ResultsTable: [{ Account: "Cash", Amount: 10 }] });
  const out = await run(
    app("onestream"),
    onestreamConn,
    "onestream_get_adapter_data",
    { application: "GolfStream", adapter: "Balance Sheet Recs" },
    fake,
    { adapters: ["Balance Sheet Recs"] },
  );
  assert.equal(fake.seen[0].method, "POST");
  assert.equal(
    fake.seen[0].url,
    "https://acme.onestreamcloud.com/OneStreamApi/api/DataProvider/GetAdoDataSetForAdapter?api-version=5.2.0",
  );
  assert.equal(fake.seen[0].headers.authorization, "Bearer pat");
  assert.deepEqual(JSON.parse(fake.seen[0].body), {
    BaseWebServerUrl: "https://acme.onestreamcloud.com/OneStreamWeb",
    ApplicationName: "GolfStream",
    AdapterName: "Balance Sheet Recs",
    ResultDataTableName: "ResultsTable",
  });
  assert.equal(JSON.parse(out.text).ResultsTable[0].Account, "Cash");

  fake = fakeApp({});
  await assert.rejects(
    run(app("onestream"), onestreamConn, "onestream_run_sequence", {
      application: "GolfStream",
      sequence: "Close",
      web_server_url: "https://evil.example/OneStreamWeb",
    }, fake),
  );
  const outside = await run(app("onestream"), onestreamConn, "onestream_run_sequence", { application: "GolfStream", sequence: "Clear Actuals" }, fake, {
    sequences: ["Close"],
  });
  assert.equal(outside.isError, true);
  assert.equal(fake.seen.length, 0);
});

test("connecting proves the sign-in: Ironclad's regional token, OneStream's logon", async () => {
  const key = new CredentialKey(randomBytes(32).toString("base64"));
  const fake = fakeApp({ count: 3, Message: "Logon succeeded." });
  const service = createService({
    apps: ["ironclad", "onestream"].map((id) => ({
      connector: app(id),
      vendor: httpVendor(app(id), fake.fetchImpl),
      toolset: httpTools(app(id).http!),
    })),
    store: fakeStore(key).store,
    key,
  });
  await service.connect({
    organizationId: ORG,
    connectorId: "ironclad",
    capabilities: ["read"],
    fields: { host: "eu1.ironcladapp.com", client_id: "cid", client_secret: "cs", user_email: "legal-ops@example.com" },
    connectedBy: "admin@x",
  });
  assert.deepEqual(
    fake.seen.map((s) => s.url),
    ["https://eu1.ironcladapp.com/oauth/token", "https://eu1.ironcladapp.com/public/api/v1/workflows?pageSize=1"],
  );
  fake.seen.length = 0;
  await service.connect({
    organizationId: ORG,
    connectorId: "onestream",
    capabilities: ["read"],
    fields: { host: "acme.onestreamcloud.com", web_server_url: "https://acme.onestreamcloud.com/OneStreamWeb", token: "pat" },
    connectedBy: "admin@x",
  });
  assert.equal(fake.seen[0].url, "https://acme.onestreamcloud.com/OneStreamApi/api/Authentication/Logon?api-version=7.2.0");
  assert.deepEqual(JSON.parse(fake.seen[0].body), {
    BaseWebServerUrl: "https://acme.onestreamcloud.com/OneStreamWeb",
  });
});
