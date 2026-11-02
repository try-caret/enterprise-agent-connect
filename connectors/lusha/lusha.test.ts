import assert from "node:assert/strict";
import { test } from "node:test";
import type { McpServer } from "@modelcontextprotocol/server";
import type { ZodType } from "zod";
import { loadConnectors } from "../../packages/connector-kit/src/connector.ts";
import type { ToolContext } from "../../packages/connector-kit/src/common.ts";
import { ConnectError } from "../../packages/connector-kit/src/types.ts";
import { assertToolsImplemented } from "../../packages/server/src/service.ts";
import { lushaTools } from "./tools.ts";
import { lushaVendor } from "./vendor.ts";

// Lusha against a fake api.lusha.com, answering as its v3 search-and-enrich and usage do.

function fakeLusha(answers: Response[] = []) {
  const calls: { url: string; init: RequestInit }[] = [];
  const fetchImpl = (async (input: string | URL, init: RequestInit = {}) => {
    calls.push({ url: String(input), init });
    return answers.shift() ?? Response.json({ credits: { total: 10, used: 1, remaining: 9 } });
  }) as typeof fetch;
  return { vendor: lushaVendor(fetchImpl, async () => {}), calls };
}

function tools(vendor: ReturnType<typeof lushaVendor>, secret = { api_key: "k1" }) {
  type Result = { isError?: boolean; content: { text: string }[] };
  const handlers = new Map<string, (a: unknown) => Promise<Result>>();
  const recorder = {
    registerTool: (name: string, c: { inputSchema: ZodType }, h: (a: unknown) => unknown) =>
      handlers.set(name, async (a) => h(c.inputSchema.parse(a)) as never),
  } as unknown as McpServer;
  const ctx = {
    limits: {},
    config: {},
    call: (r) => vendor.call({ config: {}, secret }, r),
    audit: async () => {},
  } as ToolContext;
  lushaTools.register(recorder, ctx, new Set(lushaTools.names));
  return handlers;
}

test("lusha's capabilities name only its tools", () => {
  const [connector] = loadConnectors(undefined, (id) => id === "lusha");
  assertToolsImplemented([{ connector, toolset: lushaTools, vendor: lushaVendor() }]);
});

test("connecting checks the key against v3 usage in the api_key header", async () => {
  const { vendor, calls } = fakeLusha();
  const conn = await vendor.connect({ api_key: " key1234 " });
  assert.equal(calls[0].url, "https://api.lusha.com/v3/account/usage");
  assert.equal((calls[0].init.headers as Record<string, string>).api_key, "key1234");
  assert.equal(conn.accountLabel, "API key ending 1234");
  const refused = fakeLusha([Response.json({ message: "bad" }, { status: 401 })]);
  await assert.rejects(refused.vendor.connect({ api_key: "x" }), ConnectError);
});

test("a contact is one search-and-enrich of one person, trimmed", async () => {
  const { vendor, calls } = fakeLusha([
    Response.json({
      requestId: "r",
      results: [
        {
          id: "43",
          fullName: "Ann Lee",
          jobTitle: { title: "CFO", seniority: "C-Level" },
          emails: [{ email: "ann@acme.com", type: "work", confidence: "A+", credits: 1 }],
          phones: [],
          tags: [{ id: "t" }],
          profile: { skills: ["x"] },
        },
      ],
      billing: { creditsCharged: 2 },
    }),
  ]);
  const handlers = tools(vendor);
  const res = await handlers.get("lusha_find_contact")!({
    first_name: "Ann",
    last_name: "Lee",
    company_domain: "acme.com",
    reveal: ["emails"],
  });
  assert.equal(calls[0].url, "https://api.lusha.com/v3/contacts/search-and-enrich");
  assert.equal(calls[0].init.method, "POST");
  assert.equal((calls[0].init.headers as Record<string, string>).api_key, "k1");
  assert.deepEqual(JSON.parse(String(calls[0].init.body)), {
    contacts: [{ firstName: "Ann", lastName: "Lee", companyDomain: "acme.com" }],
    reveal: ["emails"],
  });
  const data = JSON.parse(res.content[0].text);
  assert.equal(data.name, "Ann Lee");
  assert.equal(data.credits, 2);
  assert.deepEqual(data.emails, [{ email: "ann@acme.com", type: "work", confidence: "A+" }]);
  assert.equal(data.profile, undefined);

  const vague = await handlers.get("lusha_find_contact")!({ first_name: "Ann" });
  assert.equal(vague.isError, true);
  assert.equal(calls.length, 1);
});

test("a company is one search-and-enrich of one company; no match says so", async () => {
  const { vendor, calls } = fakeLusha([
    Response.json({ results: [{ id: "16", name: "Lusha", domain: "lusha.com", tags: ["x"] }] }),
    Response.json({ results: [] }),
    Response.json({ statusCode: 402, message: "no credits" }, { status: 402 }),
  ]);
  const handlers = tools(vendor);
  const res = await handlers.get("lusha_find_company")!({ domain: "lusha.com" });
  assert.equal(calls[0].url, "https://api.lusha.com/v3/companies/search-and-enrich");
  assert.deepEqual(JSON.parse(String(calls[0].init.body)), { companies: [{ domain: "lusha.com" }] });
  const data = JSON.parse(res.content[0].text);
  assert.equal(data.name, "Lusha");
  assert.equal(data.tags, undefined);
  const none = await handlers.get("lusha_find_company")!({ name: "Nobody" });
  assert.equal(JSON.parse(none.content[0].text).found, false);
  const broke = await handlers.get("lusha_find_company")!({ company_id: "1" });
  assert.equal(broke.isError, true);
  assert.match(broke.content[0].text, /out of credits/);
});
