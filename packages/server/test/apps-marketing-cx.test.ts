import assert from "node:assert/strict";
import { createServer, type IncomingMessage } from "node:http";
import type { AddressInfo } from "node:net";
import { after, before, test } from "node:test";
import type { McpServer } from "@modelcontextprotocol/server";
import type { ZodType } from "zod";
import { connectorSchema, loadConnectors, type Connector } from "../../connector-kit/src/connector.ts";
import type { ToolContext } from "../../connector-kit/src/common.ts";
import { assertToolsImplemented } from "../src/service.ts";
import { builtinConnectors } from "../src/builtin.ts";
import { httpTools, httpVendor } from "../../../connectors/http/index.ts";

// The marketing and customer apps described only in connector.json: each loads as the server
// loads it, and one tool of each reaches a local stand-in for the app as its docs describe.

const IDS = ["adestra", "cvent", "forsta-decipher", "intercom", "sprout-social", "synthesia", "wordpress"];

type Seen = { method: string; url: string; headers: IncomingMessage["headers"]; body: string };
const seen: Seen[] = [];
let origin = "";
const server = createServer(async (req, res) => {
  let body = "";
  for await (const chunk of req) body += chunk;
  seen.push({ method: req.method!, url: req.url!, headers: req.headers, body });
  res.writeHead(200, { "content-type": "application/json" });
  res.end(
    JSON.stringify(
      req.url === "/token" ? { access_token: "minted", expires_in: 3600 } : { data: [], paging: {} },
    ),
  );
});
before(async () => {
  await new Promise<void>((done) => server.listen(0, "127.0.0.1", done));
  origin = `http://127.0.0.1:${(server.address() as AddressInfo).port}`;
});
after(() => server.close());

const loaded = () => loadConnectors(undefined, (id) => IDS.includes(id));

// The app pointed at the local server: a fixed host is swapped for it, a {host} filled with it.
function local(connector: Connector): Connector {
  const http = connector.http!;
  const baseUrl = http.baseUrl.replace(/^https:\/\/(\{host\}|[^/]+)/, (_, host) =>
    host === "{host}" ? "http://{host}" : origin,
  );
  const auth =
    http.auth.type === "oauth2" ? { ...http.auth, tokenUrl: `${origin}/token` } : http.auth;
  return connectorSchema.parse({
    ...connector,
    http: { ...http, baseUrl, auth, allowedHosts: undefined, allowInsecureHttp: true },
  });
}

async function call(
  connector: Connector,
  tool: string,
  args: Record<string, unknown>,
  conn: { config?: Record<string, string>; secret: Record<string, string> },
) {
  const vendor = httpVendor(connector);
  const handlers = new Map<string, (a: unknown) => unknown>();
  const recorder = {
    registerTool: (name: string, c: { inputSchema: ZodType }, h: (a: unknown) => unknown) =>
      handlers.set(name, (a) => h(c.inputSchema.parse(a))),
  } as unknown as McpServer;
  const c = { config: { host: new URL(origin).host, ...conn.config }, secret: conn.secret };
  const ctx = {
    limits: {},
    config: c.config,
    call: (r) => vendor.call(c, r),
    audit: async () => {},
  } as ToolContext;
  const toolset = httpTools(connector.http!);
  toolset.register(recorder, ctx, new Set(toolset.names));
  const result = (await handlers.get(tool)!(args)) as { isError?: boolean; content: { text: string }[] };
  assert.ok(!result.isError, result.content[0].text);
  return seen.at(-1)!;
}

test("each app loads as the server loads it, with every tool in a capability", () => {
  const apps = loaded();
  assert.deepEqual(apps.map((c) => c.id).sort(), IDS);
  for (const id of IDS) assert.ok(builtinConnectors.some((m) => m.id === id), `${id} is offered`);
  assertToolsImplemented(
    apps.map((connector) => ({
      connector,
      toolset: httpTools(connector.http!),
      vendor: httpVendor(connector),
    })),
  );
});

test("one tool of each app reaches the app as its docs describe", async () => {
  const app = (id: string) => local(loaded().find((c) => c.id === id)!);

  let req = await call(app("adestra"), "adestra_search_campaigns", { project_id: 7 }, { secret: { api_token: "t" } });
  assert.equal(req.method, "GET");
  assert.equal(req.url, "/api/rest/1/campaigns?search%3Aproject_id=7&paging%3Apage_size=50");
  assert.equal(req.headers.authorization, "TOKEN t");

  req = await call(app("adestra"), "adestra_send_test", { campaign_id: 3, emails: ["a@x.com"] }, { secret: { api_token: "t" } });
  assert.equal(req.method, "POST");
  assert.equal(req.url, "/api/rest/1/campaigns/3/send_test");
  assert.deepEqual(JSON.parse(req.body), { emails: ["a@x.com"] });

  seen.length = 0;
  req = await call(
    app("cvent"),
    "cvent_list_attendees",
    { filter: "event.id eq 'e1'" },
    { config: { client_id: "cid" }, secret: { client_secret: "cs" } },
  );
  const token = seen.find((s) => s.url === "/token")!;
  assert.equal(token.headers.authorization, `Basic ${Buffer.from("cid:cs").toString("base64")}`);
  assert.equal(req.method, "GET");
  assert.equal(req.url, "/ea/attendees?filter=event.id+eq+%27e1%27&limit=100");
  assert.equal(req.headers.authorization, "Bearer minted");

  req = await call(
    app("forsta-decipher"),
    "decipher_get_quotas",
    { directory: "53b", project: "180401" },
    { secret: { api_key: "k" } },
  );
  assert.equal(req.url, "/api/v1/surveys/selfserve/53b/180401/quota");
  assert.equal(req.headers["x-apikey"], "k");

  req = await call(
    app("intercom"),
    "intercom_add_note",
    { conversation_id: "c9", admin_id: "a1", body: "Entitled: yes" },
    { secret: { access_token: "it" } },
  );
  assert.equal(req.method, "POST");
  assert.equal(req.url, "/conversations/c9/reply");
  assert.equal(req.headers.authorization, "Bearer it");
  assert.equal(req.headers["intercom-version"], "2.14");
  assert.deepEqual(JSON.parse(req.body), {
    admin_id: "a1",
    body: "Entitled: yes",
    message_type: "note",
    type: "admin",
  });

  req = await call(
    app("sprout-social"),
    "sprout_create_draft",
    { group_id: 5, customer_profile_ids: [11, 12], text: "Hello" },
    { config: { customer: "687751" }, secret: { api_token: "st" } },
  );
  assert.equal(req.url, "/v1/687751/publishing/posts");
  assert.equal(req.headers.authorization, "Bearer st");
  assert.deepEqual(JSON.parse(req.body), {
    group_id: 5,
    customer_profile_ids: [11, 12],
    text: "Hello",
    is_draft: true,
  });

  req = await call(app("synthesia"), "synthesia_list_videos", { source: ["workspace"] }, { secret: { api_key: "sk" } });
  assert.equal(req.url, "/v2/videos?limit=20&source=workspace");
  assert.equal(req.headers.authorization, "sk");

  req = await call(
    app("wordpress"),
    "wordpress_create_draft",
    { title: "T", content: "<p>C</p>" },
    { config: { username: "agents" }, secret: { app_password: "ap" } },
  );
  assert.equal(req.method, "POST");
  assert.equal(req.url, "/wp-json/wp/v2/posts");
  assert.equal(req.headers.authorization, `Basic ${Buffer.from("agents:ap").toString("base64")}`);
  assert.deepEqual(JSON.parse(req.body), { title: "T", content: "<p>C</p>", status: "draft" });

  // WordPress drops every post the user may not edit from an edit-context list, and refuses
  // one by id, so lists and reads ask for the view a reader gets unless told otherwise.
  const wp = { config: { username: "agents" }, secret: { app_password: "ap" } };
  req = await call(app("wordpress"), "wordpress_list_posts", { search: "launch" }, wp);
  assert.equal(req.url, "/wp-json/wp/v2/posts?search=launch&per_page=20");
  req = await call(app("wordpress"), "wordpress_get_post", { id: 5 }, wp);
  assert.equal(req.url, "/wp-json/wp/v2/posts/5?context=view");
});
