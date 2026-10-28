# Enterprise Agent Connect

Connect your company's apps once, then give each AI agent its own key with only the access you
grant it. Agents reach the apps through one MCP endpoint; every call is checked against the
agent's grants and written to an audit log. Agents never hold the apps' credentials.

- **Built-in apps:** Jira, Microsoft 365 (Outlook mail and calendar, SharePoint files and Excel,
  Teams, Entra ID, Intune), Jamf Pro, Okta, Slack, Salesforce, SAP S/4HANA, ServiceNow, FedEx, Dell TechDirect.
- **Your own apps without code:** describe an internal REST API or a SaaS app's in one
  `connector.json`, or draft one from its OpenAPI spec
  ([docs/http-connector.md](docs/http-connector.md)).
- **Per-agent access:** tick what each agent can do per app, and limit it to values (mailboxes,
  sites, projects, customers). Anything that deletes needs an explicit confirmation.
- **Self-hosted:** one Node service and Postgres. No UI ships; the admin API
  (`packages/server/openapi.json`) and the typed client (`packages/client`) are for building
  your own admin pages.

How it is secured: [docs/security.md](docs/security.md). Adding apps:
[docs/connectors.md](docs/connectors.md).

## Quick start (about 10 minutes)

Needs Docker with Compose v2, `openssl` and `curl`. This runs the service, Postgres and a small
example internal app (an orders API), then gives an agent read access to one customer's orders.

```sh
cd deploy/docker
cp .env.example .env
sed -i.bak "s|^CONNECT_ENCRYPTION_KEY=|CONNECT_ENCRYPTION_KEY=$(openssl rand -base64 32)|; \
  s|^AGENT_CONNECT_ADMIN_TOKEN=|AGENT_CONNECT_ADMIN_TOKEN=$(openssl rand -hex 32)|; \
  s|^POSTGRES_PASSWORD=|POSTGRES_PASSWORD=$(openssl rand -hex 24)|" .env && rm .env.bak
docker compose -f docker-compose.yml -f docker-compose.example.yml up --build --wait
```

Keep `CONNECT_ENCRYPTION_KEY`: it seals every stored credential, and nothing saved opens without
it. The service listens on port 3000 of this machine only and applies its database migrations.

**1. Make an organization and connect the example app** (its API key is `example-key`):

```sh
TOKEN=$(grep ^AGENT_CONNECT_ADMIN_TOKEN .env | cut -d= -f2)
admin() { curl -s -H "Authorization: Bearer $TOKEN" -H 'content-type: application/json' \
  "http://localhost:3000/agent-connect/admin/v1$@"; }

ORG=$(admin /orgs -X POST -d '{"name":"Example Org"}' | sed 's/.*"id":"\([^"]*\)".*/\1/')
admin /orgs/$ORG/connections/orders -X PUT \
  -d '{"capabilities":["read","notes"],"fields":{"api_key":"example-key"},"connectedBy":"Example Admin"}'
```

`capabilities` is the most any agent may get from this app; connecting runs the app's check.

**2. Make an agent and grant it "Read orders" for customer `acme` only.** The key is shown once:

```sh
NEW=$(admin /orgs/$ORG/agents -X POST -d '{"name":"demo","createdBy":"Example Admin"}')
KEY=$(echo "$NEW" | sed 's/.*"key":"\([^"]*\)".*/\1/')
AGENT=$(echo "$NEW" | sed 's/.*"agent":{[^}]*"id":"\([^"]*\)".*/\1/')
admin /orgs/$ORG/agents/$AGENT/grants/orders -X PUT \
  -d '{"capabilities":["read"],"limits":{"customers":["acme"]},"by":"Example Admin"}'
```

**3. Use it from Claude Code:**

```sh
claude mcp add --transport http orders http://localhost:3000/agent-connect/mcp \
  --header "Authorization: Bearer $KEY"
```

Ask Claude for acme's orders, then for order 1002 (a different customer's): the first works, the
second is refused, and both are in the audit log (`admin /orgs/$ORG/audit`). Any MCP client works
the same way, for example `curl`:

```sh
curl -s http://localhost:3000/agent-connect/mcp -H "Authorization: Bearer $KEY" \
  -H 'content-type: application/json' -H 'accept: application/json, text/event-stream' \
  -d '{"jsonrpc":"2.0","id":1,"method":"tools/call","params":{"name":"orders_get","arguments":{"id":"1001"}}}'
```

Stop with `docker compose -f docker-compose.yml -f docker-compose.example.yml down` (add `-v` to
delete the database too).

**Going further:** connect a real app from `GET /agent-connect/admin/v1/connectors` (each one
lists its setup steps, as click paths and as commands), add your own app
([docs/http-connector.md](docs/http-connector.md)), or embed the service in your own Node app
(below).

Only `/agent-connect/mcp` and `/agent-connect/teams` may be public. The admin API refuses any
request that carries `X-Forwarded-For`, so never route it through a public load balancer.

## Use it in your code

```sh
npm install @arbium/enterprise-agent-connect @arbium/enterprise-agent-connect-client
```

Both packages ship TypeScript types. Every snippet here is run by `scripts/consumer-test.sh`
(full files: `scripts/consumer/`) against the packed packages, in CI.

**Embed the service.** One instance per process, in a module of its own:

```ts
import { builtinConnectors, createAgentConnect } from "@arbium/enterprise-agent-connect";

export const eac = await createAgentConnect({
  database: { url: process.env.DATABASE_URL! }, // or a postgres.js instance
  encryptionKey: process.env.CONNECT_ENCRYPTION_KEY!,
  adminToken: process.env.AGENT_CONNECT_ADMIN_TOKEN!,
  connectors: [...builtinConnectors, orders, inventory], // default: every built-in app
});
eac.start(); // the Teams worker; await eac.stop() on shutdown
```

`connectors` is the list of apps this deployment offers. Leave it out for every built-in app;
pass `builtinConnectors.filter((c) => ["jira", "fedex"].includes(c.id))` for only some.

`eac.fetch` is a plain `(Request) => Promise<Response>` serving `/agent-connect/*`. Mount it at
that path, before any body parser (Express and plain Node use `@hono/node-server`):

```ts
// Hono
app.all("/agent-connect/*", (c) => eac.fetch(c.req.raw));

// Express 5
import { getRequestListener } from "@hono/node-server";
app.all("/agent-connect/*splat", getRequestListener(eac.fetch));

// Next.js: app/agent-connect/[...path]/route.ts (Node runtime)
const handler = (req: Request) => eac.fetch(req);
export { handler as GET, handler as POST, handler as PUT, handler as PATCH, handler as DELETE };

// Plain Node
import { serve } from "@hono/node-server";
serve({ fetch: eac.fetch, port: 3000 });
```

**The typed admin client.** The admin API refuses requests that came through a load balancer
(`X-Forwarded-For`), so call it from your server, over the network or in the same process:

```ts
import { createClient } from "@arbium/enterprise-agent-connect-client";

const admin = createClient({ baseUrl: "http://localhost:3000", token: process.env.AGENT_CONNECT_ADMIN_TOKEN! });
// In the same process as eac, with no network hop:
// createClient({ baseUrl: "http://agent-connect", token, fetch: (input, init) => eac.fetch(new Request(input, init)) })

const org = await admin.organizations.create("Example Org");
await admin.connect(org.id, "orders", {
  capabilities: ["read"],
  fields: { api_key: "example-key" },
  connectedBy: "Example Admin",
});
const { agent, key } = await admin.createAgent(org.id, { name: "demo", createdBy: "Example Admin" });
await admin.setGrant(org.id, agent.id, "orders", ["read"], { customers: ["acme"] }, "Example Admin");
const rows = await admin.audit(org.id, { agent: agent.id });
```

**Your own app, no code.** Pass a connector definition with an `http` block
([docs/http-connector.md](docs/http-connector.md)) straight into `connectors`. It is checked like
a `connector.json`; the type is `ConnectorJson`:

```ts
import type { ConnectorJson } from "@arbium/enterprise-agent-connect";

const orders: ConnectorJson = {
  id: "orders",
  name: "Orders",
  // ...setup steps, fields, capabilities, limits and verify, as in a connector.json
  http: {
    baseUrl: "https://orders.internal.example.com/api",
    auth: { type: "header", header: "X-Api-Key", value: "api_key" },
    tools: [{ name: "orders_get", description: "Read one order by its id.", method: "GET", path: "/orders/{id}",
      params: { id: { type: "string", description: "The order id", in: "path" } } }],
  },
};
```

**Your own app, as code.** Build a `ConnectorModule` on the connector kit and give it its
definition as `connector` ([docs/connectors.md](docs/connectors.md)):

```ts
import { z } from "zod";
import { audited, Denied, json, listLimit, type ConnectorModule } from "@arbium/enterprise-agent-connect/kit";

const inventory: ConnectorModule = {
  id: "inventory",
  connector: { id: "inventory", /* ...the rest of its connector.json */ },
  create: () => ({
    vendor: { connect: async (fields) => ({ ... }), call: async (connection, request) => ({ ... }) },
    toolset: {
      names: ["inventory_get_stock"],
      register(server, ctx, allowed) {
        if (allowed.has("inventory_get_stock"))
          server.registerTool("inventory_get_stock",
            { description: "How many of one item are in stock.", inputSchema: z.object({ sku: z.string() }) },
            ({ sku }) => audited(ctx, "inventory_get_stock", sku, async (mark) => {
              const skus = listLimit(ctx.limits, "skus");
              if (skus && !skus.includes(sku)) throw new Denied(`This agent may not read ${sku}`);
              const res = await ctx.call({ method: "GET", path: `/stock/${encodeURIComponent(sku)}` });
              mark(res.status);
              return json(res.data);
            }));
      },
    },
  }),
};
```

**Connect an agent.** Each agent uses its own key (`key` above) on one MCP endpoint,
`<your origin>/agent-connect/mcp`:

- Claude Code: `claude mcp add --transport http agent-connect https://agents.example.com/agent-connect/mcp --header "Authorization: Bearer $KEY"`
- The Claude API's MCP connector (beta `mcp-client-2025-11-20`): `mcp_servers: [{ type: "url", url: "https://agents.example.com/agent-connect/mcp", name: "agent-connect", authorization_token: key }]` with `tools: [{ type: "mcp_toolset", mcp_server_name: "agent-connect" }]`. The endpoint must be reachable from the internet.
- Claude Managed Agents: a vault `static_bearer` credential holding the key, keyed on the same URL.
- Any MCP client over Streamable HTTP: `Authorization: Bearer <key>`; the quick start's `curl` shows the raw call.

## Kubernetes

`deploy/helm/agent-connect` is a small chart: one replica (Teams run limits are counted in one
process), a Service, and an optional Ingress with only the public paths. It needs the image
(`image.repository`, `image.tag`; releases are `ghcr.io/try-caret/enterprise-agent-connect:<version>`) and a Secret with `DATABASE_URL`, `CONNECT_ENCRYPTION_KEY`
and `AGENT_CONNECT_ADMIN_TOKEN` (`existingSecret`).

## Layout

- `packages/server`: the service (HTTP, admin API, MCP, Teams, store, migrations) and its tests
- `packages/connector-kit`: what a connector builds on (schemas, tool helpers, vendor types)
- `packages/client`: typed admin API client
- `connectors/<app>`: one folder per app: `connector.json`, `index.ts`, `tools.ts`, `vendor.ts`
- `connectors/http`: runs any app described only in its `connector.json` (docs/http-connector.md)
- `examples/internal-app`: a small internal API and its connector, for the quick start
- `scripts/openapi-to-connector.ts`: drafts an http connector from an OpenAPI 3.x spec
- `deploy/docker`, `deploy/helm`: image, Compose file and chart

## License

[Apache License 2.0](LICENSE).

## Development

```sh
npm ci
npm run check
scripts/test-databases.sh postgres://postgres@localhost:5432   # a scratch Postgres
TEST_DATABASE_URL=postgres://postgres@localhost:5432/ac_test DATABASE_SSL=disable npm test
```
