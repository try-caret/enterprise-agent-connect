# Enterprise Agent Connect

Connect your company's apps once, then give each AI agent its own key with only the access you
grant it. Agents reach the apps through one MCP endpoint; every call is checked against the
agent's grants and written to an audit log. Agents never hold the apps' credentials.

- **Built-in apps:** Jira, Microsoft 365 (Outlook mail and calendar, SharePoint files and Excel,
  Teams, Entra ID, Intune), Slack, Salesforce, FedEx, Dell TechDirect.
- **Your own apps without code:** describe an internal REST API in one `connector.json`
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
with `createAgentConnect({...})` from the server package (`npm install
@arbium/enterprise-agent-connect`; the typed admin client is `@arbium/enterprise-agent-connect-client`).

Only `/agent-connect/mcp` and `/agent-connect/teams` may be public. The admin API refuses any
request that carries `X-Forwarded-For`, so never route it through a public load balancer.

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
