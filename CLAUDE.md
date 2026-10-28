# Enterprise Agent Connect

Lets a company admin connect Jira, Outlook mail and calendar, SharePoint (files, Excel
spreadsheets and the directory), Teams, Microsoft Entra ID, Microsoft Intune, Slack, Salesforce, FedEx and Dell TechDirect once, from the admin UI, and lets Claude Managed Agents use those
apps through the service with per-agent limits. An agent is an API key made through the admin
API. Plain Node (`@hono/node-server` + `@modelcontextprotocol/server` v2), no UI: a host app
builds its admin pages on `packages/client`. Served under `/agent-connect` on the deployment's
host. Its own tables live in its own schema (`agent_connect`, see Database below).

## Endpoints

- `GET /agent-connect/healthz`
- `POST /agent-connect/mcp`: MCP Streamable HTTP for agents. `Authorization: Bearer arb_agent_…` (or `eac_agent_…`)
  (new keys stay `arb_agent_` for this release, so a rollback to 0.7.x still takes them)
  (a Managed Agents vault `static_bearer` credential keyed on this URL). Requests with an
  `Origin` header are refused. Tools are registered per request from the agent's grants: a tool
  outside its effective capabilities is never listed, and a call to one is audited as denied
  before the MCP server answers "not found".
  `POST /agent-connect/mcp/<agent id>` is the same for that agent's key only (any other gets 401):
  a Teams trigger's Claude agent and vault must use it, so a vault holding another agent's key
  cannot pass the trigger's check.
- `/agent-connect/admin/v1/*`: internal JSON API for the admin app only, server to server:
  `Authorization: Bearer $AGENT_CONNECT_ADMIN_TOKEN`; anything carrying `X-Forwarded-For`
  (i.e. through the public load balancer) is refused with 404. Every route is declared once in
  `packages/server/src/api.ts` (method, path, body and answer schemas); `packages/server/openapi.json` (OpenAPI 3.1, committed,
  also served at `GET /admin/v1/openapi.json`) is made from it by `npm run openapi`, and the
  router parses bodies with the same schemas (tests also check every JSON answer against them).
  A refusal is `{error}`: 400 bad input, 401 token, 404 unknown route or organization, 422 the
  app or vendor refused (`ConnectError`). v1 answers `/orgs/<org>/…` only for an organization it
  knows: `GET/POST orgs` (POST {name} makes one, id made here), `GET/PUT orgs/:org` (PUT
  {name?} makes it with that id if missing; a name given replaces the kept one). DEPRECATED:
  the unversioned `/agent-connect/admin/*` serves the same routes for one release (an admin of
  the previous release during an upgrade), and creates an organization it does not know on its
  first `/orgs/<org>/` request; it goes in the release after. Routes: `GET connectors`,
  `orgs/:org/connections` (GET; `PUT/PATCH/DELETE …/:app`; `POST …/:app/check`, which marks the
  connection connected or needs_reconnect unless `?record=false` (read only: limit choices, opening an app page, the Integrations tab; only the admin's Check again records, with body {by}, and a recorded check that changes the status is audited as `admin:check_app`);
  `PUT …/:app/sites` {sites} (apps with `siteAccess`: the whole list of added sites, normalized)
  and `POST …/:app/sites/check` {site?} (each added site's live access: granted, not_granted,
  not_found or failed); `POST …/:app/agent-access` {value, by} (apps with `agentAccess`: reads one
  agent's own address, e.g. its mailbox, with the app's sign-in; granted on 2xx, not_granted on
  401/403, not_found on 404, else failed. Live, Exchange answers 404, not 403, for a mailbox
  outside the app's scope (seen 2026-10-22), so not_found's copy names both causes; a check of an address no key holds is audited as
  `admin:agent_access_check`);
  `POST …/:app/consent` {capabilities, redirectUri} → {url, nonce} and
  `POST …/:app/consent/callback` {state, nonce, params, connectedBy} for approval-based apps:
  called twice, the first answer is {redirect} to the sign-in, the second {connection, checks}),
  `orgs/:org/confirmed` (GET the setup steps an admin marked done; `PUT/DELETE …/:step` {by}:
  only a `confirm` id of a loaded connector or provider, table `confirmed`;
  each mark and unmark is audited as `admin:confirm_step` / `admin:unconfirm_step` under the
  provider owning the step, else its first app),
  `orgs/:org/agents` (GET, each with `lastUsedAt`, `agents.last_used_at`, stamped in
  the same statement as each audit row; POST {name, createdBy};
  `PATCH …/:id` {status, by}, `DELETE …/:id` {by}; `POST …/:id/key` {by} rotates the key),
  `PUT orgs/:org/agents/:id/grants/:app` {capabilities, limits, by}, `DELETE` {by}. Every change
  to an app, key or grant (connect, ceiling `PATCH` {capabilities, by}, sites `PUT` {sites, by},
  disconnect `DELETE` {by}) writes an audit row with no agent (a key's in the same transaction
  as the key change, so a failed audit leaves the old key working): `agent_name` the admin, `tool`
  `admin:<action>`, `connector_id` the app (or `agents` for keys), the grant's capabilities and
  limits in `target`. The MCP endpoint refuses bodies over 4 MiB (413) and batches of more than
  20 messages (400) before auditing anything;
  `GET orgs/:org/audit?limit&agent&app`. Teams (docs/teams.md): `GET orgs/:org/teams/bots`;
  per agent `…/agents/:id/teams-bot` (POST {displayName, by} makes the key pair; a name another agent's bot has, in any case, is refused; the setup script makes and finds the Entra app by `uniqueName` `<branding.resourcePrefix>-bot-<bot row id>` (Arbium: `arbium`), and its removal commands delete an app only when it carries that name, and an Azure Bot only when its `msaAppId` is that app; GET `/kit` the
  PowerShell setup script; PUT {botAppId, by} proves the app id by signing in; POST `/check`;
  GET `/package` the Teams app zip; GET `/certificate` the bot certificate's public half (.cer,
  for setting up by hand); POST `/test` {conversationId, by} posts a short test message where the bot was added; POST `/certificate` and `/certificate/activate`; DELETE {by}
  returns the removal commands, and for a bot Add to Teams made first deletes its Azure Bot and app
  registration (the app only once its Azure Bot is gone, so the removal commands' ownership
  check still holds), `failed` listing what it could not; the setup script refuses a same-named
  Azure Bot whose `msaAppId` is another app); Add to Teams (docs/teams.md "Add to Teams"):
  `GET orgs/:org/teams/setup` (the Teams page's live step checks), `POST …/teams-bot/add`
  {displayName, by} (makes the bot in the tenant, each part only if missing; answers each part's
  state), `POST …/teams-bot/publish` {redirectUri} → {url, nonce} and `POST
orgs/:org/teams/publish/callback` {state, nonce, params, by} (the administrator's sign-in, then
  publishing to the org catalog), `GET …/teams-bot/targets?kind=team|person&q` and `POST
…/teams-bot/install` {kind team|chat|person, id, by}; `PUT orgs/:org/providers/:provider/settings`
  {values, by} (values a provider app's steps ask for, `settings`, kept in the provider's config and
  across a reconnect with the same sign-in, audited as `admin:settings`),
  `GET/PUT …/agents/:id/teams-trigger`, `GET
…/agents/:id/teams-runs?limit`; the organization's Claude API key at `GET/PUT/DELETE
orgs/:org/claude-key` (write-only, shown as its last four characters) and `GET
orgs/:org/claude-key/agents`.
- `POST /agent-connect/teams/bots/:botAppId/messages`: public, for Teams. Every request carries
  Microsoft's Bot Framework token, checked by hand (packages/server/src/teams/jwt.ts): signature from
  login.botframework.com keys, issuer, audience = this bot, 5 min skew, serviceUrl claim,
  msteams endorsement; then recipient, tenant and a Teams service address. The token is checked
  before the body is read; bodies over 256 KB get 413. Refused: 403, a log line, and for a request
  with a valid token, a plain reason on the bot (`last_refused_reason`, shown in the admin).
  Accepted: queued (dedupe on the activity id) and answered 200 at once. An approval card's
  button (`invoke` `adaptiveCard/action`, Action.Execute) is the exception: after the same checks
  it is answered in the response (Universal Actions' invoke answer: the card as it now is, or a
  message) and never queued, so it starts no run, trigger on or off (Rules: Teams, approvals). Connections (`PUT` {capabilities, fields, connectedBy},
  `PATCH` {capabilities}) carry the capability ceiling. Provider connections (one sign-in shared
  by several apps): `GET providers` (definitions, with `approval` when the deployment offers
  its shared app), `GET orgs/:org/providers` (each with certificate thumbprint/expiry, the
  waiting `nextCertificate`, and its `apps`), `PUT …/:provider` {fields, connectedBy} (proven by
  a sign-in before saving), `POST …/:provider/check` (sets connected or needs_reconnect only over the credential it checked, so a reconnect meanwhile wins), `POST …/:provider/certificate` (makes a key
  pair: before connecting the certificate to upload, after it one to rotate to),
  `POST …/:provider/certificate/activate` (signs in with the new one, then switches),
  `DELETE …/:provider` (disconnects every app signing in through it).
  Credentials are write-only: no route returns them. Agent keys are returned once, at creation
  or rotation.

## Env

`DATABASE_URL` (+ `DATABASE_SSL=disable` for local), `DATABASE_SCHEMA` (`agent_connect`),
`MIGRATE_ON_START` (`true`; `false` for installs that run `node dist/packages/server/src/server.js migrate` as their
own step), `CONNECT_ENCRYPTION_KEY` (32 bytes,
base64), `AGENT_CONNECT_ADMIN_TOKEN` (≥32 chars), `PORT` (3000), `PUBLIC_BASE_URL` (the
deployment's origin, the service under `/agent-connect`; where Teams reaches the bots; without it
the bot setup script and package say to set `global.publicBaseUrl`), `BRANDING` (JSON, see
Branding), `CONNECTOR_DIRS` (`:`-separated folders laid out like `connectors/`, `<id>/connector.json` each; one with an
`http` block runs on the generic module `connectors/http`, docs/http-connector.md; any other needs a
`ConnectorModule` passed to `createAgentConnect`).
`MICROSOFT_SHARED_APP=enabled` (off by default) offers approving the deployment's own
multi-tenant Entra app instead of the customer's own, and then needs
`MICROSOFT_SHARED_APP_CLIENT_ID` / `MICROSOFT_SHARED_APP_CLIENT_SECRET`: only for a deployment
that runs it, since anyone holding that secret reaches every tenant that approved it.
Renamed in 0.1.0: `SELF_HOST_PUBLIC_BASE_URL`, `ARBIUM_MICROSOFT_CLIENT_ID` and
`ARBIUM_MICROSOFT_CLIENT_SECRET` are still read for one release when the new name is unset, with
a deprecation line in the log (server.ts `renamed`; drop it in the release after). Key rotation: put the new key in `CONNECT_ENCRYPTION_KEY` and the old in
`CONNECT_ENCRYPTION_KEY_PREVIOUS`; old rows still open and are re-sealed on their next check,
recorded or read only (an app page or the Integrations tab; only over the credential it read, so a
reconnect meanwhile wins). Teams bot credentials
(`teams_bots`), Claude keys (`providers`, provider `anthropic`) and trigger signing secrets
(`teams_triggers.secret`) are re-sealed by a pass at startup while `CONNECT_ENCRYPTION_KEY_PREVIOUS` is set. Drop the
previous key once every row's `credential_key_id` is the new one in `connections`,
`providers` and `teams_bots`. Approvals in flight during a rotation must be restarted.

## Branding

`packages/connector-kit/src/branding.ts`: product name, logo URL, resource prefix (bot apps `<prefix>-bot-<id>`, Azure
Bots `<prefix>-<8>`, MCP server name), the Microsoft app, mailbox group (and alias) and Azure
role names, the Teams resource group and manifest developer. All have neutral defaults
("Enterprise Agent Connect", `eac`) except the developer, without which no Teams package is
made. Connector JSON says `{brand.productName}`, `{brand.microsoft.appName}` and so on, filled
when the connectors load (one naming nothing fails startup), so `GET /connectors` carries final
text; `GET /admin/v1/branding` serves it for UIs. Names are letters, digits and ` .,&()'-`
only (they reach PowerShell strings). A deployment that already created names in tenants
must keep passing the same values (`BRANDING`). One branding
per process. `packages/server/test/branding.test.ts` fails on any new mention of the first host's name in the
packages' `src/` or `connectors/`; the ones allowed are the stored `arbium_app` (`SHARED_APP_AUTH`), the deprecated
`madeByArbium` field (added only on the deprecated prefix's answers, not in the v1 schema) and
the old env names.

## Trigger runners

`packages/server/src/teams/runner.ts`: a Teams trigger names its runner (`runner`, default
`claude-managed-agents`); the core keeps who may ask, reply mode, the hourly, daily and
concurrency caps, the reply token and `teams_reply`, the watcher's timeouts.
`claude-managed-agents`: the organization's Claude key, the trigger's managedAgentId,
environmentId and vaultIds (unchanged). `webhook`: settings `{url}` (https), and a
`signingSecret` (32+ chars, write-only, sealed under `teams-trigger:<agent id>`; the trigger
shows `hasSigningSecret`). Saving pings it signed; each message is a POST of `{type: "message",
run, agent, bot, requestedBy, text, replyToken, mcpUrl}` with
`X-Agent-Connect-Signature: t=<unix>,v1=<hex HMAC-SHA256 of "<t>.<body>">` (2xx accepts; anything
else fails the run). The receiver answers with MCP `teams_reply` and that agent's own key; the
run is answered once a reply arrives, stopped after 30 min without one; it reports no cost (the
caps still count each answer's budget). A run may reply while still `starting` (a receiver
replying inside its POST handler). Each run records its runner (`teams_runs.runner`): only that
runner continues, watches or stops it, and saving the trigger with another runner stops the old
runner's live sessions. A run whose runner the deployment no longer offers is marked failed,
but its session is left running (nothing here can stop it). The webhook URL is an admin-only capability: the service POSTs from its
own network, so receivers check `t` freshness and dedupe on `replyToken` (`docs/teams.md`, As built).
`createAgentConnect({runners})` takes runner factories.

## Database

Schema `agent_connect` (`DATABASE_SCHEMA`), tables without a prefix (`agents`, `providers`,
`connections`, `grants`, `audit`, `confirmed`, `teams_*`, `approvals`, `organizations`); index and
constraint names keep the `agent_connect_<table>_...` spelling they had in public (`store.ts`
matches `agent_connect_providers_one_org_per_tenant`). `packages/server/src/migrate.ts` applies
`packages/server/migrations/NNNN_*.sql` on start, all pending in one transaction under an advisory lock, with
`search_path` set to the schema, recording each in `<schema>.schema_migrations` (version, name,
sha256). An install of the earlier layout (tables `public.agent_connect_*`, created by the host's
own migration before the service had a schema) is adopted
on the first start: with no history yet, the 12 tables are moved with `alter table ... set
schema` (rows, ids and sequences untouched) and recorded as 0001 (`adopted`); 0002 renames
them, adds a policy per table for the role running migrations (RLS stays forced, so a role
that cannot bypass RLS still reads the rows the backfill needs) and makes `organizations` from
the org ids already in them, 0003 drops two never-used columns, 0004 gives triggers a runner
(`claude-managed-agents` for every earlier row), its settings and a sealed signing secret, 0005
records each run's runner (a trigger moved to another runner stops the old runner's live
sessions and never reuses them), 0006 makes `approvals` (Teams approval cards; the
rollback drops it). The adopting run then makes simple views
`public.agent_connect_<old name>` over the moved tables and revokes every grant on them but the
owner's, so the earlier image brought back by an atomic rollback keeps reading and writing (only
for orgs already in `organizations`: the FK now points there). They are for one release: the
next release drops them in a migration before any column change, which they would block.
`rollback/to-legacy.sql` drops them first. Legacy
tables (not views) appearing after that (Flyway ran after the service on a fresh install) are dropped when
empty; with rows, start is refused. Organizations: the admin app names its own org ids; the
first admin request for an unknown one creates it (no name). `upgrade.test.ts` holds the
equivalence of both paths, the rows, the key and the rollback.

## Deploy

`deploy/docker` (image, `docker-compose.yml` with Postgres, `docker-compose.example.yml` adding
the example internal app) and `deploy/helm/agent-connect` (one replica, Service, optional
Ingress with only `/agent-connect/mcp` and `/agent-connect/teams`; `existingSecret` holds
`DATABASE_URL`, `CONNECT_ENCRYPTION_KEY`, `AGENT_CONNECT_ADMIN_TOKEN`). One replica only: Teams
run limits (concurrent runs, runs per hour, daily spend) are counted and admitted by one process,
so a rollout recreates rather than overlaps. The admin API is never public: it refuses any
request carrying `X-Forwarded-For`. Give the service a database role that owns only its schema.

## Layout

```
package.json          # one npm package for the server, kit and connectors (one lockfile);
                      # tsconfig.json builds all three into dist/ with their paths kept
connectors/<id>/      # one folder per app: connector.json, index.ts (its ConnectorModule:
                      # vendor client + toolset), tools.ts (MCP tools), vendor.ts (auth + HTTP)
                      # and any helpers of its own. Loaded by folder; the id is the folder name
connectors/<id>/connector.json # one per app: logo (admin /brand path), setup steps, form fields,
                      # capabilities → tools + vendor scopes (+ token `roles` when they differ;
                      # connector-level `scopes`: needed whatever is ticked, e.g. Teams' target search)
                      # + in-app permissions + step fills, value limits, verify/discover
                      # probes, optional `provider`, `siteAccess` (grant steps and a note to
                      # IT per added site). Fields may be choices (`options`, the first the
                      # default; an option's `scopes` add to the token roles expected, and the
                      # admin page asks such a choice in step 1, before the steps naming
                      # {scopes}; `requires` {field, value}: asked only while an earlier choice
                      # is that option, else its first option holds). Limits
                      # are list, text or map (`keys`: one value per chosen value of an earlier
                      # list limit, e.g. a folder per site). zod-validated at startup (connector-kit/src/connector.ts)
                      # Steps may carry `click` (the same step in the vendor's admin portals;
                      # with `copy` the admin page offers both, as tabs; every Microsoft 365
                      # step has both, apps of their own need not; a step with no portal page
                      # sets `clickUnavailable: true` instead, never a click line pointing at
                      # the commands, and the page says so), `live` (checked by the app's own
                      # setup check under that id, e.g. Teams' Azure steps; startup refuses an
                      # unknown id), `settings` (values entered at the step, kept in the provider
                      # connection's config, filled in as {id}; an id that is one of the
                      # connection's own keys, tenantId clientId auth certificate sites, fails
                      # at startup), `agentCheck` (Check
                      # access after it) and `confirm` (an id: not checkable live, so the admin
                      # marks it, stored per organization; steps sharing an id share the mark),
                      # `option` {field, value} (shown and holding back turning on only while
                      # that choice is made, e.g. Teams' Azure steps with add_to_teams on). A capability with
                      # `claudeKey` needs the organization's Claude key while ticked.
                      # A capability's `forbidden` is added to the agent's error when the app
                      # answers 403 to one of its tools (the missing role or site grant, with
                      # {scopes} filled), in the audit row too. A call a tool makes under
                      # another capability's permission passes that capability's tool as
                      # VendorFailed's `hintTool`, so the 403 names the right permission.
                      # `agentAccess` {limit, label, path with {value}, messages}: the agent's
                      # own address and its live check; `probes` per capability (a call
                      # that changes nothing: sendMail with an empty message answers 400 when
                      # allowed, PATCH {} on the inbox or calendar 200, all 403 without the
                      # role, seen live 2026-10-22) and `missing`, the copy naming those refused
connectors/microsoft/  # provider.json (one per shared sign-in): first-app setup steps (a step
                      # with `certificate: true` shows the certificate to download), fields
                      # (`optional` ones may be left empty), accountKey, principalKeys,
                      # sharedIdentity (shown while the shared app is used), revoke, addAgent
                      # (what one more agent needs, shown when its key is made with one of the
                      # apps each step names; checked at startup)
packages/connector-kit/src/ # what a connector builds on: connector.ts (connector and provider
                      # schemas, ConnectorModule, loading connectors/), common.ts (Toolset,
                      # audited: the audit row for every call), types.ts (Vendor, ConnectError),
                      # oauth.ts (client credentials with a cached token, vendor error text),
                      # branding.ts, certificate.ts. It still imports types from the server
                      # (Limits, TeamsAccess); untangling that is phase 3
packages/connector-kit/src/certificate.ts # RSA key pair + self-signed X.509 (DER written by hand; Node only parses
                      # X.509), SHA-1 thumbprint for display, PS256 (RSA-PSS) client assertion
                      # (x5t#S256), as Microsoft documents for certificate credentials
packages/server/src/index.ts # createAgentConnect(options): checks and wires everything (connectors and
                      # their modules, providers, branding, runners, migrations), returns
                      # {fetch, start, stop}; what an embedder imports
packages/server/src/builtin.ts # the built-in connectors' modules (connectors/<id>/index.ts)
packages/connector-kit/src/branding.ts # names shown and created in tenants (BRANDING), {brand.*} placeholders
packages/server/src/server.ts # the container's entry: env (old names for one release) → createAgentConnect, listen
packages/server/src/app.ts # HTTP routing, agent auth, admin API handlers (typed against api.ts)
packages/server/src/api.ts # the admin API's contract: every route, its body and answer schemas
packages/server/src/openapi.ts # openapi.json from api.ts (zod's toJSONSchema; named shapes as components)
packages/server/openapi.json # generated, committed (npm run openapi); CI fails on a diff
packages/client/      # @arbium/enterprise-agent-connect-client: createClient (the admin API v1 calls,
                      # named as the admin app calls them, plus appChecks/appStates/checkedSites),
                      # src/schema.ts (generated by openapi-typescript), src/types.ts (named
                      # shapes), src/steps.ts (pure step and status logic the admin pages share).
                      # `npm run client` regenerates schema.ts and builds lib/; CI fails when
                      # openapi.json or schema.ts is stale
packages/server/src/service.ts # connect (verify before save), check, grants, agent sessions
packages/server/src/store.ts # Postgres; every table named through tables() in the configured schema,
                      # never through search_path. Timestamps out as ISO 8601
packages/server/src/migrate.ts # the schema runner (see Database)
packages/server/migrations/ # 0001_tables.sql ... (plain SQL, unqualified names, never edited once
                      # applied: the runner refuses a changed checksum); rollback/to-legacy.sql
                      # (by hand only: back to the earlier layout in public)
packages/server/src/secrets.ts # AES-256-GCM credential sealing (AAD = org + app, or org + provider:<id>; key
                      # id per row), agent keys
scripts/test-databases.sh # makes an empty ac_test on a scratch Postgres, for the DB-backed tests
packages/server/src/consent.ts # signed approval state (org, app, capability ceiling, return address, tenant once
                      # approved, nonce, 15 min expiry)
connectors/jira/jql.ts # confines agent JQL to granted projects
connectors/salesforce/soql.ts # finds the one object a SOQL query reads; refuses anything reaching another
connectors/*/vendor.ts # per-app auth + HTTP (jira: Atlassian service account client credentials;
                      # microsoft/client.ts: the Microsoft connection (certificate, client secret or
                      # the deployment's shared app), Graph calls (or another resource's, by origin),
                      # the token-roles check, and mailboxVendor for apps needing only the
                      # connection: Outlook mail and calendar, Teams, Entra ID, Intune (it keeps
                      # only the app's own choice fields, from its connector.json; microsoftApp()
                      # makes their modules); sharepoint-files: SharePoint sites (sitesVendor);
                      # salesforce: External Client App client credentials on the My Domain;
                      # fedex, dell-techdirect on the kit's oauth.ts (client credentials, a
                      # cached token re-minted once on a 401); slack: a bot token, ok:false and
                      # 429 handling)
connectors/*/tools.ts # per-app MCP tools (sharepoint-files has tools.ts for files and excel.ts
                      # for spreadsheets, joined by the kit's `toolsets`; entra and intune reuse
                      # its Graph paging, outlook-calendar outlook-mail's). Folder, mail, calendar and people listings page with `cursor`
                      # (Graph's $skiptoken, or mail's and calendar's $skip as `s<n>` / token as `t<token>`,
                      # returned as `next`); shortcuts to other sites are never listed nor read by id
docs/teams.md         # Teams design (a bot per agent, real-time responses)
deploy/docker/        # Dockerfile (context: the repository root),
                      # docker-compose.yml (Postgres + the service) and .env.example
deploy/helm/agent-connect/ # small standalone chart (Deployment, Service, optional Ingress with
                      # only the public paths)
README.md             # quick start
.github/workflows/ci.yml # check, tests on Postgres 17, generated files current, package, image and chart build
.github/workflows/release.yml # v* tag: publishes the server and client to npmjs as
                      # @arbium/enterprise-agent-connect(-client) and the image to
                      # ghcr.io/try-caret/enterprise-agent-connect
packages/server/src/teams/ # Teams: approval.ts (approval cards: token, card before and after,
                      # invoke answers), jwt.ts (inbound token checks), bot.ts (the bot's certificate
                      # sign-in, Bot Connector + Graph calls), activity.ts (places, the decision
                      # whether a message starts a response, prompt text, reply tokens),
                      # runtime.ts (inbound queue worker, card clicks, session watcher, tools' access),
                      # admin.ts (bots, setup, Add to Teams, publishing, triggers, Claude key),
                      # provision.ts (Add to Teams' Graph and Azure calls, live step checks,
                      # catalog publish, installs, plain refusals), package.ts (manifest,
                      # generated icons, zip, setup script), runner.ts (TriggerRunner: what
                      # answers a trigger; Claude Managed Agents and the signed webhook),
                      # claude.ts (Managed Agents via @anthropic-ai/sdk), store.ts
packages/server/test/ # node:test, fakes for store and vendors; typechecked by tsconfig.test.json.
                      # The harness validates every admin answer against api.ts.
                      # packages/client/test: the client's calls against fakes and against the app
                      # upgrade.test.ts: an install from fixtures/legacy-20261023200000.sql (a
                      # frozen copy of the earlier layout) adopted, against a fresh one, and
                      # by an owner that cannot bypass RLS
                      # teams.test.ts: the Teams runtime and trigger admin with fake runners
                      # (replies while starting, sessions kept to their runner) and approval
                      # card clicks (approver only, once, unexpired, this bot's, this place's,
                      # trigger off; each store guard and two clicks at once)
                      # Generic machinery only: no suite or case for one connector's behavior
connectors/<id>/*.test.ts # one connector's own tool behavior, on the server's test harness
                      # (packages/server/test/helpers.ts) with its fake vendor answers; run by
                      # npm test, typechecked by tsconfig.test.json, left out of the build
```

## Rules

- Capabilities: each connector lists checkbox capabilities (`required` ones are always on,
  `destructive` ones are off unless the admin confirms a warning; `costs` ones, also destructive,
  are badged as spending money instead of deleting data), each naming its tools,
  vendor `scopes` (filled into `{scope}`/`{scopes}` in steps) and in-app `permissions` (listed
  by steps with `permissions: true`). The connection holds the ceiling (`capabilities text[]`),
  each grant the agent's ticks; effective = ticks ∩ ceiling, applied on every request. Lowering the ceiling (or
  reconnecting with a lower one) also strips the dropped capabilities from that app's grants, so
  raising it again gives agents nothing until they are ticked again. A grant is refused above
  the ceiling or for an app not connected, so nothing latent waits for a raise. Saved
  lists are validated, put in connector order and always include the required ones. Every
  toolset tool must belong to exactly one capability, and a capability with a tool annotated
  `destructiveHint` (it can destroy or overwrite) must be `destructive` (both checked at startup). Value limits
  (`limits` jsonb) are validated against the connector's `limits` (kind + pattern) and enforced
  inside each tool, which also filters results to them.
- Every tool call writes `audit` (allowed / denied / failed, reason, vendor
  status, the agent's name so removal keeps it). The target is what was sent (Jira: the
  project-scoped JQL); a call refused before any tool ran (not granted, or arguments its schema
  rejects, which app.ts records by wrapping each tool's schema) gets its arguments, shortened. Internal errors are logged and the agent gets a generic failure.
- Each connector names its `accountKey` (the config value for the account reached). A
  reconnect that reaches a different account, or a disconnect, deletes agents' grants for that
  app: their value limits meant something in the old account.
- Probes can carry a JSON `body` and an `ifEmpty` warning shown when they find nothing. A probe
  `{ label, check }` instead runs the vendor client's `checks[check](connection, expected)`
  (expected = the scopes the ceiling needs; it warns only about missing ones, never extra
  ones) for what no single API call shows; a missing check
  fails startup. A capability's `fill` sets `{name}` placeholders in steps while ticked (later
  ticks win), and its `note` is a caution shown next to its checkbox.
- Jira: `accessible-resources` is empty for service-account tokens, so the cloudId comes from
  `https://<site>.atlassian.net/_edge/tenant_info` and access is proven with `serverInfo`.
  The new search endpoint refuses unbounded JQL. The service account needs a project role
  with Browse projects, otherwise projects list but issues 404, so discover and
  `jira_list_projects` use `POST /rest/api/3/permissions/project` (BROWSE_PROJECTS). Calls send
  `Accept-Language: en` (errors otherwise follow the account's language). Every issue tool
  reads the issue first and refuses one outside the granted projects. `jira_transition_issue`
  resolves a transition by id or name from the issue's current transitions.
  `jira_create_issue` resolves the type by name from createmeta (no sub-task types) and refuses
  when the type needs fields beyond summary and description. `jira_delete_issue` sends
  `deleteSubtasks=false` unless the agent sets `delete_subtasks`. Service projects (Jira Service
  Management, `projectTypeKey` `service_desk`, read from the issue's own `project` field, never
  a project call): a platform
  comment can be public, so the customer sees it and is emailed (Atlassian's Comment `jsdPublic`:
  "Defaults to true when comments are created in the Jira Cloud Platform"). Seen live 2026-10-25
  (a test Jira Cloud site): a comment from a service account without a Service Management agent seat
  came out `jsdPublic` false anyway, even on a customer request. `jira_add_comment` takes `internal`
  (default false): true posts the platform comment with the `sd.public.comment` entity property
  `{internal: true}` (what JSM's agent view sends; Atlassian does not document it, only community
  posts do, so the answer's `jsdPublic` is the guard); it is refused outside a service project.
  When Jira answers `jsdPublic` other than false, the comment exists and is public (the customer
  may have been emailed): the tool answers normally, `internal` false, with a `warning` naming the
  comment id and saying not to retry (an error would be retried into a second public copy), and
  the audit row is allowed with target `<key> (internal asked, posted PUBLIC as comment <id>)`.
  Every answer's `internal` is what Jira saved: `jsdPublic === false` in a service project, null
  elsewhere. JSM's documented
  Create request comment (`/rest/servicedeskapi/request/{key}/comment`, `public: false`) is not
  used: tried live, it answers 403 "You do not have permission to view this request" to a service
  account without an agent seat, which service accounts were not offered on the same site.
  Needs only Add comments. `jira_get_issue` gives the reporter's and each
  comment author's `account_id` and `account_type` (atlassian, customer, app), from UserDetails,
  and no email (every agent with read would get customers' addresses); comments come from
  `GET /rest/api/3/issue/{key}/comment?orderBy=-created&maxResults=20` (the 20 newest, returned
  oldest first; a failed read fails the tool), each with `internal`: `!jsdPublic` in a service
  project, null elsewhere (where Atlassian says jsdPublic is always true) or when absent.
  `labels_added_by`: for each current label, the newest changelog entry adding it (author, id,
  type, `added` time; nulls when not found, e.g. set at creation), from
  `GET /rest/api/3/issue/{key}/changelog` (oldest first, 100 a page): the first page, then back
  from the newest, 5 calls at most; only after the project check, only with labels, and null
  when the changelog cannot be read (the rest of the issue is still returned). Not yet run
  against a real service project. Sources: developer.atlassian.com/cloud/jira/platform/swagger-v3.v3.json (Comment,
  UserDetails, ProjectDetails, issue comment GET/POST, issue changelog GET) and
  developer.atlassian.com/cloud/jira/service-desk/swagger.v3.json (Create request comment,
  CommentCreateDTO, CommentDTO, DateDTO), read October 25, 2026.
- Microsoft connection (provider `microsoft`, table `providers`, one per org):
  the apps with `"provider": "microsoft"` sign in through it, so its credentials are entered
  once; each app row keeps its own ceiling and settings (`provider_id`, no credential of its
  own; a check constraint allows exactly one of the two). Calls use the provider's config with
  the app's on top. `config.auth`: `certificate` (default; the service makes an RSA key pair, the
  admin uploads the certificate to their single-tenant app, tokens use a signed client
  assertion: PS256, `x5t#S256`, aud = the token endpoint, 5 minutes; still to confirm against a
  real tenant), `client_secret`, or
  `arbium_app` (the deployment's shared app, a stored value kept from the first host; approved by admin consent; only with
  `MICROSOFT_SHARED_APP=enabled`, otherwise its sign-in fails with a message to reconnect).
  Before connecting, a certificate lives in a `pending` row. Rotation: a new certificate waits
  in `nextCertificate`/`next_private_key` (calls keep the old one) until activate signs in with it
  and swaps; activating also replaces a client secret. Connect checks the token's `tid`, and
  `servicePrincipalId` is its `oid` (for Exchange's `New-ServicePrincipal`); `accountLabel` is
  the tenant id (reading the organization's name would need an org-wide permission). Reaching
  another tenant, or the same tenant through another app (`principalKeys`: clientId,
  servicePrincipalId, compared when both sides have one), disconnects its apps: Exchange roles
  and site grants belong to the old principal. Disconnecting the provider disconnects every
  app and their grants. Provider writes are conditional on the credential the caller read
  (`replaces`); a lost race is a `ProviderChanged` error telling the admin to reload. Each
  change to the sign-in (connect, certificate made or switched, disconnect) writes an
  `audit` row with no agent: `agent_name` is the admin, `connector_id` the
  provider, `tool` `admin:<action>`, `target` the old and new app/method/thumbprint and the apps
  disconnected. For a shared-app row with no stored clientId (moved from microsoft-365),
  `publicProvider` shows the shared app's id. Its credential is sealed under `provider:microsoft` (`credential_scope`; rows
  moved from an app keep that app's id until the next check re-seals them). Each app's
  `permissions` check decodes the `roles` of fresh Graph and Exchange Online
  (`outlook.office365.com`, shown with " (Exchange Online)") tokens: missing ones are the app's
  own (`roles`, else `scopes`, of its ceiling) and are the only warning; extra ones are the
  admin's choice. The approval
  flow (shared app only): its `state` is signed and its nonce must come back from the admin's
  own browser (the admin app keeps it in a cookie); the ceiling comes from the state. The
  approval's `tenant` is an unsigned query value, so it is only trusted after the admin signs
  in there (OIDC code flow, `scope=openid`): the id token must carry that `tid`, our `aud`, the
  derived nonce, no foreign `idp` (guests are refused) and, in `wids`, a role that can approve apps
  (Global, Privileged Role, Cloud Application or Application Administrator), since the approval
  itself is only a browser query value; `wids` needs the shared app's manifest to set
  `groupMembershipClaims` to `DirectoryRole`, and its absence fails the approval. It connects the provider, and the
  app too when the app has no form fields. A unique index binds a tenant approved for the
  shared app to one org.

- Outlook mail (`outlook-mail`): the app needs no Graph mail permission; mail access comes
  from the customer's Exchange RBAC for Applications role assignments to their own app's
  service principal, scoped by a management scope on membership of the mail-enabled security
  group named by `branding.microsoft.mailboxGroup` (`MemberOfGroup`; never a CustomAttribute, which tenants use
  for other data). After-connect steps have one role line per ticked capability. Its
  capabilities' token `roles` are empty. Mail
  tools take no mailbox argument: the mailbox is the grant's `mailbox` limit. `mail_move` goes
  only to inbox or archive: Deleted Items and Junk Email (purged automatically) would be
  deleting, a separate capability. `mail_delete` moves the message to Deleted Items (a Graph
  DELETE skips it and lands in Recoverable Items, which only an administrator reaches). The check runs inside `audited()` so the refusal is logged;
  refusals belong there, not in a zod enum, which the SDK rejects before any audit row. The
  admin app fills `{<config key>}` in steps from the provider's and app's config (e.g.
  `{clientId}`, `{servicePrincipalId}`) and `{scope}`/`{scopes}` from the chosen ceiling.

- Salesforce: client credentials against the customer's My Domain (an External Client App
  running as an integration user). The token response gives `instance_url` (must be an
  `https://*.my.salesforce.com` origin, checked on every call) and the identity URL, whose last
  segment is the Run As user's id (`config.userId`, read back by the verify probe through the
  `{userId}` path placeholder). Salesforce returns no token lifetime, so tokens are reused for
  10 minutes and re-minted once on a 401. With an `objects` limit, `salesforce_query` allows
  only single-object SOQL (no subqueries, dots, TYPEOF) and filters records by
  `attributes.type`. Update and delete read the record's type back first and refuse an id of
  another object. With an `objects` limit, delete also walks the `cascadeDelete` child
  relationships breadth-first (describing the object, then each allowed child once) and refuses
  when any cascade descendant, at any depth, is an object outside the limit (Salesforce would
  delete it too). Not counted: names ending Feed, History, Share or ChangeEvent
  (bookkeeping on any record) and children the global describe shows as not queryable (a child only
  automation creates still counts); the refusal names ten and says how many more. Probes can filter list items with `match` (discover keeps queryable,
  layoutable, non-setting objects; not `createable`, which follows the integration user's own
  rights, so a read-only user's objects would drop out; the admin page shows 20 values until opened).

- SharePoint (`sharepoint-files`, shown as "SharePoint"; the id stays for compatibility):
  files and Excel spreadsheets, one app with one site list, one site access check, one setup
  and one status. Capabilities, reads before writes so any write tick's `site_role` fill wins:
  `read_files` (graph_list_sites, graph_list_files, graph_read_file), `read_sheets` (the five
  excel_* reads), `write_files` (graph_upload_file, destructive), `write_sheets`
  (excel_update_range, excel_add_table_rows), `read_directory` (graph_list_users). None is
  required, so a key can have spreadsheets without files. Both write capabilities are
  `siteAccess.write` (the access check then proves write). Through the Microsoft
  connection. Tokens are client credentials for `graph.microsoft.com/.default`, cached per
  credential until near expiry and re-minted once on a 401. The connect form's `access` choice:
  `selected` (default, Sites.Selected: the app reaches a site only once it is granted) or
  `all_read` ("All sites, read-only", Sites.Read.All: then expected by its own
  check; it covers files and spreadsheets alike, and editing either still needs a write grant on
  the site). Sites are not entered at connect: the admin adds them
  one at a time afterwards (`config.sites`, normalized addresses: a bare host is the root site,
  any other path must start /sites/, /teams/ or /personal/, so a typo is refused rather than
  widened to the root; newline-separated, at most 50;
  kept across a reconnect; saved only over the config read, so a reconnect meanwhile is kept and
  the admin is asked to reload). A key's `sites` must be added sites, and a `folders` value has
  no `.` or `..` segment (the limit's pattern). A site search does not work under Sites.Selected (Microsoft documents
  `sites?search` as unsupported there). Per added site the admin page shows the grant
  (`siteAccess.grant`: New-MgSitePermission only. The Azure CLI cannot get
  Sites.FullControl.All: AADSTS65002, preauthorization, tried 2026-10-21) with the app id and role
  filled. Site names are limited to `[a-z0-9][a-z0-9._-]*` and every fill sits in single quotes,
  so a pasted command cannot run anything a site or an agent's address carries. Admin consent
  by command is an app role assignment through `az rest` (`az ad app permission admin-consent`
  exits 0 without granting anything, seen 2026-10-22). Graph answers an item's
  `parentReference.siteId` as the site collection id alone, not the whole site id (`inSite`). Steps name their
  `shell` (bash or powershell), shown above the commands, a note to IT (`siteAccess.request`), and Check access
  (`sitesVendor.sites.access`: resolves the address, then reads the library root; 401/403 is
  not_granted, 404/400 not_found). While a `siteAccess.write` capability is in the ceiling it
  also creates an upload session for `<prefix>-access-check.txt` and cancels it at once (DELETE on
  its pre-authenticated `uploadUrl`, no token; SharePoint can hold the name as an empty file
  while a session is open, seen on staging 2026-10-22): 403 means the site was given read only,
  so not_granted; any other non-2xx but 409 is failed. Seen live against a read-only grant
  (staging, 2026-10-22). Grant steps can carry `siteRole` (shown only for that `site_role`, e.g. the raise of a
  read grant to write with `Update-MgSitePermission`) and offer `click` too, saying Microsoft has no portal page for a site grant. Agents have
  no OneDrive of their own; a person's OneDrive (`https://<tenant>-my.sharepoint.com/personal/<name>`)
  is a site, reached only when added and granted. A token the `permissions` check cannot decode (Microsoft calls
  them opaque) gives a warning, never a failed connection; the `sites` check reopens each site,
  and its values (those that open) are the choices for the grant's `sites` limit (required). The
  grant's optional `folders` limit (map: site address to a folder path such as Contracts/Signed)
  keeps the key to that folder there. Tools resolve the granted address to a site id on every
  call. Paths (folder and file names) refuse `.`/`..`, `%` (nothing encoded reaches Graph),
  SharePoint's reserved characters, control characters and names SharePoint would trim; with a
  folder limit an empty folder means the key's folder and any other must be it or inside it
  (case-insensitive, NFC). An item read by id must have `parentReference.siteId` of the site,
  `driveId` of the site's default library (`GET /sites/{id}/drive`, the only one paths reach) and,
  with a folder limit, a parent path (percent-encoded, split before decoding) inside the folder;
  a missing or unreadable one is refused. Downloads only text types up to 1 MB, and uploads up to
  3 MB (the MCP endpoint takes 4 MiB bodies), never replacing a file unless asked (so it is
  marked destructive; base64 content is validated, never silently repaired). A key's sites are
  also kept to the connection's `config.sites`, so a stale or typed-in address is never reached
  even if the app holds a wider permission. `graph_list_users` returns name, email, title and
  department only, and pages with `cursor`.

- Outlook calendar (`outlook-calendar`): the mail model applied to calendars: no Graph
  Calendars permission, Exchange RBAC roles
  `Application Calendars.Read` / `.ReadWrite` on the same mailbox group scope (one
  after-connect line per ticked capability), and the grant's `mailbox` limit; no tool takes a
  mailbox. Event ids are only looked up under that mailbox (Exchange finds no other mailbox's
  events there). `calendar_list_events` needs a window with offsets, at most 62 days, and uses
  `calendarView`; create/update take wall-clock times plus a time zone name (update has no
  default zone: a new start or end without one is refused, and so is a zone without a start or
  end; a day that does not exist, such as February 31 or 24:00, is refused rather than rolled
  into the next) and answer with the event read back
  in UTC, since Microsoft answers a write in the zone it was sent in. Attendees get an
  invitation email from the mailbox to any address, which the write capability's label says.
  `calendar_respond_to_event` (write) posts Microsoft's `accept` / `tentativelyAccept` /
  `decline` action (202, no body) with an optional comment and `send_response` (default true);
  no proposed new time. Repeating events: create (and update with `series: true`) take a
  `recurrence` of daily / weekly / monthly / yearly, an interval, weekly days (default the
  start's weekday) and exactly one of `end_date` or `count` (at most 999), so no series is
  endless; monthly and yearly are Microsoft's absolute patterns on the start's date, and the
  range's `startDate` and `recurrenceTimeZone` are the start's date and `time_zone` (an update
  without a new start keeps the series' own date and zone). `calendarView` already lists each occurrence; events
  carry `type`, `series_id`, `recurrence`, `is_organizer` and `my_response`. Update and delete
  look the id up first: an occurrence's id changes that occurrence only, `series: true` sends
  the change to its series master, and a master's own id is refused without `series: true`,
  so a whole series is never changed or deleted by accident (deleting stays the delete
  capability). Sources: learn.microsoft.com/graph/api/event-accept, event-tentativelyaccept,
  event-decline, event-update, event-delete, resources/event and
  learn.microsoft.com/graph/outlook-schedule-recurring-events.
- SharePoint spreadsheets (excel_* tools, `read_sheets` / `write_sheets`): the files tools'
  site and folder limits apply. Every workbook tool reads the item back with the files tools'
  site, library and folder checks, and only opens `.xlsx` files.
  Ranges are A1 or A1:D20 (no whole rows or columns), at most 10,000 cells; tables return at
  most 500 rows. Reads are sessionless; each edit runs in a `persistChanges: true` session that
  is closed afterwards. Written text starting with `=`, `+`, `-`, `@`, tab or carriage return (OWASP's
  formula starts) gets a leading apostrophe (`literal()`):
  never a formula, which could link to other workbooks or outside addresses. The `site`
  parameter's description names the key's sites, since the agent may have no files tool to
  list them. Excel was once an app of its own (`excel`); older audit rows may name it.
- Microsoft Entra ID (`entra`, provider microsoft, read only): one capability `read` with
  Graph application permissions User.Read.All, GroupMember.Read.All and
  LicenseAssignment.Read.All. Group.Read.All is not needed: GroupMember.Read.All reads groups (group-get,
  group-list). Licenses come as `assignedLicenses` sku ids named through `GET /subscribedSkus`
  (LicenseAssignment.Read.All least; `licenseDetails` has no application permission). The
  manager comes from `$expand=manager` on user-get (the list-manager page shows no application
  permission). Tools: `entra_get_person` (id or sign-in name; title, department, office,
  account enabled, manager, direct reports, licenses), `entra_search_people` (`$search` on
  displayName and mail), `entra_person_groups` (`transitiveMemberOf/microsoft.graph.group`),
  `entra_group_members` (direct members, people and groups), `entra_search_groups`; listings
  page with `cursor`. Directory `$search`, casts and member `$select` send
  `ConsistencyLevel: eventual` with `$count=true`. Optional limits `groups` and `units` (group
  and administrative unit ids, at most 20 together, checkMemberObjects' cap): a person is in
  scope when `POST /users/{id}/checkMemberObjects` (transitive; User.ReadBasic.All and
  GroupMember.Read.All) finds any of them; a group when it is one of the key's groups or
  checkMemberObjects on it finds one. A group in scope only through a unit has each member
  checked. Every tool checks, lookups by id included; managers and reports outside are left
  out. Under limits a sign-in name Graph does not find gets the same refusal as one outside them,
  so a kept key cannot learn which names exist. Filtering follows Graph's paging, so under
  limits a listing (searches, a person's groups, a unit-scoped group's members) reads on, at
  most 5 pages (`MAX_SCOPED_PAGES`), until it has `max_results` or the listing ends; it can
  return up to a page more than asked, and `next` continues after the last page read. Administrative unit membership through checkMemberObjects is documented, not yet seen
  live. Sources: learn.microsoft.com/graph/api/user-get, user-list, user-list-manager,
  user-list-directreports, user-list-transitivememberof, user-list-licensedetails,
  subscribedsku-list, group-get, group-list, group-list-members,
  directoryobject-checkmemberobjects (read October 22, 2026).
- Microsoft Intune (`intune`, provider microsoft, read only): capability `read` with
  DeviceManagementManagedDevices.Read.All (also enough for detected apps, per
  intune-devices-detectedapp-list; DeviceManagementApps.Read.All is not needed). The connect
  choice `people_groups` (`off` default, or `groups`, which adds User.ReadBasic.All and
  GroupMember.Read.All) is what lets a key's optional `groups` limit work: a key with groups
  while the choice is off is refused before any call, never shown every device. In scope: a
  device whose `userId` (Intune's primary user) checkMemberObjects finds in the key's groups
  (one check per user per call; a user Graph answers 404 for is in none; a device with no user
  is out). Tools: `intune_list_devices` (`$filter` `contains(deviceName,…)` and `complianceState
eq`, quotes doubled), `intune_get_device`, `intune_device_apps` (Graph beta
  `managedDevices/{id}/detectedApps`: the relationship is beta only), `intune_person_devices`
  (`GET /users/{id or sign-in name}/managedDevices`, the user's managedDevices relationship in
  resources/intune-devices-user; `$filter=userId eq` on managedDevices is refused live with 400
  "Unsupported parameter", seen on staging 2026-10-22; under a groups limit the person is
  checked first with checkMemberObjects, and one outside the groups or not found gets Entra's
  refusal "This agent may not look up that person"; that check is the scope check, since this
  listing leaves each device's `userId` and `userPrincipalName` blank, seen live), `intune_compliance_summary` (counts by operating
  system and state over at most 5000 devices read, `complete` false past it; with a groups
  limit it lists each group's transitive user members once, `/groups/{id}/transitiveMembers/
microsoft.graph.user`, instead of one checkMemberObjects per user). Sources:
  learn.microsoft.com/graph/api/intune-devices-manageddevice-list, intune-devices-manageddevice-get,
  resources/intune-devices-manageddevice (v1.0 and beta), resources/intune-devices-user,
  intune-devices-detectedapp-list, directoryobject-checkmemberobjects (read October 22, 2026).
  Run against the staging tenant's Intune on 2026-10-22; `contains` on deviceName is documented
  on beta.
  Autopilot: optional capability `autopilot` ("Read Autopilot status", not destructive) with
  Graph application permission DeviceManagementServiceConfig.Read.All, the least Microsoft lists
  for list/get windowsAutopilotDeviceIdentity and for reading a device's deploymentProfile. Tool `intune_autopilot_status`
  ({serial_number}, letters, digits and dashes up to 50, refused before any Graph call; strict on
  purpose, so real serials with spaces, dots or underscores, as some virtual machines have, are
  refused, and the refusal says so):
  `GET /v1.0/deviceManagement/windowsAutopilotDeviceIdentities?$filter=contains(serialNumber,'<serial>')`
  (quotes doubled), then the exact serial (any case) is picked here, up to 5 pages of 50, else
  "Too many … match". Later pages follow `@odata.nextLink` as Graph gives it (whatever its paging
  parameter), only when it is on `https://graph.microsoft.com` and the same list path; any other
  link fails the call. Microsoft's list page documents no `$filter`, so `contains` on serialNumber
  is what Intune tooling is seen to use, not something the docs promise: unverified, never run live.
  Answer `{serial_number, found, deployment_state, enrollment_profile, group_tag, last_contact,
  managed_device}`; not found is `{serial_number, found: false}`. `deployment_state` is the
  identity's `enrollmentState` (v1.0: unknown, enrolled, pendingReset, failed, notContacted; beta
  adds blocked). `enrollment_profile` is the display name of the beta-only `deploymentProfile`
  relationship (`GET /beta/deviceManagement/windowsAutopilotDeviceIdentities/{id}/deploymentProfile`;
  the docs do not say what a device with no profile answers: 404 reads as null; any other
  failure also gives null, plus `enrollment_profile_error` "Microsoft Graph (beta) answered
  <status> for the profile", and the v1.0 answer stands). `group_tag` is
  null when empty. `managed_device` `{id, compliance, primary_user}` is the identity's
  `managedDeviceId` read from `managedDevices/{id}` (null for none, the all-zero id, or a 404;
  a 403 there carries the `read` capability's hint, since that read is its permission).
  Docs do not say when `enrollmentState` turns `enrolled` for a pre-provisioning (technician)
  run, so "ready for the user" is not claimed by the tool. Groups limit: every person tied to the
  device must be in the key's groups, else "This agent may not read that device" (as
  `intune_get_device`): the managed device's primary user (`userId`) and a user assigned on the
  Autopilot record (`userPrincipalName`, checked through checkMemberObjects). A device with
  neither (a fresh spare) gets only the minimum: `{serial_number, found, deployment_state,
  managed (whether a managed device exists), compliance, limited (a sentence saying why)}`, with
  no managed device id, names, group tag, last contact or profile (the beta call is not made).
  The other Intune tools treat a device with no user as out of scope; whether this one should
  refuse too is the owner's call, pending: `UNASSIGNED_UNDER_GROUPS` in tools.ts, set to
  `"refuse"`, does that. The `groups` limit's help text says the same. The `people_groups`
  choice rule is the same. Sources: learn.microsoft.com/graph/api/
  intune-enrollment-windowsautopilotdeviceidentity-list and -get, resources/
  intune-enrollment-windowsautopilotdeviceidentity (v1.0 and beta), resources/
  intune-enrollment-enrollmentstate, intune-shared-windowsautopilotdeploymentprofile-get (beta)
  (read October 25, 2026). Tests: `connectors/intune/autopilot.test.ts`, against a fake Graph.
- FedEx (`fedex`): a FedEx Developer Portal project's API key and secret key, OAuth client
  credentials at `/oauth/token` on `apis-sandbox.fedex.com` or `apis.fedex.com` (the connect
  form's `environment` choice; no default is assumed: the admin app always sends the choice),
  token cached until a minute before `expires_in` (3600) and re-minted once on a 401; a token
  answer without a lifetime is refused. Request shapes come from FedEx's OpenAPI files
  (developer.fedex.com/api/en-us/catalog/{authorization,track,rate,ship}/v1/docs.html, read
  October 21, 2026): `POST /track/v1/trackingnumbers` {includeDetailedScans, trackingInfo[] of
  trackingNumberInfo.trackingNumber} (at most 30), `POST /rate/v1/rates/quotes` {accountNumber
  {value}, requestedShipment}, `POST /ship/v1/shipments` {accountNumber, labelResponseOptions
  `URL_ONLY` (12-hour label links), requestedShipment}, `PUT /ship/v1/shipments/cancel`
  {accountNumber, trackingNumber, deletionControl `DELETE_ALL_PACKAGES`}; errors are
  `{transactionId, errors[{code, message}]}`. The connection keeps its account numbers
  (`config.accounts`, 9 to 12 digits, at most 50); config `project` is a hash of the API key
  (the `accountKey`). Capabilities: track (required), rate, ship (destructive: buys labels).
  The grant's `accounts` limit narrows a key's accounts; every `accountNumber` anywhere in a
  rate, ship or cancel request (the shipper's, a payor's for charges or duties, any case of
  the key, `{value}` or a plain string) must be the key's and the connection's, else the call
  is refused before FedEx sees it. Check: a Track API call for 000000000000 (any 2xx proves the
  Track API is in the project; whether production answers 2xx for an unknown number is not yet
  confirmed with real keys), and the account numbers as the limit's choices.
- Dell TechDirect (`dell-techdirect`): a client ID and secret Dell issues once it approves the
  company's Warranty API request. Token `https://apigtwb2c.us.dell.com/auth/oauth/v2/token`
  (client credentials), warranty `GET /PROD/sbil/eapi/v5/asset-entitlements?servicetags=` (up
  to 100, comma-separated, uppercased). Dell's reference needs a TechDirect sign-in, so these
  come from Dell's community threads and integrators' docs; nothing has run against real keys. Missing keys are refused with a plain message saying Dell issues them
  only after approval. `dell_warranty_lookup` returns model, ship date and each entitlement
  with an `active` flag; an answer that is not a list fails loudly. Asset components (parts)
  are left out: a separate API not confirmed under the same key. No limits.
- Slack (`slack`): a Slack app made from the manifest in the first setup step (YAML: a JSON
  manifest cannot take the `{scope}` line repeat without a trailing comma), installed to one
  workspace; the connect form takes its Bot User OAuth Token (`xoxb-`). `auth.test` proves it
  (`teamId` is the `accountKey`); the `workspace` check compares Slack's `x-oauth-scopes`
  header with the ceiling's scopes and warns about missing ones, and `channels` (discover)
  lists the channels the bot is in (`users.conversations`, 5 pages of 200 at most) as the
  `channels` limit's choices. Reads are GET with query arguments, writes POST JSON, all on
  `https://slack.com/api`. Slack answers most failures 200 `{ok: false, error}`: the vendor turns
  them into statuses (auth errors 401, missing_scope and the like 403, not found 404,
  ratelimited 429, any other 400) so probes and VendorFailed work as for other apps, and
  `slackError` puts common codes in words. A 429 is waited out (Retry-After, at most 10 s, twice:
  Slack did not run the call, so even a post is safe to resend); longer, the agent is told when
  to retry. The bot reaches only channels it is a member of (Slack refuses the rest with
  not_in_channel; no `chat:write.public`), and no im or mpim scope is asked, so no direct
  messages. Every tool takes a channel id (`C`/`G`), checked against the key's `channels` limit
  before any call; `slack_list_channels` filters to it. Capabilities: read (required: list
  channels, read a channel, read a thread, permalink), post (also thread replies), react
  (`already_reacted` answers success), upload (text up to 1 MB through
  files.getUploadURLExternal, the bytes POSTed to its `upload_url` only on `files.slack.com` with
  no token, then files.completeUploadExternal), people (users.lookupByEmail; no email in the
  answer), delete (destructive, chat.delete: a bot token deletes only the bot's own messages).
  Messages come trimmed: ts and its ISO time, author id and name (users.info, at most 20 a call;
  a failed lookup keeps the id), text with mentions, channel links and links made plain,
  thread_ts, reply_count, reactions, file titles. Search is left out: search.messages takes only
  a user token, which reads everything that person can (direct messages included), so a key's
  channel limit could only filter answers after the person's whole reach was searched. Sources:
  docs.slack.dev/reference/methods (files.getUploadURLExternal and files.completeUploadExternal
  read October 27, 2026, both taking JSON or form arguments; the other methods, rate limits and
  manifests from Slack's same reference, not re-read). Tested only against a fake Slack, never a
  real workspace.
- Teams (`teams`, provider microsoft, `signsInAs: "agentBot"`): the app row holds only the
  ceiling; each agent signs in as its own bot (its own single-tenant Entra app, Azure Bot and
  Teams app, made by the generated script with a service-made certificate sealed under
  `teams-bot:<bot id>`). An agent gets Teams tools only once its bot is connected. Capabilities:
  read (list teams, channels, read messages and threads; optional, since whoever messages the
  agent could ask it to read any place it is in), post, respond ("React in real time":
  `teams_reply`); the connector has no required capability, so an empty choice is refused. Limits `teams` (group ids) and `channels`; either one leaves chats out.
  The manifest (1.24) always asks for RSC ChannelMessage.Read.Group and ChatMessage.Read.Chat:
  every message where the bot is added reaches it, and Graph reads only there. Direct chats are
  not readable app-only. Inbound messages go through `decide()` (agent, respond capability,
  trigger on, its runner ready (Claude: the key); loop guard against any bot; tenant; a sender with an Entra id who is in
  allowed people or, with `anyone_in_organization`, a member Teams confirms (same tenant,
  `userRole` user, so never a guest); teams and key limits; reply mode; hourly cap (answers and
  failed starts), daily spending cap (spent in 24 h, from `teams_spend`: each
  turn's increase of the session's cost, timed, so a continued session's older turns age out;
  plus each live answer's budget) and
  concurrency, with at most one limit notice an hour per place; a run's own outcome notice is
  always sent) and start or continue one Managed
  Agents session per thread and requester (12 h / 20 turns; another person gets a new session),
  with the organization's Claude key, the trigger's agent, environment, vaults and budget (raised
  by one answer's budget on each continued turn). Saving a trigger checks the Claude agent has
  this agent's own MCP address (/mcp/<agent id>) and that a vault holds a bearer key for it,
  and refuses any other address of this service on that Claude agent or in any credential of its
  vaults, bearer or OAuth (the shared /mcp, another agent's /mcp/<id>, or one with a query or
  fragment; URLs are parsed, so ports, dot segments and case do not hide one). The same check
  (`verifyTrigger` in claude.ts) runs again before each start or turn (a pass is trusted for a
  minute), since both can change in Claude's console: a failure fails the run with its reason,
  says so in the thread and is audited as a denied `teams:run`. Every start or turn is
  audited as `teams:run` with the requester. The header's names are single-line JSON strings; the
  thread is quoted as untrusted text. `teams_reply` takes the run's reply token (hashed, the
  calling agent's, while the run is running or idle, under 12 h, at most 3 replies per message)
  and posts only to that thread, after the key's limits are checked again and the channel is
  confirmed to be in the team (as for every channel tool: a channel id alone could name another
  team's); a reply is counted only when its post goes through. A busy thread holds later messages until its session is idle.
  The watcher records a session that goes idle without replying as `no_reply`, one stopped at its
  budget as failed (both with a notice in the thread), interrupts turns running over 30 min, and
  fails runs whose session is gone. A run is `starting` until its message is sent (a continued
  turn too, so the previous turn's idle is never read as this one's); one left starting over 10
  minutes (the process stopped, or a write failed) is failed so it stops holding its thread. Removing the bot or the Claude key interrupts live sessions.
  A bot app with any Graph application roles is refused at connect and check (it should have
  resource-specific consent only).
  Add to Teams (docs/teams.md "Add to Teams"): the Microsoft connection's own app makes each
  agent's bot app-only (Application.ReadWrite.OwnedBy upsert by `uniqueName` `<prefix>-bot-<bot row
id>`, service principal, ARM Azure Bot `<prefix>-<8>` (prefix `branding.resourcePrefix`) and Teams channel through a custom role on
  `azureResourceGroup`), retries the bot's sign-in while Microsoft replicates (then says it is
  waiting; running it again makes nothing twice), and keeps `appObjectId`, `appId`, `azureBot`,
  `catalogAppId`, `publishedVersion` in the bot's config. Publishing needs an administrator's
  delegated AppCatalog.ReadWrite.All sign-in (signed state + browser nonce as for the shared-app
  approval; id token checked for app, tenant, nonce, expiry, no guest); installs consent to the
  package's RSC list. Not for a connection on the shared app. The teams connector's `scopes`
  (those permissions plus Team.ReadBasic.All and User.ReadBasic.All for the picker) are expected
  by every Teams ceiling's permission check; TeamsAppInstallation.ReadWriteAndConsentForUser.All
  only with the page's choice `person_installs` = `on` (kept in the teams connection's config;
  a person install is refused without it). Together these let the Microsoft app's credential read
  any team or chat by adding one of its bots there: the first step and the identity say so. A
  bot removed while Add to Teams runs is cleaned up by that run (its save finds no row). Nothing has run against a real tenant yet.
  Approvals (capability `approve`, "Ask people to approve or reject", not destructive):
  `teams_request_approval` ({team, channel} or {chat}, the same place and limit checks as
  `teams_post_message`; `approver` a sign-in name or Entra object id, resolved to its object id
  and name through the Microsoft connection with `GET /v1.0/users/{x}?$select=id,displayName`
  (User.ReadBasic.All, already expected by every Teams ceiling; the only call an `agentBot` app
  may make through the Microsoft connection, service.ts `DIRECTORY_LOOKUP`; so an agent with
  `approve` can test whether a sign-in name exists in the tenant); `title`, `text`,
  `facts` ≤ 20) posts an Adaptive Card 1.5 as the bot, card only (no text beside it, so Teams
  answers the card's own activity id), a new post in a channel, with Approve and Reject
  `Action.Execute` buttons (data: the approval row id) and an optional comment input, and
  returns `{approval_id, message_id, expires}`: `appr_` + 32 random bytes, stored only as its
  SHA-256, expiring after 7 days; a card that fails to post deletes its row. Action.Execute, not
  Action.Submit: Microsoft's Universal Actions send the click as an `adaptiveCard/action` invoke
  the bot answers in the same response with a card that replaces the clicker's view, in
  channels and chats alike, where Action.Submit arrives as a message (it would enter the message
  queue) and cannot answer with a card; no Action.Submit fallback, so a Teams client older than
  Universal Actions shows no buttons. The click arrives in a request authenticated by a Bot
  Framework token (jwt.ts; the body is not signed, so no layer may log the `Authorization`
  header); the approver is matched on the Entra object id Teams sends. A click counts only for
  an approval of the bot that was clicked (row id and bot id), posted in the team or chat the
  click came from (a forwarded card or a crafted 1:1 invoke finds nothing), pending and
  unexpired, from `from.aadObjectId` equal to the approver's object id (after the token,
  recipient and tenant checks), and is recorded once (one conditional update repeating every
  condition): `approved` / `rejected`, `decided_by` (the clicker's object id and
  Teams name), `decided_at`, the comment (≤ 1000). The card is then edited for everyone (`PUT
  {serviceUrl}/v3/conversations/{thread}/activities/{id}`; a failed edit is logged, the
  decision stands). Anyone else gets "Only <approver> can answer this request." and a second
  click or one after expiry gets the card as it is; each click is an audit row `teams:approval`
  (allowed or denied, with why). `teams_get_approval` {approval_id} answers only the agent that
  asked: status (pending, approved, rejected, expired), approver, decided_by, decided_at,
  comment, expires. Table `approvals` (title, text and facts kept in clear, to redraw the card).
  The agent picks its approver and nothing checks the choice: the guarantee is that the person
  the agent named clicked, not that this is the right person (an allowlist is the owner's
  decision, not built). A click before the card's activity id is saved is edited in by the tool
  once it is. Cards stay answerable after the capability or the agent is revoked, until they
  expire (current behaviour, docs/teams.md).
  Not yet run against a real tenant. Sources: docs/teams.md [28]-[31] (read October 25, 2026).
  Read messages carry `from` (display name), `from_id` (Graph `from.user.id`, the sender's Entra
  object id; null for apps and for users whose `userIdentityType` is not `aadUser` or
  `onPremiseAadUser`) and `from_kind` (`user` / `application`). Sources:
  learn.microsoft.com/graph/api/resources/chatmessagefromidentityset and teamworkuseridentity
  (read October 25, 2026).

## Commands

`npm test`, `npm run check` (server, connector kit, connectors, tests and client), `npm run dev`
(tsx watch), `npm run build` (tsc into dist/, plus the migrations and connector JSON beside it).
Local stack: `docker compose up --build --wait` in `deploy/docker` (README.md). After changing
a route or a shape in `packages/server/src/api.ts`, or anything in `packages/client/src`, run `npm run client` (regenerates
`packages/server/openapi.json`, `packages/client/src/schema.ts` and the admin app's vendored copy) and commit all three. The DB-backed tests
skip without their database: `scripts/test-databases.sh postgres://postgres@localhost:5432`
makes it (the user needs CREATEDB: the upgrade test makes its own databases), then `TEST_DATABASE_URL=…/ac_test DATABASE_SSL=disable npm test`. CI
(agent-connect-ci.yml) runs them against a Postgres service.

Connector files may also set `order` (position among its provider's apps on the Integrations
tab, which groups them under the provider).
