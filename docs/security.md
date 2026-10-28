# How Enterprise Agent Connect is secured

## Who can do what

- **Admins** use the admin API with `AGENT_CONNECT_ADMIN_TOKEN`. The admin API is for your own
  admin pages, inside your network: it refuses any request carrying `X-Forwarded-For`, so it
  cannot be reached through a public load balancer by mistake.
- **Agents** use `POST /agent-connect/mcp` with their own key. A key is shown once and stored only
  as its SHA-256 hash. Requests with an `Origin` header are refused, so a browser page cannot
  use a key it does not hold.
- **What an agent can do** in an app is the capabilities ticked on its key, within that app's
  ceiling (the most any agent may get from it). Tools outside that are never listed, and a call
  to one is refused and audited.
- **Limits** narrow a capability to values: the mailbox of Outlook mail and calendar, SharePoint
  sites and a folder per site, Teams teams and channels, Entra ID groups and administrative
  units, Intune groups, Dynamics 365 business units, owning teams and other tables, Jira
  projects, Jamf Pro sites, Okta groups, Slack channels, Salesforce objects, Gong workspaces,
  customer email domains and CRM accounts, SAP S/4HANA company codes, sales and purchasing organizations and plants, NetSuite subsidiaries, locations and
  saved searches, ServiceNow assignment groups and companies, Workday organizations and reports,
  FedEx account numbers, or any list an `http`, `soap`, `odata` or `sql` connector declares.
  They are checked on every call, including against ids the agent passes in and, for those
  connectors, against fields of the answer (an OData update reads the record first and an SQL
  write is rolled back unless every row it changed is inside the agent's values).
- **Destructive capabilities** (anything that deletes or overwrites) are off unless the admin
  confirms a warning, and every tool that deletes must sit in one. The service refuses to start
  otherwise.

## Credentials

- App credentials are sealed with AES-256-GCM under `CONNECT_ENCRYPTION_KEY`. The additional
  data binds each one to its organization and app, so a sealed value cannot be moved to another
  row. Keys rotate with `CONNECT_ENCRYPTION_KEY_PREVIOUS` (see CLAUDE.md).
- Microsoft connections use the customer's own app registration with a certificate the service
  makes; the private key never leaves the service. Mailboxes are limited with Exchange RBAC for
  Applications and SharePoint with per-site grants, so the app itself cannot reach more than
  what the admin granted.
- Agents never see credentials, and errors passed to agents never contain them.

## Calls to apps

- Each built-in connector reaches only its vendor's addresses, and no connector follows a
  redirect: a redirect is refused as an error, so a token or request body is never sent on to
  another host. The one exception is a SharePoint file's content, which Microsoft Graph serves
  from a SharePoint download address; that address is called without the token, and only when
  it is on SharePoint. An `http` connector reaches only its `baseUrl` (https unless
  `allowInsecureHttp` is set) and its token URL, a `soap` connector only its endpoint, an `odata`
  connector only its service root and token URL. Request and answer sizes and call time are
  capped.
- Agents never write the app's query language or markup: SOAP values are XML-escaped into a fixed
  body, and OData filters, keys and literals are written by the service from typed values. SOAP
  answers are read without any document type declaration, so XML entities cannot reach files or
  expand.
- Content from apps (emails, issues, messages) is passed to agents as data. Teams thread content
  is marked untrusted for the agent.
- Microsoft Teams messages to agent bots are accepted only with a valid Bot Framework token for
  that bot.
- App events (`/agent-connect/events/<subscription id>`) are accepted only with the
  subscription's secret (an HMAC of the body or the secret in a header, compared in constant
  time, before the body is parsed), once per delivery id, and within the app's replay window.
  An event starts only agents granted on that app and within their limits, and its content is
  passed to the agent as untrusted data ([triggers.md](triggers.md)).

## Audit

Every tool call (allowed, denied or failed), every app event delivery and trigger decision, and
every admin change writes an audit row: who,
which agent, which app and tool, a shortened target, the decision and the app's status.
`GET /orgs/{org}/audit` reads it.

## Running it

- Give the service a database role that owns only its schema.
- Keep `CONNECT_ENCRYPTION_KEY` and `AGENT_CONNECT_ADMIN_TOKEN` in a secrets manager.
- Expose only `/agent-connect/mcp`, `/agent-connect/teams` and `/agent-connect/events`.
- Report vulnerabilities privately: see [SECURITY.md](../SECURITY.md).
