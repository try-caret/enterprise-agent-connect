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
- **Limits** narrow a capability to values: mailboxes, SharePoint sites and folders, Jira
  projects, Okta groups, Salesforce objects, ServiceNow assignment groups and companies, Workday
  organizations and reports, account numbers, or any list an `http` connector declares. They are checked on every call, including
  against ids the agent passes in and, for `http` connectors, against fields of the answer.
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

- Each built-in connector reaches only its vendor's addresses. An `http` connector reaches only
  its `baseUrl` (https unless `allowInsecureHttp` is set) and its token URL; redirects are not
  followed. Request and answer sizes and call time are capped.
- Content from apps (emails, issues, messages) is passed to agents as data. Teams thread content
  is marked untrusted for the agent.
- Microsoft Teams messages to agent bots are accepted only with a valid Bot Framework token for
  that bot.

## Audit

Every tool call (allowed, denied or failed) and every admin change writes an audit row: who,
which agent, which app and tool, a shortened target, the decision and the app's status.
`GET /orgs/{org}/audit` reads it.

## Running it

- Give the service a database role that owns only its schema.
- Keep `CONNECT_ENCRYPTION_KEY` and `AGENT_CONNECT_ADMIN_TOKEN` in a secrets manager.
- Expose only `/agent-connect/mcp` and `/agent-connect/teams`.
- Report vulnerabilities privately: see [SECURITY.md](../SECURITY.md).
