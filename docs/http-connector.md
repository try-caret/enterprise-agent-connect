# Add your own app without code

Most internal tools and SaaS apps have a REST API. You can let agents use one by describing it
in a single `connector.json`. No code, no rebuild: the service runs it with its built-in `http` connector,
and your app gets what every other app gets: an API key per agent, the actions you tick for
each agent, value limits, a confirmation for anything that deletes, and an audit row for every
call.

## Try the example

`examples/internal-app` is a tiny orders API and its connector. From `deploy/docker`:

```sh
docker compose -f docker-compose.yml -f docker-compose.example.yml up --build --wait
```

"Orders" now shows up with the other apps. Connect it with the API key `example-key`, give an
agent a key with "Read orders", and ask it to look up order 1001, or to list the open orders
(it pages through them two at a time).

## Start from an OpenAPI spec

If the app publishes an OpenAPI 3.x spec, draft the connector from it and pick the operations
agents need, by id or by tag:

```sh
npx tsx scripts/openapi-to-connector.ts https://api.example.com/openapi.json \
  --id tracker --operations listIssues,getIssue,createIssue --out my-apps/tracker/connector.json
```

The draft loads as it is, but it is a starting point, not a finished connector. The command
lists what it left out and what to check: reads go in a "Read" capability, other writes in
"Make changes" and deletes in a destructive "Delete"; answers are trimmed to fields guessed from
the spec; there are no limits and no paging yet; and the name, setup and other copy the admin
reads need rewriting. The spec must be JSON (convert a YAML one first).

## Write your own

Make a folder named after your app's id, put `connector.json` in it, and point the service at
the folder that holds it with `CONNECTOR_DIRS` (several folders: separate them with `:`).

A connector with an `http` block is like any other connector (name, setup steps, connect form
`fields`, `capabilities`, `limits`, `verify`), plus:

```json
"http": {
  "baseUrl": "https://orders.internal.example.com/api",
  "auth": { "type": "header", "header": "X-Api-Key", "value": "api_key" },
  "tools": [
    {
      "name": "orders_get",
      "description": "Read one order by its id.",
      "method": "GET",
      "path": "/orders/{id}",
      "params": { "id": { "type": "string", "description": "The order id", "in": "path" } }
    }
  ]
}
```

- **baseUrl**: every call goes to this address and under this path, and nowhere else. Redirects
  are not followed. It must be `https://`; for an app reachable only inside your network over
  plain http, add `"allowInsecureHttp": true` (this also allows a plain http `tokenUrl`).
  For an app each company runs at its own address, name a plain (not secret) connect form field
  in it as `{field}`, filled in per connection: the whole host (`"https://{host}/api/v4"`), one
  label of it (`"https://{tenant}.example.com/v1"`) or one path segment
  (`"https://api.example.com/customers/{customer}"`), nothing else. A host must be a host name
  only (a port is allowed), a label one word of letters, digits and hyphens, and a segment holds
  no slash and is not `.` or `..`; anything else is refused when the app is connected and before
  every call. An OData `serviceRoot` and a SOAP `endpoint` take the same placeholders.
- **allowedHosts** (optional): the hosts a filled-in address may reach, each exact or
  `"*.<domain>"` for any host under that domain, e.g. `["*.qualys.com"]`. Checked when the app is
  connected and before every call; an OData block takes it too.
- **auth**: names the connect form fields that hold the credentials. Secrets must be fields with
  `"secret": true`; they are stored encrypted and never shown again. Every field must be filled
  in unless it has `"optional": true`.
  - `{"type": "header", "header": "X-Api-Key", "value": "<field>"}`, with `"prefix"` for a word
    before the value, e.g. `"header": "Authorization", "prefix": "SSWS "` (Okta) or
    `"prefix": "Token token="` (PagerDuty)
  - `{"type": "bearer", "token": "<field>"}`
  - `{"type": "basic", "username": "<plain field>", "password": "<secret field>"}`
  - `{"type": "oauth2", "tokenUrl": "https://…", "clientId": "<plain field>", "clientSecret": "<secret field>"}` (client credentials; add `"scope"` for a token service that needs one, such as Microsoft Entra's, `"audience"` for one that asks for an audience, and `"clientAuth": "basic"` for one that wants the client id and secret as HTTP Basic rather than in the form, which is the default `"body"`). `tokenUrl` takes `{field}` placeholders like `baseUrl`, e.g. `"https://{host}/oauth/token"` for an app whose token service is on each company's own host; a filled-in one is held to `allowedHosts` too
  - `{"type": "query", "param": "api_key", "value": "<secret field>"}`: the key goes in every
    call's query string. No tool param may use the same name.
- **headers**: sent with every call, e.g. `{"X-GitHub-Api-Version": "2022-11-28", "Accept":
  "application/vnd.github+json"}`. They are not secret, so they cannot carry credentials: names
  like `Authorization`, `Cookie` or anything with `key`, `token`, `secret` or `auth` in it are
  refused, as are headers the service sets itself (`Content-Type`, `Host`).
- **fieldHeaders**: headers filled from connect form fields, secret or not, sent with every call
  beside `auth`, e.g. `{"x-client-id": "client_id"}` for an app that wants a second key. A header
  whose field is left empty is not sent. They cannot set the auth header, a fixed header or one
  the service sets.
- **tools**: one per action. `path` takes `{param}` placeholders. Each param has a `type`
  (`string`, `number`, `integer`, `boolean` or `array`), a `description` the agent reads, and
  `in`: `path`, `query` or `body` (sent as one JSON object). Params are required unless
  `"required": false` or they have a `default`. Also:
  - `"enum": ["open", "closed"]`: the only values the agent may send.
  - `"default": "open"`: sent when the agent leaves the param out.
  - `"type": "array", "items": "string"`: a list of values. In the query it is sent as one
    `name=` per value, or as one value joined with `"join": ","`.
  - `"name": "pageToken"`: the name the app uses, when it is not a lower-case param name
    (letters, digits and `_ . : - [ ]`, e.g. `"search:name"`).
  - `"field": "<plain field id>"`: filled from that connect form field, never by the agent (it
    is not offered to the agent at all), for a value the app wants in every request, such as
    the company's own server address. A `string` in the query or body, with no `default`,
    `enum` or `limit`.
- **result**: what the agent gets of a successful answer, so it reads the few fields it needs
  instead of the whole payload. `"items"` is the dot path of the list in the answer and
  `"fields"` the dot paths kept of each item (or of the answer itself without `items`):

  ```json
  "result": { "items": "data.issues", "fields": ["id", "title", "state", "user.login", "labels.name"] }
  ```

  gives `[{"id": 1, "title": "…", "state": "open", "user.login": "ann", "labels.name": ["bug"]}]`.
  A path through a list keeps that field of each of its items. A list the app leaves out of the
  answer reads as empty.
- **paging**: one page per call. The agent passes a cursor in `param` and gets
  `{"items": …, "next": …}` back, `next` being `null` on the last page. The next cursor comes
  from the answer's body (`"body": "meta.next_cursor"`), a header (`"header": "X-Next-Page"`) or
  the `rel="next"` address of the Link header (`"link": true`, taking that address's value of
  `param`, e.g. GitHub's `page`). The service never follows the address itself:

  ```json
  "paging": { "param": "page_token", "body": "nextPageToken" }
  ```
- **capabilities** group the tools into what you tick per agent. A `DELETE` tool, or one marked
  `"destructive": true`, must sit in a capability marked `"destructive": true`.
- **limits**: give a param `"limit": "<list limit id>"` and an agent with values in that limit can
  only use those values, and must always send that param. A limit covers only the params that
  name it.
- **answerLimits**: for a read that cannot be limited by its params, such as an order by id,
  check the answer instead: `"answerLimits": [{"field": "customer", "limit": "customers"}]`.
  While the agent has values in `customers`, the answer's `customer` (a dot path such as
  `owner.id`; every item when the answer is a list) must be one of them, or the agent gets a
  refusal and nothing of the answer. An answer without the field is refused too. A path through
  a list checks each of its items, e.g. `data.orders.customer`. Answer limits are checked on the
  whole answer, before `result` trims it. Only `GET` tools can have them, because a write has
  already happened by the time its answer arrives.
- A limited `array` param must have at least one value, and every value must be one of the
  agent's.
- **verify**: a request the service makes when the app is connected, to prove the credentials,
  e.g. `{"method": "GET", "path": "/health", "label": "Orders answers", "show": "name"}`. A
  `{field}` in its path, or a body value that is exactly `"{field}"`, is that connect form
  field's value.
- **discover** (optional): a request listing the values a limit chooses from, shown to the admin,
  e.g. `{"method": "GET", "path": "/customers", "label": "Customers", "show": "items[].id",
  "name": "items[].name"}`. `name` gives each value a label read from the same list item, so the
  admin picks "Acme Corp" rather than an id. For an answer that is itself a list, start the
  path with `[]`: `"show": "[].id", "name": "[].name"`.
- **Setup steps** may carry `"capabilities": ["<capability id>"]`: shown only while one of those
  capabilities is ticked, e.g. a script only one feature needs.

What the service enforces on every call: path values are encoded and `.`/`..` are refused,
addresses stay under 8 KB and request bodies under 256 KB, answers under 1 MB and 30 seconds, only JSON or text answers
are passed back, and the agent's own headers are never sent to your app.

A SOAP service or an OData service (SAP, Dynamics, SuccessFactors) is described the same way,
with a `soap` or `odata` block instead: see [soap-connector.md](soap-connector.md) and
[odata-connector.md](odata-connector.md).

To ship an app described this way as a built-in, put its folder in `connectors/` and add
`describedApp("<id>")` (or `describedApp("<id>", odataConnector)` for an `odata` block) to `builtinConnectors` (`packages/server/src/builtin.ts`).

## Add internal apps through the admin API

Each organization can also add its own apps while the service runs, as many as it wants: no
folder, no restart, no chart values. An internal app is the same definition as a
`connector.json` (with an `http`, `soap`, `odata` or `sql` block), saved for one organization
and checked by the same parser. Only that organization sees it or can use it; every other part
works as for any app: connecting and checking it, its capabilities and limits per agent, the
tools an agent's key gets, and the audit.

```sh
curl -X PUT "$EAC/agent-connect/admin/v1/orgs/$ORG/internal-apps/orders" \
  -H "Authorization: Bearer $AGENT_CONNECT_ADMIN_TOKEN" -H "content-type: application/json" \
  -d "{\"definition\": $(cat orders/connector.json), \"by\": \"ann@example.com\"}"
```

With the client: `saveInternalApp(org, "orders", definition, by)`, `internalApps(org)`,
`internalApp(org, id)`, `deleteInternalApp(org, id, by)`, and `orgConnectors(org)`, the apps an
organization can connect (the deployment's and its own). Then connect it and give agents
access as for any other app.

- The id in the address must be the definition's `id`. It cannot be one of the deployment's apps
  (built in or added with `CONNECTOR_DIRS`), and each tool name must be new to the organization:
  agents see every app's tools side by side.
- The definition holds no secret. Credentials are connect form fields, entered when the app is
  connected and stored encrypted, as for every app.
- `events` and `signsInAs` are not taken yet.
- Saving a changed definition of a connected app keeps the connection when it signs in and
  reaches the same place (same form fields, `accountKey`, and the `http`, `soap`, `odata` or
  `sql` block apart from its tools). The app's and agents' capabilities are cut to those the new
  definition still has, and an agent whose limits no longer fit it loses its access rather than
  getting more. Anything else, such as a new `baseUrl`, removes the connection and every
  agent's access to the app, since its credential was entered for the old address. The answer
  says what happened (`disconnected`, `grantsRemoved`), and an audit row records it.
- Deleting an internal app removes its connection and every agent's access with it, in one step,
  with an audit row.
- If the deployment later adds an app with the same id, or a tool with the same name, the
  organization's app is not used and `internalApps` shows it with a `problem`. For the same id,
  the deployment's app is not offered to that organization either, since the connection kept
  under that id was entered for the organization's own app: save it under another id, then
  delete the old one, which removes that connection.
- An internal app reaches whatever address its definition names, from inside the service's
  network, as an app loaded from a folder does. Saving one takes the admin token, so whoever
  can save definitions decides where the service connects: if people who should not reach the
  service's own network can save them (an organization's admins, through your own admin
  pages), limit where the service may connect with your network rules, and in particular
  block the cloud's instance metadata address (169.254.169.254).

To start from an OpenAPI spec, `POST /agent-connect/admin/v1/internal-apps/draft` with
`{"spec": <the spec as JSON>, "id": "tracker", "tags": ["issues"]}` (or `"operations"`) answers
the same draft as `scripts/openapi-to-connector.ts`, with its `notes` and a `problem` when it
would not save as it is.

When an app needs logic JSON cannot express, write a connector in code instead
(`connectors/<app>/index.ts`, see the built-in ones) and pass it to `createAgentConnect`.
