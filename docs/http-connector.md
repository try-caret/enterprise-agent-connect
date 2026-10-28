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
- **auth**: names the connect form fields that hold the credentials. Secrets must be fields with
  `"secret": true`; they are stored encrypted and never shown again.
  - `{"type": "header", "header": "X-Api-Key", "value": "<field>"}`, with `"prefix"` for a word
    before the value, e.g. `"header": "Authorization", "prefix": "SSWS "` (Okta) or
    `"prefix": "Token token="` (PagerDuty)
  - `{"type": "bearer", "token": "<field>"}`
  - `{"type": "basic", "username": "<plain field>", "password": "<secret field>"}`
  - `{"type": "oauth2", "tokenUrl": "https://…", "clientId": "<plain field>", "clientSecret": "<secret field>"}` (client credentials)
  - `{"type": "query", "param": "api_key", "value": "<secret field>"}`: the key goes in every
    call's query string. No tool param may use the same name.
- **headers**: sent with every call, e.g. `{"X-GitHub-Api-Version": "2022-11-28", "Accept":
  "application/vnd.github+json"}`. They are not secret, so they cannot carry credentials: names
  like `Authorization`, `Cookie` or anything with `key`, `token`, `secret` or `auth` in it are
  refused, as are headers the service sets itself (`Content-Type`, `Host`).
- **tools**: one per action. `path` takes `{param}` placeholders. Each param has a `type`
  (`string`, `number`, `integer`, `boolean` or `array`), a `description` the agent reads, and
  `in`: `path`, `query` or `body` (sent as one JSON object). Params are required unless
  `"required": false` or they have a `default`. Also:
  - `"enum": ["open", "closed"]`: the only values the agent may send.
  - `"default": "open"`: sent when the agent leaves the param out.
  - `"type": "array", "items": "string"`: a list of values. In the query it is sent as one
    `name=` per value, or as one value joined with `"join": ","`.
  - `"name": "pageToken"`: the name the app uses, when it is not a lower-case param name.
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
  e.g. `{"method": "GET", "path": "/health", "label": "Orders answers", "show": "name"}`.

What the service enforces on every call: path values are encoded and `.`/`..` are refused,
addresses stay under 8 KB and request bodies under 256 KB, answers under 1 MB and 30 seconds, only JSON or text answers
are passed back, and the agent's own headers are never sent to your app.

When an app needs logic JSON cannot express, write a connector in code instead
(`connectors/<app>/index.ts`, see the built-in ones) and pass it to `createAgentConnect`.
