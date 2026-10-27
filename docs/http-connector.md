# Add your own app without code

Most internal tools have a REST API. You can let agents use one by describing it in a single
`connector.json`. No code, no rebuild: the service runs it with its built-in `http` connector,
and your app gets what every other app gets: an API key per agent, the actions you tick for
each agent, value limits, a confirmation for anything that deletes, and an audit row for every
call.

## Try the example

`examples/internal-app` is a tiny orders API and its connector. From `deploy/docker`:

```sh
docker compose -f docker-compose.yml -f docker-compose.example.yml up --build --wait
```

"Orders" now shows up with the other apps. Connect it with the API key `example-key`, give an
agent a key with "Read orders", and ask it to look up order 1001.

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
  - `{"type": "header", "header": "X-Api-Key", "value": "<field>"}`
  - `{"type": "bearer", "token": "<field>"}`
  - `{"type": "basic", "username": "<plain field>", "password": "<secret field>"}`
  - `{"type": "oauth2", "tokenUrl": "https://…", "clientId": "<plain field>", "clientSecret": "<secret field>"}` (client credentials)
- **tools**: one per action. `path` takes `{param}` placeholders. Each param has a `type`
  (`string`, `number`, `integer`, `boolean`), a `description` the agent reads, and `in`: `path`,
  `query` or `body` (sent as one JSON object). Params are required unless `"required": false`.
- **capabilities** group the tools into what you tick per agent. A `DELETE` tool, or one marked
  `"destructive": true`, must sit in a capability marked `"destructive": true`.
- **limits**: give a param `"limit": "<list limit id>"` and an agent with values in that limit can
  only use those values, and must always send that param. A limit covers only the params that
  name it.
- **answerLimits**: for a read that cannot be limited by its params, such as an order by id,
  check the answer instead: `"answerLimits": [{"field": "customer", "limit": "customers"}]`.
  While the agent has values in `customers`, the answer's `customer` (a dot path such as
  `owner.id`; every item when the answer is a list) must be one of them, or the agent gets a
  refusal and nothing of the answer. An answer without the field is refused too. Only `GET`
  tools can have them, because a write has already happened by the time its answer arrives.
- **verify**: a request the service makes when the app is connected, to prove the credentials,
  e.g. `{"method": "GET", "path": "/health", "label": "Orders answers", "show": "name"}`.

What the service enforces on every call: path values are encoded and `.`/`..` are refused,
addresses stay under 8 KB and request bodies under 256 KB, answers under 1 MB and 30 seconds, only JSON or text answers
are passed back, and the agent's own headers are never sent to your app.

When an app needs logic JSON cannot express, write a connector in code instead
(`connectors/<app>/index.ts`, see the built-in ones) and pass it to `createAgentConnect`.
