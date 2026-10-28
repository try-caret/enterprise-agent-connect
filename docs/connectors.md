# Adding an app

Every app is a folder in `connectors/` (or in a folder you pass with `CONNECTOR_DIRS`) holding its
`connector.json`: name, setup steps (as click paths and as commands), the connect form, the
capabilities an admin can tick (each a set of tools), the limits a grant can carry, and how a
connection is checked. Schema: `packages/connector-kit/src/connector.ts`.

There are two ways to give an app its tools.

## Without code: the `http` connector

For an app with a REST API, add an `http` block to `connector.json` and describe each tool as a
method and path. The image runs it as is. See [http-connector.md](http-connector.md) and the
example in `examples/internal-app`.

## With code

For an app that needs logic JSON cannot express (paging, multi-step writes, checks before a
delete), write a `ConnectorModule` next to the JSON:

- `vendor.ts`: how it signs in and calls the app (`connect`, `call`), built on
  `packages/connector-kit/src/types.ts`.
- `tools.ts`: the MCP tools, each wrapped in `audited(...)` from `connector-kit/src/common.ts`,
  enforcing the grant's limits with `listLimit`, `mapLimit` and friends.
- `index.ts`: `export default { id, create }`.

Then either add it to `packages/server/src/builtin.ts` (to contribute it here) or pass it to
`createAgentConnect({ connectors: [...builtinConnectors, myApp] })` in your own Node app. Outside
this repo, import the same helpers and types from `@arbium/enterprise-agent-connect/kit`, and give
the module its definition inline (`connector: { ...its connector.json }`) instead of a folder in
`extraConnectorDirs`. A definition with an `http` block can go into `connectors` alone, with no
module. A built-in app is offered only while its module is in `connectors`. Example: the
README's "Use it in your code".

At startup the service checks that every tool a capability names exists, that each tool is in
at least one capability, and that every tool marked destructive sits in a destructive
capability. A mismatch stops the service.

## Tests

Put a connector's tests next to it, as `connectors/<app>/*.test.ts` (see `connectors/jira`). Run `npm test`
(see the README for the test database).
