# Contributing

Thanks for helping. Issues and pull requests are welcome.

- **Bugs and ideas:** open an issue with the template that fits. Security problems go to
  [SECURITY.md](SECURITY.md), never to a public issue.
- **Pull requests:** one change per PR, with tests for anything beyond a trivial fix.
  `npm run check` and `npm test` must pass (see the README for the test database). If you change
  the admin API, run `npm run client` and commit the regenerated `openapi.json` and client.
- **New apps:** prefer the `http` connector when an app has a plain REST API
  ([docs/http-connector.md](docs/http-connector.md)); otherwise see
  [docs/connectors.md](docs/connectors.md).
- **Migrations** in `packages/server/migrations` are never edited once released; add a new one.
- **Compatibility:** don't remove or rename admin API fields, stored values or MCP tool names
  without a deprecation release.

By contributing you agree that your contribution is licensed under the
[Apache License 2.0](LICENSE).
