import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { test } from "node:test";
import { connectorSchema } from "../../connector-kit/src/connector.ts";
import { assertToolsImplemented } from "../src/service.ts";
import { httpTools, httpVendor } from "../../../connectors/http/index.ts";
import { draftConnector } from "../../../scripts/openapi-to-connector.ts";

const script = fileURLToPath(new URL("../../../scripts/openapi-to-connector.ts", import.meta.url));
const fixture = fileURLToPath(
  new URL("../../../scripts/fixtures/tracker.openapi.json", import.meta.url),
);
const spec = JSON.parse(readFileSync(fixture, "utf8"));

test("a spec's operations become a draft that loads and runs on the http connector", () => {
  const { connector: draft, notes } = draftConnector(spec, {
    id: "tracker",
    tags: ["issues", "admin"],
  });
  const connector = connectorSchema.parse(draft);
  assertToolsImplemented([
    { connector, toolset: httpTools(connector.http!), vendor: httpVendor(connector) },
  ]);
  assert.equal(connector.http!.baseUrl, "https://tracker.example.com/api/v2");
  assert.deepEqual(connector.http!.auth, {
    type: "header",
    header: "X-Tracker-Key",
    value: "api_key",
  });
  assert.deepEqual(
    connector.capabilities.map((c) => [c.id, c.destructive, c.tools]),
    [
      ["read", false, ["tracker_list_issues", "tracker_get_issue"]],
      ["write", false, ["tracker_create_issue"]],
      ["delete", true, ["tracker_delete_issue"]],
    ],
  );
  const [list, create, get] = connector.http!.tools;
  assert.deepEqual(list.params.state, {
    type: "string",
    enum: ["open", "closed"],
    default: "open",
    description: "state",
    in: "query",
    required: true,
  });
  assert.equal(list.params.labels.join, ",");
  assert.equal(list.params.page_token.name, "pageToken");
  assert.equal(list.params.page_token.required, false);
  assert.deepEqual(list.result, {
    items: "issues",
    fields: ["title", "body", "priority", "id", "state"],
  });
  // allOf is merged; the required title stays required, an object field is left out.
  assert.deepEqual(Object.keys(create.params), ["title", "body", "priority", "labels"]);
  assert.equal(create.params.title.required, true);
  assert.equal(create.params.body.required, false);
  assert.equal(get.path, "/issues/{issue_id}");
  assert.equal(get.result?.items, undefined);
  assert.deepEqual(connector.verify, {
    method: "GET",
    path: "/issues",
    label: "Tracker answers",
    show: "title",
  });
  for (const left of [/header param X-Request-Id/, /left out filter/, /body field assignee/])
    assert.ok(
      notes.some((n) => left.test(n)),
      String(left),
    );
});

test("operations are picked by id, and a missing one or another spec version is refused", () => {
  const { connector } = draftConnector(spec, { id: "tracker", operations: ["getIssue"] });
  assert.deepEqual(
    connector.http.tools.map((t) => t.name),
    ["tracker_get_issue"],
  );
  assert.throws(
    () => draftConnector(spec, { id: "tracker", operations: ["getIssue", "nope"] }),
    /No operation has the id nope/,
  );
  assert.throws(() => draftConnector(spec, { id: "tracker", tags: ["none"] }), /No operation/);
  assert.throws(
    () =>
      draftConnector(
        { ...spec, openapi: undefined, swagger: "2.0" },
        { id: "t", tags: ["issues"] },
      ),
    /OpenAPI 3/,
  );
});

test("each sign-in the spec names maps to the http connector's, or a token with a note", () => {
  const draft = (scheme: unknown) =>
    draftConnector(
      { ...spec, components: { ...spec.components, securitySchemes: { apiKey: scheme } } },
      { id: "tracker", operations: ["getIssue"] },
    );
  const cases: [unknown, unknown][] = [
    [
      { type: "apiKey", in: "query", name: "key" },
      { type: "query", param: "key", value: "api_key" },
    ],
    [
      { type: "http", scheme: "basic" },
      { type: "basic", username: "username", password: "password" },
    ],
    [
      { type: "http", scheme: "bearer" },
      { type: "bearer", token: "token" },
    ],
    [
      { type: "oauth2", flows: { clientCredentials: { tokenUrl: "/oauth/token", scopes: {} } } },
      {
        type: "oauth2",
        tokenUrl: "https://tracker.example.com/oauth/token",
        clientId: "client_id",
        clientSecret: "client_secret",
      },
    ],
  ];
  for (const [scheme, auth] of cases) {
    const { connector } = draft(scheme);
    assert.deepEqual(connector.http.auth, auth);
    connectorSchema.parse(connector);
  }
  const { connector, notes } = draft({
    type: "oauth2",
    flows: { authorizationCode: { authorizationUrl: "/a", tokenUrl: "/t", scopes: {} } },
  });
  assert.deepEqual(connector.http.auth, { type: "bearer", token: "token" });
  assert.ok(notes.some((n) => /does not map/.test(n)));
});

test("the command writes the draft and says it is one", () => {
  const dir = mkdtempSync(join(tmpdir(), "openapi-draft-"));
  try {
    const out = join(dir, "tracker", "connector.json");
    const said = execFileSync(
      process.execPath,
      ["--import", "tsx", script, fixture, "--id", "tracker", "--tags", "issues", "--out", out],
      { encoding: "utf8" },
    );
    assert.match(said, /a DRAFT with 3 tools, not a finished connector/);
    assert.doesNotMatch(said, /does not load/);
    connectorSchema.parse(JSON.parse(readFileSync(out, "utf8")));
  } finally {
    rmSync(dir, { recursive: true });
  }
});
