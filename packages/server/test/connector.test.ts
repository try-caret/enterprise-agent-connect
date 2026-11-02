import assert from "node:assert/strict";
import { cpSync, mkdirSync, mkdtempSync, symlinkSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { pathToFileURL } from "node:url";
import { test } from "node:test";
import {
  connectorSchema,
  effectiveCapabilities,
  expandCopy,
  loadConnectors,
  loadProviders,
  pickDisplay,
  pickLabels,
  providerSettings,
} from "../../connector-kit/src/connector.ts";
import { assertToolsImplemented, InvalidInput, validateCapabilities } from "../src/service.ts";
import { graphTools } from "../../../connectors/sharepoint-files/tools.ts";
import type { Toolset } from "../../connector-kit/src/common.ts";
import { jiraTools } from "../../../connectors/jira/tools.ts";
import { jira } from "./helpers.ts";

test("every shipped connector parses and its tools exist", () => {
  const all = loadConnectors();
  assert.ok(all.some((c) => c.id === "jira"));
  assertToolsImplemented([{ connector: jira, toolset: jiraTools, vendor: {} as never }]);
});

test("a connector naming a tool nobody implements fails at startup", () => {
  const broken = {
    ...jira,
    capabilities: [
      ...jira.capabilities,
      { ...jira.capabilities[1], id: "x", tools: ["jira_delete_everything"] },
    ],
  };
  assert.throws(
    () => assertToolsImplemented([{ connector: broken, toolset: jiraTools, vendor: {} as never }]),
    /jira_delete_everything/,
  );
});

test("an invalid connector file is refused with its name", () => {
  const dir = mkdtempSync(join(tmpdir(), "connectors-"));
  mkdirSync(join(dir, "bad"));
  writeFileSync(join(dir, "bad", "connector.json"), JSON.stringify({ id: "bad" }));
  assert.throws(() => loadConnectors(pathToFileURL(dir + "/")), /bad\/connector\.json/);
});

test("only the folders asked for are read, so an app not offered is never parsed", () => {
  const dir = mkdtempSync(join(tmpdir(), "connectors-"));
  mkdirSync(join(dir, "bad"));
  writeFileSync(join(dir, "bad", "connector.json"), JSON.stringify({ id: "bad" }));
  assert.deepEqual(
    loadConnectors(pathToFileURL(dir + "/"), (id) => id !== "bad"),
    [],
  );
  assert.deepEqual(
    loadConnectors(undefined, (id) => id === "jira").map((c) => c.id),
    ["jira"],
  );
});

test("a connector folder may be a link, as in a Kubernetes ConfigMap mount", () => {
  const dir = mkdtempSync(join(tmpdir(), "connectors-"));
  const orders = new URL("../../../examples/internal-app/connectors/orders", import.meta.url);
  cpSync(orders, join(dir, "..data", "orders"), { recursive: true });
  symlinkSync(join(dir, "..data", "orders"), join(dir, "orders"));
  assert.deepEqual(
    loadConnectors(pathToFileURL(dir + "/")).map((c) => c.id),
    ["orders"],
  );
});

test("a toolset tool that no capability names fails at startup", () => {
  const partial = { ...jira, capabilities: jira.capabilities.filter((c) => c.id !== "delete") };
  assert.throws(
    () => assertToolsImplemented([{ connector: partial, toolset: jiraTools, vendor: {} as never }]),
    /jira_delete_issue/,
  );
});

test("a capability with a tool that can destroy or overwrite must be destructive", () => {
  const toolset: Toolset = {
    names: ["fake_read", "fake_wipe"],
    register(server, _ctx, allowed) {
      for (const name of this.names)
        if (allowed.has(name))
          server.registerTool(
            name,
            { description: name, annotations: { destructiveHint: name === "fake_wipe" } },
            async () => ({ content: [] }),
          );
    },
  };
  const fake = (destructive: boolean) => ({
    ...jira,
    id: "fake",
    capabilities: [
      { ...jira.capabilities[0], id: "read", tools: ["fake_read"] },
      { ...jira.capabilities[0], id: "wipe", tools: ["fake_wipe"], required: false, destructive },
    ],
  });
  const load = (destructive: boolean) =>
    assertToolsImplemented([{ connector: fake(destructive), toolset, vendor: {} as never }]);
  assert.throws(() => load(false), /fake: capability wipe has fake_wipe/);
  load(true);
});

test("effective capabilities are the agent's ticks that the connection also allows", () => {
  const ids = (grant: string[], ceiling: string[]) =>
    effectiveCapabilities(jira, grant, ceiling).map((c) => c.id);
  assert.deepEqual(ids(["read", "comment", "delete"], ["read", "comment"]), ["read", "comment"]);
  assert.deepEqual(ids(["read"], ["read", "comment", "delete"]), ["read"]);
  assert.deepEqual(ids(["delete", "read"], ["delete", "read"]), ["read", "delete"]);
  assert.deepEqual(ids(["nonsense"], ["nonsense"]), []);
});

test("ticked capabilities are checked, ordered and always include the required ones", () => {
  assert.deepEqual(validateCapabilities(jira, ["delete", "comment"]), [
    "read",
    "comment",
    "delete",
  ]);
  assert.deepEqual(validateCapabilities(jira, []), ["read"]);
  assert.throws(() => validateCapabilities(jira, ["admin"]), InvalidInput);
  assert.throws(() => validateCapabilities(jira, "read"), InvalidInput);
  // Without a required capability, an empty choice is refused rather than saved.
  const teams = loadConnectors().find((c) => c.id === "teams")!;
  assert.throws(() => validateCapabilities(teams, []), /at least one/);
  assert.deepEqual(validateCapabilities(teams, ["respond"]), ["respond"]);
});

test("connector capabilities are validated", () => {
  const bad = (capabilities: unknown) =>
    connectorSchema.safeParse({ ...jira, capabilities }).success;
  assert.equal(bad(jira.capabilities), true);
  // A tool in two capabilities, a required destructive one.
  assert.equal(bad([...jira.capabilities, { ...jira.capabilities[1], id: "again" }]), false);
  assert.equal(bad([{ ...jira.capabilities[0], destructive: true }]), false);
  // Spending money is asked about first too.
  assert.equal(bad([jira.capabilities[0], { ...jira.capabilities[1], costs: true }]), false);
  assert.equal(
    bad([jira.capabilities[0], { ...jira.capabilities[1], costs: true, destructive: true }]),
    true,
  );
});

test("a probe naming a vendor check the client lacks fails at startup", () => {
  const graph = loadConnectors().find((c) => c.id === "sharepoint-files")!;
  assert.throws(
    () => assertToolsImplemented([{ connector: graph, toolset: graphTools, vendor: {} as never }]),
    /check not implemented: permissions/,
  );
});

test("copy templates expand scopes", () => {
  assert.equal(expandCopy(["{scopes}"], ["a", "b"]), "a b");
  assert.equal(expandCopy(["head", "role {scope}"], ["a", "b"]), "head\nrole a\nrole b");
});

test("display paths read nested values", () => {
  assert.deepEqual(pickDisplay({ values: [{ key: "A" }, { key: "B" }] }, "values[].key"), [
    "A",
    "B",
  ]);
  assert.deepEqual(pickDisplay({ baseUrl: "u" }, "baseUrl"), ["u"]);
});

test("a probe's name labels each value from the same list item", () => {
  const data = { items: [{ id: 1, name: "US" }, { id: 2 }, { id: 3, name: "UK", inactive: true }] };
  assert.deepEqual(pickLabels(data, "items[].id", "items[].name"), { "1": "US", "3": "UK" });
  assert.deepEqual(pickLabels(data, "items[].id", "items[].name", { inactive: true }), {
    "3": "UK",
  });
  const netsuite = loadConnectors().find((c) => c.id === "netsuite")!;
  const discover = { ...netsuite.discover!, name: "rows[].name" };
  assert.match(
    String(connectorSchema.safeParse({ ...netsuite, discover }).error),
    /same list as its show/,
  );
});

test("a step's capabilities and a field's optional are checked and kept", () => {
  const netsuite = loadConnectors().find((c) => c.id === "netsuite")!;
  assert.deepEqual(netsuite.afterConnect.at(-1)!.capabilities, ["saved_searches"]);
  const afterConnect = [{ ...netsuite.afterConnect[0], capabilities: ["nope"] }];
  assert.match(
    String(connectorSchema.safeParse({ ...netsuite, afterConnect }).error),
    /capabilities name capabilities/,
  );
  const okta = loadConnectors().find((c) => c.id === "okta")!;
  assert.equal(okta.fields.find((f) => f.id === "key_id")!.optional, true);
  assert.equal(okta.fields.find((f) => f.id === "org")!.optional, false);
});

test("a map limit is keyed by a list limit listed before it", () => {
  const graph = loadConnectors().find((c) => c.id === "sharepoint-files")!;
  const [sites, folders] = graph.limits;
  const parse = (limits: unknown[]) => connectorSchema.safeParse({ ...graph, limits }).success;
  assert.equal(parse([sites, folders]), true);
  assert.equal(parse([folders, sites]), false);
  assert.equal(parse([sites, { ...folders, keys: undefined }]), false);
  assert.equal(parse([sites, { ...folders, keys: "nope" }]), false);
  assert.equal(parse([sites, { ...sites, id: "more", keys: "sites" }]), false);
});

test("an app with sites to add needs a client that checks them", () => {
  const graph = loadConnectors().find((c) => c.id === "sharepoint-files")!;
  const checks = { permissions: async () => ({ values: [] }), sites: async () => ({ values: [] }) };
  assert.throws(
    () =>
      assertToolsImplemented([
        { connector: graph, toolset: graphTools, vendor: { checks } as never },
      ]),
    /cannot check them/,
  );
});

test("addresses and sites are filled into commands only inside single quotes", () => {
  const steps = [
    ...loadConnectors().flatMap((c) => [
      ...c.setup,
      ...c.afterConnect,
      ...(c.siteAccess?.grant ?? []),
    ]),
    ...loadProviders().flatMap((p) => p.addAgent ?? []),
  ];
  let seen = 0;
  for (const line of steps.flatMap((s) => s.copy ?? []))
    for (const m of line.matchAll(/\{(agentEmail|agentName|siteRef)\}/g)) {
      seen++;
      const at = m.index!;
      assert.equal(line[at - 1], "'", line);
      assert.equal(line[at + m[0].length], "'", line);
    }
  assert.ok(seen >= 5);
});

test("settings are a provider's and may not take the connection's own keys", () => {
  for (const c of loadConnectors())
    for (const s of [...c.setup, ...c.afterConnect]) if (s.settings) assert.ok(c.provider, c.id);
  // A setting is saved beside the connection's own config, so it may not take one of its keys.
  const all = loadConnectors();
  for (const p of loadProviders()) providerSettings(all, p.id);
  const teams = all.find((c) => c.id === "teams")!;
  for (const id of ["tenantId", "clientId", "auth", "certificate", "sites"]) {
    const bad = { ...teams, setup: [{ ...teams.setup[0], settings: [{ id, label: "x" }] }] };
    assert.throws(() => providerSettings([bad as never], "microsoft"), /the connection's own/);
  }
});

test("a step's option and a field's requires must name an earlier choice's option", () => {
  const teams = loadConnectors().find((c) => c.id === "teams")!;
  const refused = (c: unknown) =>
    assert.match(String(connectorSchema.safeParse(c).error), /earlier choice/);
  refused({ ...teams, setup: [{ ...teams.setup[0], option: { field: "nope", value: "on" } }] });
  const [adds, person] = teams.fields;
  refused({ ...teams, fields: [person, adds] });
  assert.ok(connectorSchema.safeParse(teams).success);
});
