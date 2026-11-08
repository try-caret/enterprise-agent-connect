import assert from "node:assert/strict";
import { test } from "node:test";
import type { VendorRequest, VendorResponse } from "../../packages/connector-kit/src/types.ts";
import { harness } from "../../packages/server/test/helpers.ts";
import { scopeAql } from "./aql.ts";
import { jiraVendor } from "./vendor.ts";

// The Jira Assets tools against an in-memory Assets (the server's test harness). Object schema
// 137 holds object type 505 (Laptop) with objects 11 and 12; schema 9 holds object 99. Answer
// shapes follow Atlassian's Assets OpenAPI examples: AQL values carry only
// objectTypeAttributeId, resolved through objectTypeAttributes; GET object embeds each
// attribute's objectTypeAttribute.

const WS = "/jsm/assets/workspace/{workspaceId}/v1/";
const LAPTOP = { id: "505", name: "Laptop", objectSchemaId: "137" };
const OTHER = { id: "77", name: "Server", objectSchemaId: "9" };
const ATTRS = [
  {
    id: "1",
    name: "Key",
    type: 0,
    defaultType: { id: 0, name: "Text" },
    editable: false,
    system: true,
    maximumCardinality: 1,
    minimumCardinality: 1,
  },
  {
    id: "2",
    name: "Name",
    type: 0,
    defaultType: { id: 0, name: "Text" },
    editable: true,
    system: false,
    maximumCardinality: 1,
    minimumCardinality: 1,
  },
  {
    id: "3",
    name: "Status",
    type: 7,
    typeValueMulti: ["20", "21"],
    editable: true,
    system: false,
    maximumCardinality: 1,
    minimumCardinality: 0,
  },
  {
    id: "4",
    name: "Date Received",
    type: 0,
    defaultType: { id: 4, name: "Date" },
    editable: true,
    system: false,
    maximumCardinality: 1,
    minimumCardinality: 0,
  },
  {
    id: "5",
    name: "Recipient",
    type: 2,
    editable: true,
    system: false,
    maximumCardinality: 1,
    minimumCardinality: 0,
  },
  {
    id: "6",
    name: "Notes",
    type: 0,
    defaultType: { id: 0, name: "Text" },
    editable: true,
    system: false,
    maximumCardinality: 1,
    minimumCardinality: 0,
  },
  {
    id: "7",
    name: "Location",
    type: 0,
    defaultType: { id: 10, name: "Select" },
    options: "Boston, London",
    editable: true,
    system: false,
    maximumCardinality: 1,
    minimumCardinality: 0,
  },
].map((a) => ({ ...a, objectType: LAPTOP }));

const value = (v: string, extra: object = {}) => ({ value: v, displayValue: v, ...extra });
const laptop = (id: string, status: string, received: string, embed: boolean, type = LAPTOP) => ({
  id,
  objectKey: `IT-${id}`,
  label: `LT-${id}`,
  objectType: type,
  updated: "2026-11-01T00:00:00.000Z",
  attributes: [
    { objectTypeAttributeId: "2", objectAttributeValues: [value(`LT-${id}`)] },
    {
      objectTypeAttributeId: "3",
      objectAttributeValues: [{ displayValue: status, status: { id: "20", name: status } }],
    },
    {
      objectTypeAttributeId: "4",
      objectAttributeValues: [{ value: received, displayValue: "01/Mar/26" }],
    },
    { objectTypeAttributeId: "5", objectAttributeValues: [] },
  ].map((a) =>
    embed ? { ...a, objectTypeAttribute: ATTRS.find((d) => d.id === a.objectTypeAttributeId) } : a,
  ),
});

type Options = { fail?: number; statusRefused?: boolean };

function assets(h: ReturnType<typeof harness>, options: Options = {}) {
  const sent: VendorRequest[] = [];
  const objects: Record<string, ReturnType<typeof laptop>> = {
    "11": laptop("11", "In stock", "2026-03-01", true),
    "12": laptop("12", "In stock", "2026-01-15", true),
    "99": laptop("99", "In stock", "2026-02-01", true, OTHER),
  };
  h.vendor.call = async (_conn, req): Promise<VendorResponse> => {
    sent.push(req);
    if (options.fail)
      return { status: options.fail, data: { errorMessages: ["No access to this schema"] } };
    assert.ok(req.path.startsWith(WS), req.path);
    const [p] = req.path.slice(WS.length).split("?");
    if (p === "object/aql" && req.method === "POST")
      return {
        status: 200,
        data: {
          startAt: 0,
          maxResults: 25,
          total: 3,
          isLast: "true",
          objectTypeAttributes: ATTRS,
          values: [
            laptop("12", "In stock", "2026-01-15", false),
            laptop("11", "In stock", "2026-03-01", false),
            laptop("99", "In stock", "2026-02-01", false, OTHER),
          ],
        },
      };
    const object = /^object\/(\d+)$/.exec(p);
    if (object && req.method === "GET")
      return objects[object[1]]
        ? { status: 200, data: objects[object[1]] }
        : { status: 404, data: { errorMessages: ["No object"] } };
    if (object && req.method === "PUT") {
      // Apply only what was sent, as Assets does.
      const o = objects[object[1]];
      for (const a of (req.body as any).attributes) {
        const at = o.attributes.find(
          (x: any) => x.objectTypeAttributeId === a.objectTypeAttributeId,
        );
        const vals = a.objectAttributeValues.map((v: any) =>
          a.objectTypeAttributeId === "3"
            ? {
                displayValue: v.value === "21" ? "In use" : "?",
                status: { id: v.value, name: v.value === "21" ? "In use" : "?" },
              }
            : value(v.value),
        );
        if (at) at.objectAttributeValues = vals;
        else
          o.attributes.push({
            objectTypeAttributeId: a.objectTypeAttributeId,
            objectAttributeValues: vals,
            objectTypeAttribute: ATTRS.find((d) => d.id === a.objectTypeAttributeId),
          } as any);
      }
      return { status: 200, data: o };
    }
    if (p === "objecttype/505/attributes") return { status: 200, data: ATTRS };
    if (p === "objecttype/77/attributes")
      return { status: 200, data: ATTRS.map((a) => ({ ...a, objectType: OTHER })) };
    if (p === "config/statustype" && options.statusRefused)
      return { status: 401, data: { code: 401, message: "Unauthorized; scope does not match" } };
    if (p === "config/statustype")
      return {
        status: 200,
        data: [
          { id: "20", name: "In stock" },
          { id: "21", name: "In use" },
          { id: "22", name: "Retired" },
        ],
      };
    return { status: 404, data: { errorMessages: ["Not found"] } };
  };
  return { sent, objects };
}

const call = async (h: ReturnType<typeof harness>, key: string, tool: string, args: object) => {
  const r = await h.callTool(key, tool, args);
  return { ...r, data: r.text && !r.isError ? JSON.parse(r.text) : null };
};

test("scopeAql puts the agent's AQL inside the schema filter and keeps ORDER BY last", () => {
  assert.equal(
    scopeAql('Status = "In stock" ORDER BY "Date Received" ASC', ["137"]),
    'objectSchemaId = 137 AND (Status = "In stock") ORDER BY "Date Received" ASC',
  );
  assert.equal(
    scopeAql("objectType = Laptop", ["1", "2"]),
    "objectSchemaId in (1, 2) AND (objectType = Laptop)",
  );
  assert.equal(scopeAql("objectType = Laptop", null), "objectType = Laptop");
  // AQL has no single-quoted strings, so these parentheses count and could close the filter's.
  assert.throws(() => scopeAql("Name = 'x') OR (Name = 'y'", ["137"]), /Unbalanced/);
  assert.throws(() => scopeAql('Name = "x" ORDER BY Name; DROP', ["137"]), /ORDER BY/);
  assert.throws(() => scopeAql("ORDER BY Name", null), /AQL condition/);
});

test("jira_assets_search names attributes, sends the schema-scoped AQL and drops other schemas", async () => {
  const h = harness();
  const key = h.addAgent(["read", "assets_read"], { asset_schemas: ["137"] });
  const { sent } = assets(h);
  const r = await call(h, key, "jira_assets_search", {
    aql: 'objectType = Laptop AND Status = "In stock" ORDER BY "Date Received" ASC',
    max_results: 10,
  });
  assert.equal(r.isError, false, r.text ?? "");
  const scoped =
    'objectSchemaId = 137 AND (objectType = Laptop AND Status = "In stock") ORDER BY "Date Received" ASC';
  assert.deepEqual(sent[0], {
    method: "POST",
    path: `${WS}object/aql?startAt=0&maxResults=10&includeAttributes=true`,
    body: { qlQuery: scoped },
  });
  assert.deepEqual(
    r.data.objects.map((o: any) => o.id),
    ["12", "11"],
  );
  assert.deepEqual(r.data.objects[0], {
    id: "12",
    key: "IT-12",
    label: "LT-12",
    object_type: "Laptop",
    object_type_id: "505",
    schema_id: "137",
    updated: "2026-11-01T00:00:00.000Z",
    attributes: {
      Name: "LT-12",
      Status: "In stock",
      "Date Received": "2026-01-15",
      Recipient: null,
    },
  });
  assert.equal(r.data.more, false);
  assert.equal(h.state.audit.at(-1)!.target, scoped);
  assert.equal(h.state.audit.at(-1)!.decision, "allowed");
});

test("jira_assets_get_object refuses an object in another schema", async () => {
  const h = harness();
  const key = h.addAgent(["read", "assets_read"], { asset_schemas: ["137"] });
  assets(h);
  const ok = await call(h, key, "jira_assets_get_object", { object_id: "11" });
  assert.equal(ok.data.attributes.Status, "In stock");
  const r = await call(h, key, "jira_assets_get_object", { object_id: "99" });
  assert.equal(r.isError, true);
  assert.match(r.text!, /object schema this agent is not allowed/);
  assert.equal(h.state.audit.at(-1)!.decision, "denied");
  // Without a schema limit, every schema the account sees.
  const open = harness();
  const all = open.addAgent(["read", "assets_read"]);
  assets(open);
  assert.equal(
    (await call(open, all, "jira_assets_get_object", { object_id: "99" })).data.schema_id,
    "9",
  );
});

test("jira_assets_list_attributes lists names and kinds, and refuses another schema's type", async () => {
  const h = harness();
  const key = h.addAgent(["read", "assets_read"], { asset_schemas: ["137"] });
  assets(h);
  const r = await call(h, key, "jira_assets_list_attributes", { object_type_id: "505" });
  const byName = Object.fromEntries(r.data.map((a: any) => [a.name, a]));
  assert.equal(byName.Status.kind, "Status");
  assert.equal(byName.Recipient.kind, "User");
  assert.equal(byName["Date Received"].kind, "Date");
  assert.deepEqual(byName.Location.options, ["Boston", "London"]);
  assert.equal(byName.Key.editable, false);
  const other = await call(h, key, "jira_assets_list_attributes", { object_type_id: "77" });
  assert.match(other.text!, /not allowed/);
});

test("jira_assets_update_object sends only the named attributes, resolves a status, reads back", async () => {
  const h = harness();
  const key = h.addAgent(["read", "assets_update"], { asset_schemas: ["137"] });
  const { sent } = assets(h);
  const r = await call(h, key, "jira_assets_update_object", {
    object_id: "11",
    attributes: { status: "in use", Recipient: "acc-123", Notes: "Issued for HELP-7" },
  });
  assert.equal(r.isError, false, r.text ?? "");
  const put = sent.filter((s) => s.method === "PUT");
  assert.equal(put.length, 1);
  assert.deepEqual(put[0], {
    method: "PUT",
    path: `${WS}object/11`,
    body: {
      objectTypeId: "505",
      attributes: [
        { objectTypeAttributeId: "3", objectAttributeValues: [{ value: "21" }] },
        { objectTypeAttributeId: "5", objectAttributeValues: [{ value: "acc-123" }] },
        { objectTypeAttributeId: "6", objectAttributeValues: [{ value: "Issued for HELP-7" }] },
      ],
    },
  });
  // The statuses are read for the object's own schema.
  assert.ok(sent.some((s) => s.path === `${WS}config/statustype?objectSchemaId=137`));
  // Read back after the PUT.
  assert.equal(sent.at(-1)!.method, "GET");
  assert.equal(sent.at(-1)!.path, `${WS}object/11`);
  assert.deepEqual(r.data.updated, ["Status", "Recipient", "Notes"]);
  assert.equal(r.data.object.attributes.Status, "In use");
  assert.equal(r.data.object.attributes.Notes, "Issued for HELP-7");
  assert.equal(r.data.object.attributes["Date Received"], "2026-03-01");
  assert.equal(h.state.audit.at(-1)!.target, "11 (status, Recipient, Notes)");
});

test("jira_assets_update_object refuses before any write: other schema, unknown, read-only, bad status", async () => {
  const h = harness();
  const key = h.addAgent(["read", "assets_update"], { asset_schemas: ["137"] });
  const { sent } = assets(h);
  const cases: [object, RegExp][] = [
    [{ object_id: "99", attributes: { Notes: "x" } }, /object schema this agent is not allowed/],
    [{ object_id: "11", attributes: { Colour: "red" } }, /no attribute Colour/],
    [{ object_id: "11", attributes: { Key: "IT-1" } }, /Key cannot be edited/],
    // Retired exists but this Status attribute allows only 20 and 21.
    [{ object_id: "11", attributes: { Status: "Retired" } }, /no status Retired/],
    [{ object_id: "11", attributes: { Notes: ["a", "b"] } }, /takes one value/],
    [{ object_id: "11", attributes: {} }, /at least one/],
  ];
  for (const [args, message] of cases) {
    const r = await call(h, key, "jira_assets_update_object", args);
    assert.equal(r.isError, true);
    assert.match(r.text!, message);
    assert.equal(h.state.audit.at(-1)!.decision, "denied");
  }
  assert.equal(sent.filter((s) => s.method === "PUT").length, 0);
});

test("without the status list (service accounts get no read:cmdb-config) statuses come from the type's objects", async () => {
  const h = harness();
  const key = h.addAgent(["read", "assets_update"], { asset_schemas: ["137"] });
  const { sent } = assets(h, { statusRefused: true });
  const held = await call(h, key, "jira_assets_update_object", {
    object_id: "11",
    attributes: { Status: "In stock" },
  });
  assert.equal(held.isError, false, held.text ?? "");
  const put = sent.filter((s) => s.method === "PUT");
  assert.deepEqual((put[0].body as any).attributes, [
    { objectTypeAttributeId: "3", objectAttributeValues: [{ value: "20" }] },
  ]);
  const aql = sent.find((s) => s.method === "POST");
  assert.deepEqual(aql!.body, { qlQuery: "objectSchemaId = 137 AND objectTypeId = 505" });
  const unheld = await call(h, key, "jira_assets_update_object", {
    object_id: "11",
    attributes: { Status: "In use" },
  });
  assert.equal(unheld.isError, true);
  assert.match(unheld.text!, /no status In use\. Use one of: In stock\. Only statuses some object/);
  assert.equal(sent.filter((s) => s.method === "PUT").length, 1);
});

test("Assets 401, 403 and 404 come back as Jira errors do, 403 with the capability's hint", async () => {
  for (const status of [401, 403, 404]) {
    const h = harness();
    const key = h.addAgent(["read", "assets_read"]);
    assets(h, { fail: status });
    const r = await call(h, key, "jira_assets_get_object", { object_id: "11" });
    assert.equal(r.isError, true);
    assert.match(r.text!, new RegExp(`^Jira Assets answered ${status}: No access to this schema`));
    assert.equal(/role on the object schema/.test(r.text!), status === 403);
    const row = h.state.audit.at(-1)!;
    assert.equal(row.decision, "failed");
    assert.equal(row.vendorStatus, status);
  }
});

test("an agent with read assets only is not given the update tool", async () => {
  const h = harness();
  const key = h.addAgent(["read", "assets_read"]);
  const names = ((await h.rpc(key, "tools/list")).body.result.tools as { name: string }[]).map(
    (t) => t.name,
  );
  assert.ok(names.includes("jira_assets_search"));
  assert.ok(!names.includes("jira_assets_update_object"));
  const issuesOnly = harness();
  const plain = issuesOnly.addAgent(["read"]);
  const listed = (
    (await issuesOnly.rpc(plain, "tools/list")).body.result.tools as { name: string }[]
  ).map((t) => t.name);
  assert.ok(!listed.some((n) => n.startsWith("jira_assets_")));
});

test("the vendor finds the Assets workspace once per site and fills it into Assets paths", async () => {
  const CLOUD = "11111111-2222-3333-4444-555555555555";
  const urls: string[] = [];
  let workspaces: unknown = { values: [{ workspaceId: "1060ba0e-178b-4e0e-g0h1-jedb02cccb5f" }] };
  const fake = (async (input: string | URL | Request) => {
    const url = String(input);
    urls.push(url);
    if (url === "https://auth.atlassian.com/oauth/token")
      return Response.json({ access_token: "t", expires_in: 3600 });
    if (url.endsWith("/rest/servicedeskapi/assets/workspace"))
      return typeof workspaces === "number"
        ? Response.json({ errorMessages: ["no"] }, { status: workspaces })
        : Response.json(workspaces);
    return Response.json({ ok: url });
  }) as typeof fetch;
  const conn = { config: { cloudId: CLOUD }, secret: { client_id: "id", client_secret: "s" } };

  // Discovery refused: its answer is passed back and nothing is kept.
  workspaces = 401;
  const refused = await jiraVendor(fake).call(conn, { method: "GET", path: `${WS}object/1` });
  assert.equal(refused.status, 401);
  // A site without Assets.
  workspaces = { values: [] };
  const none = await jiraVendor(fake).call(conn, { method: "GET", path: `${WS}object/1` });
  assert.equal(none.status, 404);
  assert.match(JSON.stringify(none.data), /no Assets workspace/);

  workspaces = { values: [{ workspaceId: "1060ba0e-178b-4e0e-g0h1-jedb02cccb5f" }] };
  const vendor = jiraVendor(fake);
  urls.length = 0;
  await vendor.call(conn, { method: "GET", path: `${WS}object/1` });
  await vendor.call(conn, { method: "POST", path: `${WS}object/aql`, body: { qlQuery: "x" } });
  await vendor.call(conn, { method: "GET", path: "/rest/api/3/myself" });
  const base = `https://api.atlassian.com/ex/jira/${CLOUD}`;
  assert.deepEqual(
    urls.filter((u) => !u.startsWith("https://auth.")),
    [
      `${base}/rest/servicedeskapi/assets/workspace`,
      `${base}/jsm/assets/workspace/1060ba0e-178b-4e0e-g0h1-jedb02cccb5f/v1/object/1`,
      `${base}/jsm/assets/workspace/1060ba0e-178b-4e0e-g0h1-jedb02cccb5f/v1/object/aql`,
      `${base}/rest/api/3/myself`,
    ],
  );
});
