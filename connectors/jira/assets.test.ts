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
    options: "Boston, London,Needham, MA",
    editable: true,
    system: false,
    maximumCardinality: 1,
    minimumCardinality: 0,
  },
  {
    id: "9",
    name: "Condition",
    type: 7,
    typeValueMulti: ["30", "31", "32"],
    editable: true,
    system: false,
    maximumCardinality: 1,
    minimumCardinality: 0,
  },
  {
    id: "8",
    name: "Owner",
    type: 1,
    editable: true,
    system: false,
    maximumCardinality: 1,
    minimumCardinality: 0,
  },
].map((a) => ({ ...a, objectType: LAPTOP }));

const value = (v: string, extra: object = {}) => ({ value: v, displayValue: v, ...extra });
// Status ids by attribute and name, as objects hold them.
const STATUSES: Record<string, Record<string, string>> = {
  "3": { "In stock": "20", "In use": "21", Retired: "22" },
  "9": { Good: "30", Damaged: "31", "2024": "32" },
};
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
      objectAttributeValues: [
        { displayValue: status, status: { id: STATUSES["3"][status] ?? "20", name: status } },
      ],
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

// held: the status names some object of the type holds (default all); putEmpty: the PUT answers
// no object, so the tool reads it back; readBackFail: the GET
// after a PUT answers that status.
type Options = { fail?: number; held?: string[]; putEmpty?: boolean; readBackFail?: number };

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
    // A status name looked up on an object of the type that holds it.
    const lookup = /AND "([^"]+)" = "([^"]+)"$/.exec(String((req.body as any)?.qlQuery ?? ""));
    if (p === "object/aql" && req.method === "POST" && lookup) {
      const def = ATTRS.find((a) => a.name === lookup[1])!;
      const name = Object.keys(STATUSES[def.id]).find(
        (n) => n.toLowerCase() === lookup[2].toLowerCase(),
      );
      const held = !name || (options.held && !options.held.includes(name)) ? [] : [name];
      return {
        status: 200,
        data: {
          isLast: true,
          values: held.map((n) => ({
            id: "50",
            objectType: LAPTOP,
            attributes: [
              {
                objectTypeAttributeId: def.id,
                objectAttributeValues: [{ status: { id: STATUSES[def.id][n], name: n } }],
              },
            ],
          })),
        },
      };
    }
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
    if (object && req.method === "GET" && options.readBackFail && sent.some((s) => s.method === "PUT"))
      return { status: options.readBackFail, data: { errorMessages: ["Try again later"] } };
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
        const vals = a.objectAttributeValues.map((v: any) => {
          const names = STATUSES[a.objectTypeAttributeId];
          if (!names) return value(v.value);
          const n = Object.keys(names).find((k) => names[k] === v.value) ?? "?";
          return { displayValue: n, status: { id: v.value, name: n } };
        });
        if (at) at.objectAttributeValues = vals;
        else
          o.attributes.push({
            objectTypeAttributeId: a.objectTypeAttributeId,
            objectAttributeValues: vals,
            objectTypeAttribute: ATTRS.find((d) => d.id === a.objectTypeAttributeId),
          } as any);
      }
      return { status: 200, data: options.putEmpty ? null : o };
    }
    if (p === "objecttype/505/attributes") return { status: 200, data: ATTRS };
    if (p === "objecttype/77/attributes")
      return { status: 200, data: ATTRS.map((a) => ({ ...a, objectType: OTHER })) };
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
  assert.match(r.text!, /not in an object schema this agent may use/);
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
  assert.equal(byName.Location.options, "Boston, London,Needham, MA");
  assert.equal(byName.Key.editable, false);
  const other = await call(h, key, "jira_assets_list_attributes", { object_type_id: "77" });
  assert.match(other.text!, /not in an object schema this agent may use/);
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
  // The status id is looked up on an object of the type, in the object's own schema.
  assert.deepEqual(sent.find((s) => s.method === "POST")!.body, {
    qlQuery: 'objectSchemaId = 137 AND objectTypeId = 505 AND "Status" = "in use"',
  });
  // The PUT's own answer is the object shown: no second read.
  assert.equal(sent.at(-1)!.method, "PUT");
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
    [{ object_id: "99", attributes: { Notes: "x" } }, /not in an object schema this agent may use/],
    [{ object_id: "11", attributes: { Colour: "red" } }, /no attribute Colour/],
    [{ object_id: "11", attributes: { Key: "IT-1" } }, /Key cannot be edited/],
    // Retired exists but this Status attribute allows only 20 and 21.
    [{ object_id: "11", attributes: { Status: "Retired" } }, /does not allow status Retired/],
    // Not an allowed id and no status of that name: refused, never sent as an id.
    [{ object_id: "11", attributes: { Status: "22" } }, /no status 22 on any object/],
    [{ object_id: "11", attributes: { Status: 'In "use' } }, /cannot be looked up by name/],
    // An empty list is not a way to clear: only null is.
    [{ object_id: "11", attributes: { Notes: [] } }, />=1 items/],
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

test("a status no object holds yet is refused by name and taken by id; each Status attribute on its own", async () => {
  const h = harness();
  const key = h.addAgent(["read", "assets_update"], { asset_schemas: ["137"] });
  const { sent } = assets(h, { held: ["In stock", "Damaged"] });
  // A status whose name is digits is found by name when the digits are not an allowed id.
  const named = harness();
  const namedKey = named.addAgent(["read", "assets_update"], { asset_schemas: ["137"] });
  const { sent: namedSent } = assets(named);
  const byName = await call(named, namedKey, "jira_assets_update_object", {
    object_id: "11",
    attributes: { Condition: "2024" },
  });
  assert.equal(byName.isError, false, byName.text ?? "");
  assert.deepEqual((namedSent.find((s) => s.method === "PUT")!.body as any).attributes, [
    { objectTypeAttributeId: "9", objectAttributeValues: [{ value: "32" }] },
  ]);
  const unheld = await call(h, key, "jira_assets_update_object", {
    object_id: "11",
    attributes: { Status: "In use" },
  });
  assert.equal(unheld.isError, true);
  assert.match(unheld.text!, /no status In use on any object.*Ask a person to set it in Assets/);
  // No ids are offered to guess from.
  assert.doesNotMatch(unheld.text!, /20|21/);
  const byId = await call(h, key, "jira_assets_update_object", {
    object_id: "11",
    attributes: { Status: "21", Condition: "damaged" },
  });
  assert.equal(byId.isError, false, byId.text ?? "");
  assert.deepEqual((sent.find((s) => s.method === "PUT")!.body as any).attributes, [
    { objectTypeAttributeId: "3", objectAttributeValues: [{ value: "21" }] },
    { objectTypeAttributeId: "9", objectAttributeValues: [{ value: "31" }] },
  ]);
  // Condition's name was looked up as Condition, not with Status's list.
  assert.ok(
    sent.some((s) => String((s.body as any)?.qlQuery).endsWith('AND "Condition" = "damaged"')),
  );
  // The status list (read:cmdb-config) is never asked for.
  assert.ok(!sent.some((s) => s.path.includes("config/statustype")));
});

test("a failed read-back after the change answers the change, with a warning, audited as allowed", async () => {
  const h = harness();
  const key = h.addAgent(["read", "assets_update"], { asset_schemas: ["137"] });
  const { sent } = assets(h, { putEmpty: true, readBackFail: 503 });
  const r = await call(h, key, "jira_assets_update_object", {
    object_id: "11",
    attributes: { Notes: "x" },
  });
  assert.equal(r.isError, false, r.text ?? "");
  assert.deepEqual(r.data.updated, ["Notes"]);
  assert.equal(r.data.object, null);
  assert.match(r.data.warning, /saved, but reading the object back failed: do not repeat it/);
  assert.equal(h.state.audit.at(-1)!.decision, "allowed");
  assert.equal(sent.filter((s) => s.method === "PUT").length, 1);
});

test("more than 20 attributes are refused before the tool, with a short audit target", async () => {
  const h = harness();
  const key = h.addAgent(["read", "assets_update"], { asset_schemas: ["137"] });
  const { sent } = assets(h);
  const attributes = Object.fromEntries(
    Array.from({ length: 500 }, (_, i) => [`${"x".repeat(190)}${i}`, "v"]),
  );
  const r = await call(h, key, "jira_assets_update_object", { object_id: "11", attributes });
  assert.equal(r.isError, true);
  assert.equal(sent.length, 0);
  assert.ok(String(h.state.audit.at(-1)!.target ?? "").length <= 400);
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
  // The refusal names the seat and scope, not the schema role the capability's hint is about.
  assert.match(JSON.stringify(refused.data), /refused the Assets workspace lookup.*agent seat/);
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

test("scopeAql refuses, for a limited agent, AQL that looks through references or uses single quotes", () => {
  for (const [aql, message] of [
    ['object HAVING outboundReferences(objectSchemaId = 5 AND "Cost" > 5000)', /References cannot/],
    ["objectType = Laptop AND inboundReferences(Name = x)", /References cannot/],
    ['Owner.Department = "IT"', /dot notation/],
    ['"Owner"."Name" = "Kim"', /dot notation/],
    ["Name = '(' ) OR (objectSchemaId = 5 OR Name = ')'", /single quotes|Unbalanced/],
    ["Name = 'x' AND objectType = Laptop", /single quotes/],
    // A digit-led name before a quoted one, and dotted names inside quotes.
    ['1."Name" = "Alice"', /dot notation/],
    ['"Owner.Name" = "Alice"', /dot notation/],
    ['"Owner.Name" LIKE "Al"', /dot notation/],
    ['objectType = Laptop ORDER BY "Owner.Name"', /dot notation/],
    ["objectType = Laptop ORDER BY Owner.Name", /dot notation/],
    ['Owner.2FA = "a"', /dot notation/],
    ["Owner1.Name = x", /dot notation/],
    ['connectedTickets(project = SECRET AND summary ~ "merger")', /Connected tickets/],
    [
      'Name = "a\\\n" AND object HAVING outboundReferences(objectSchemaId = 5 AND Name = "x")',
      /backslash/,
    ],
    ['2024.Name = "x"', /dot notation/],
    ['Owner-1.Name = "x"', /dot notation/],
    ['Bürö2.Name = "x"', /dot notation/],
    ['Owner .2FA = "x"', /dot notation/],
    ['Owner .2 = "x"', /dot notation/],
    ['"Owner" .2 = "x"', /dot notation/],
    ['Owner$.5 = "x"', /dot notation/],
    ['Owne\u0301.2 = "x"', /dot notation/],
  ] as [string, RegExp][])
    assert.throws(() => scopeAql(aql, ["137"]), message, aql);
  // Inside double quotes the same characters are plain text; numbers keep their decimal point.
  assert.equal(
    scopeAql('Notes = "O\'Brien, see HAVING x.y" AND Cost > 3.5 AND Rate < .5 AND Gain > -1.5 AND Size IN (1.5,2.)', ["137"]),
    'objectSchemaId = 137 AND (Notes = "O\'Brien, see HAVING x.y" AND Cost > 3.5 AND Rate < .5 AND Gain > -1.5 AND Size IN (1.5,2.))',
  );
});

test("jira_assets_update_object checks Select options, refuses references and an attribute named twice", async () => {
  const h = harness();
  const key = h.addAgent(["read", "assets_update"], { asset_schemas: ["137"] });
  const { sent } = assets(h);
  const refused: [object, RegExp][] = [
    [{ Location: "Paris" }, /no option Paris. Use one of: Boston, London,Needham, MA/],
    [{ Owner: "IT-12" }, /Owner links to other objects/],
    [{ Status: "In use", status: "In use" }, /Status is named twice/],
  ];
  for (const [attributes, message] of refused) {
    const r = await call(h, key, "jira_assets_update_object", { object_id: "11", attributes });
    assert.equal(r.isError, true, JSON.stringify(attributes));
    assert.match(r.text!, message);
  }
  assert.equal(sent.filter((s) => s.method === "PUT").length, 0);
  const r = await call(h, key, "jira_assets_update_object", {
    object_id: "11",
    attributes: { Location: "boston" },
  });
  assert.equal(r.isError, false, r.text ?? "");
  assert.deepEqual((sent.find((s) => s.method === "PUT")!.body as any).attributes, [
    { objectTypeAttributeId: "7", objectAttributeValues: [{ value: "Boston" }] },
  ]);
  // An option holding a comma ("Needham, MA") is matched whole.
  const comma = await call(h, key, "jira_assets_update_object", {
    object_id: "11",
    attributes: { Location: "needham, ma" },
  });
  assert.equal(comma.isError, false, comma.text ?? "");
  assert.deepEqual((sent.filter((s) => s.method === "PUT").at(-1)!.body as any).attributes, [
    { objectTypeAttributeId: "7", objectAttributeValues: [{ value: "Needham, MA" }] },
  ]);
});

test("a referenced object outside the agent's schemas is not named", async () => {
  const h = harness();
  const key = h.addAgent(["read", "assets_read"], { asset_schemas: ["137"] });
  const { objects } = assets(h);
  const ref = (id: string, type: typeof LAPTOP) => ({
    displayValue: `Owner ${id}`,
    referencedObject: { id, label: `Owner ${id}`, objectType: type },
  });
  objects["11"].attributes.push({
    objectTypeAttributeId: "8",
    objectAttributeValues: [
      ref("12", LAPTOP),
      ref("99", OTHER),
      // referencedObject is optional in the OpenAPI: a reference without it is not named either.
      { displayValue: "Secret server", referencedType: true },
      { displayValue: "Secret too" },
    ],
    objectTypeAttribute: ATTRS.find((d) => d.id === "8"),
  } as any);
  const r = await call(h, key, "jira_assets_get_object", { object_id: "11" });
  assert.equal(r.isError, false, r.text ?? "");
  // With no schema limit a reference value carrying nothing but its attribute's type is null.
  const open = harness();
  const all = open.addAgent(["read", "assets_read"]);
  const { objects: openObjects } = assets(open);
  openObjects["11"].attributes.push({
    objectTypeAttributeId: "8",
    objectAttributeValues: [{}],
    objectTypeAttribute: ATTRS.find((d) => d.id === "8"),
  } as any);
  const bare = await call(open, all, "jira_assets_get_object", { object_id: "11" });
  assert.equal(bare.isError, false, bare.text ?? "");
  assert.equal(bare.data.attributes.Owner, null);
  const hidden = "(an object in a schema this agent may not use)";
  // The last value has no reference marks at all: Owner's definition (type 1) hides it.
  assert.deepEqual(r.data.attributes.Owner, ["Owner 12", hidden, hidden, hidden]);
});

test("scopeAql: a projects limit alone refuses connectedTickets; only \\\" may follow a backslash", () => {
  const BS = "\\";
  assert.throws(
    () => scopeAql("object HAVING connectedTickets(project = SECRET)", null, true),
    /Connected tickets/,
  );
  // To this check every parenthesis below sits in a string; if Atlassian does not read \\ as
  // an escaped backslash they would be live, so a backslash before anything but " is refused.
  const split = `Name = "a${BS}${BS}" AND Notes = ") OR objectSchemaId = 5 OR (Name = "x${BS}${BS}" y"`;
  assert.throws(() => scopeAql(split, ["137"]), /backslash/);
  assert.throws(() => scopeAql(`Name = "a${BS}${BS}"`, null, true), /backslash/);
  assert.equal(
    scopeAql(`Name = "say ${BS}"hi${BS}""`, ["137"]),
    `objectSchemaId = 137 AND (Name = "say ${BS}"hi${BS}"")`,
  );
  // With no limit at all, nothing is refused here.
  assert.equal(
    scopeAql("object HAVING connectedTickets(project = X)", null),
    "object HAVING connectedTickets(project = X)",
  );
});

test("search: an empty page that is not marked last has no next start", async () => {
  const h = harness();
  const key = h.addAgent(["read", "assets_read"], { asset_schemas: ["137"] });
  assets(h);
  const prev = h.vendor.call;
  h.vendor.call = async (conn, req) =>
    req.path.includes("object/aql")
      ? { status: 200, data: { values: [], isLast: false, total: 3 } }
      : prev(conn, req);
  const r = await call(h, key, "jira_assets_search", { aql: "objectType = Laptop", start_at: 40 });
  assert.equal(r.isError, false, r.text ?? "");
  assert.equal(r.data.more, false);
  assert.equal(r.data.next_start_at, null);
});

test("search: a projects limit refuses connectedTickets even with no schema limit", async () => {
  const h = harness();
  const key = h.addAgent(["read", "assets_read"], { projects: ["HELP"] });
  const { sent } = assets(h);
  const r = await call(h, key, "jira_assets_search", {
    aql: 'object HAVING connectedTickets(project = SECRET AND summary ~ "merger")',
  });
  assert.equal(r.isError, true);
  assert.match(r.text!, /Connected tickets/);
  assert.equal(sent.length, 0);
});

test("scopeAql: quoted values keep their dots; a projects limit alone refuses single quotes", () => {
  assert.equal(
    scopeAql('Vendor = "Dell Inc." AND Version IN ("1.2", "1.3")', ["137"]),
    'objectSchemaId = 137 AND (Vendor = "Dell Inc." AND Version IN ("1.2", "1.3"))',
  );
  // Single quotes could hide connectedTickets( from a check that reads only double quotes.
  assert.throws(
    () =>
      scopeAql(
        `Name = 'x"' OR object HAVING connectedTickets(project = SECRET) OR Name = '"'`,
        null,
        true,
      ),
    /single quotes/,
  );
});

test("under a schema limit a missing object or type gets the refusal another schema's gets", async () => {
  const h = harness();
  const key = h.addAgent(["read", "assets_read"], { asset_schemas: ["137"] });
  assets(h);
  const missing = await call(h, key, "jira_assets_get_object", { object_id: "555" });
  const other = await call(h, key, "jira_assets_get_object", { object_id: "99" });
  assert.equal(missing.text, "Assets object 555 is not in an object schema this agent may use");
  assert.equal(other.text, "Assets object 99 is not in an object schema this agent may use");
  const noType = await call(h, key, "jira_assets_list_attributes", { object_type_id: "404" });
  assert.equal(noType.text, "Object type 404 is not in an object schema this agent may use");
  // With no limit there is nothing to hide: Atlassian's 404 comes back.
  const open = harness();
  const all = open.addAgent(["read", "assets_read"]);
  assets(open);
  const plain = await call(open, all, "jira_assets_get_object", { object_id: "555" });
  assert.match(plain.text!, /^Jira Assets answered 404/);
  // Ids that are not numbers are refused before any call.
  const bad = await call(h, key, "jira_assets_get_object", { object_id: "12a" });
  assert.match(bad.text!, /Give an Assets object id/);
  const badType = await call(h, key, "jira_assets_list_attributes", { object_type_id: "x" });
  assert.match(badType.text!, /Give an object type id/);
});

test("a reference attribute known only by referenceObjectTypeId is still a reference", async () => {
  const h = harness();
  const key = h.addAgent(["read", "assets_read", "assets_update"], { asset_schemas: ["137"] });
  const { sent } = assets(h);
  // As in the OpenAPI's own example: no `type`, only referenceObjectTypeId.
  ATTRS.push({
    id: "10",
    name: "City",
    referenceObjectTypeId: "24",
    editable: true,
    system: false,
    maximumCardinality: 1,
    minimumCardinality: 0,
    objectType: LAPTOP,
  } as any);
  try {
    const r = await call(h, key, "jira_assets_update_object", {
      object_id: "11",
      attributes: { City: "ITSM-87" },
    });
    assert.match(r.text!, /City links to other objects/);
    assert.equal(sent.filter((s) => s.method === "PUT").length, 0);
    const list = await call(h, key, "jira_assets_list_attributes", { object_type_id: "505" });
    assert.equal(list.data.find((a: any) => a.name === "City").kind, "Object reference");
  } finally {
    ATTRS.pop();
  }
});

test("a status attribute whose name holds a quote is not looked up by name", async () => {
  const h = harness();
  const key = h.addAgent(["read", "assets_update"], { asset_schemas: ["137"] });
  const { sent } = assets(h);
  ATTRS.push({
    id: "11",
    name: 'Odd "state"',
    type: 7,
    typeValueMulti: ["20"],
    editable: true,
    system: false,
    maximumCardinality: 1,
    minimumCardinality: 0,
    objectType: LAPTOP,
  } as any);
  try {
    const r = await call(h, key, "jira_assets_update_object", {
      object_id: "11",
      attributes: { 'Odd "state"': "In stock" },
    });
    assert.match(r.text!, /cannot be looked up by name/);
    assert.ok(!sent.some((s) => s.method === "POST"));
  } finally {
    ATTRS.pop();
  }
});

test("a User value reads as its name and account id", async () => {
  const h = harness();
  const key = h.addAgent(["read", "assets_read"], { asset_schemas: ["137"] });
  const { objects } = assets(h);
  objects["11"].attributes.push({
    objectTypeAttributeId: "5",
    objectAttributeValues: [
      { value: "acc-9", displayValue: "Kim Lee", user: { key: "acc-9", displayName: "Kim Lee" } },
    ],
    objectTypeAttribute: ATTRS.find((d) => d.id === "5"),
  } as any);
  objects["11"].attributes = objects["11"].attributes.filter(
    (a: any, i: number, all: any[]) =>
      a.objectTypeAttributeId !== "5" ||
      i === all.map((x) => x.objectTypeAttributeId).lastIndexOf("5"),
  );
  const r = await call(h, key, "jira_assets_get_object", { object_id: "11" });
  assert.deepEqual(r.data.attributes.Recipient, { name: "Kim Lee", account_id: "acc-9" });
});
