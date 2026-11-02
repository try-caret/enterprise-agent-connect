import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { createServer, type IncomingMessage, type ServerResponse } from "node:http";
import type { AddressInfo } from "node:net";
import { after, before, test } from "node:test";
import type { McpServer } from "@modelcontextprotocol/server";
import { connectorSchema, type Connector } from "../../packages/connector-kit/src/connector.ts";
import type { ToolContext } from "../../packages/connector-kit/src/common.ts";
import { assertToolsImplemented } from "../../packages/server/src/service.ts";
import soapModule, { soapTools, soapVendor } from "./index.ts";
import { parseXml, toJson, XmlError } from "./xml.ts";

// The generic soap module, run against a local server standing in for a SOAP service.

const example = JSON.parse(
  readFileSync(
    new URL("../../examples/legacy-apps/connectors/hr-soap/connector.json", import.meta.url),
    "utf8",
  ),
);

type Seen = { url: string; headers: IncomingMessage["headers"]; body: string };
const seen: Seen[] = [];
let respond: (req: IncomingMessage, res: ServerResponse) => void;
let origin = "";
const server = createServer(async (req, res) => {
  let body = "";
  for await (const chunk of req) body += chunk;
  seen.push({ url: req.url!, headers: req.headers, body });
  respond(req, res);
});
before(async () => {
  await new Promise<void>((done) => server.listen(0, "127.0.0.1", done));
  origin = `http://127.0.0.1:${(server.address() as AddressInfo).port}`;
});
after(() => server.close());

const env11 = (body: string) =>
  `<?xml version="1.0"?><s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/" xmlns:h="urn:example:hr:v1"><s:Body>${body}</s:Body></s:Envelope>`;
const xml =
  (body: string, status = 200, type = "text/xml; charset=utf-8") =>
  (_req: IncomingMessage, res: ServerResponse) => {
    res.writeHead(status, { "content-type": type });
    res.end(body);
  };
const employee = (department: string) =>
  env11(
    `<h:GetEmployeeResponse><h:employee><h:id>E1</h:id><h:name>Ada</h:name><h:department>${department}</h:department><h:salary>1</h:salary></h:employee></h:GetEmployeeResponse>`,
  );

function hr(soap: Record<string, unknown> = {}, extra: Record<string, unknown> = {}) {
  return connectorSchema.parse({
    ...example,
    ...extra,
    soap: { ...example.soap, endpoint: `${origin}/services/hr`, allowInsecureHttp: true, ...soap },
  });
}

// Registers the tools on a server that only records, so a test can call them directly.
async function run(
  connector: Connector,
  tool: string,
  args: Record<string, unknown>,
  { limits = {} as ToolContext["limits"] } = {},
) {
  const vendor = soapVendor(connector);
  const handlers = new Map<string, (a: Record<string, unknown>) => Promise<unknown>>();
  const recorder = {
    registerTool: (name: string, _c: unknown, h: never) => handlers.set(name, h),
  } as unknown as McpServer;
  const audit: { decision: string; reason: string | null; vendorStatus: number | null }[] = [];
  const conn = { config: { host: "x", username: "svc" }, secret: { password: "p&ss" } };
  const ctx = {
    limits,
    config: conn.config,
    call: (r) => vendor.call(conn, r),
    audit: async (e) => void audit.push(e),
  } as ToolContext;
  const toolset = soapTools(connector.soap!);
  toolset.register(recorder, ctx, new Set(toolset.names));
  const result = (await handlers.get(tool)!(args)) as {
    isError?: boolean;
    content: { text: string }[];
  };
  return { result, text: result.content[0].text, audit };
}

test("the example loads and runs on the generic module", () => {
  const connector = hr();
  const app = soapModule.create({ connector } as never);
  assertToolsImplemented([{ connector, ...app }]);
  const soft = {
    ...connector,
    capabilities: connector.capabilities.map((c) => ({ ...c, destructive: false })),
  };
  assert.throws(
    () =>
      assertToolsImplemented([
        { connector: soft, ...soapModule.create({ connector: soft } as never) },
      ]),
    /capability terminate has hr_terminate_employee/,
  );
});

test("soap definitions are validated", () => {
  const [get, search] = example.soap.tools;
  const tools = (...t: unknown[]) => ({ tools: t, verify: undefined });
  assert.throws(() => hr(tools({ ...get, params: {} })), /placeholder/);
  assert.throws(
    () => hr(tools({ ...get, body: `${get.body}{employee_id}` })),
    /placeholder in the tool's body, once/,
  );
  assert.throws(
    () =>
      hr(
        tools({
          ...search,
          body: "<hr:S><hr:q>name {name_prefix}</hr:q><hr:d>{department}</hr:d></hr:S>",
        }),
      ),
    /whole content of an element/,
  );
  assert.throws(() => hr(tools({ ...get, readOnly: false })), /answerLimits are for readOnly/);
  assert.throws(() => hr({ wsSecurity: undefined }), /needs auth or wsSecurity/);
  assert.throws(() => hr({ allowInsecureHttp: false }), /must be https/);
  assert.throws(() => hr({ namespaces: { soap: "urn:x" } }), /Reserved/);
  assert.throws(() => hr({ verify: { tool: "nope", show: "x" } }), /verify.tool/);
  // The WS-Security password must be a secret field.
  assert.throws(
    () =>
      hr(
        {},
        {
          fields: [
            { id: "username", label: "u", secret: false },
            { id: "password", label: "p", secret: false },
          ],
        },
      ),
    /auth names its fields/,
  );
  assert.throws(
    () =>
      hr(
        {},
        {
          http: {
            baseUrl: "https://x.example",
            auth: { type: "bearer", token: "password" },
            tools: [],
          },
        },
      ),
    /at most one of http, soap, odata and sql/,
  );
  // A body that is not XML fails at startup, not on a call.
  const broken = hr(tools({ ...get, body: "<hr:GetEmployee>{employee_id}</hr:Other>" }));
  assert.throws(() => soapModule.create({ connector: broken } as never), /body is not XML/);
});

test("a call is an escaped SOAP 1.1 envelope with the action and a UsernameToken", async () => {
  respond = xml(employee("IT"));
  seen.length = 0;
  const { result } = await run(hr(), "hr_get_employee", {
    employee_id: `E1</hr:employeeId><hr:x>&"'{employee_id}`,
  });
  assert.equal(result.isError, undefined);
  const sent = seen[0];
  assert.equal(sent.url, "/services/hr");
  assert.equal(sent.headers["content-type"], "text/xml; charset=utf-8");
  assert.equal(sent.headers.soapaction, '"urn:example:hr:v1/GetEmployee"');
  assert.match(
    sent.body,
    /<hr:employeeId>E1&lt;\/hr:employeeId&gt;&lt;hr:x&gt;&amp;&quot;&apos;\{employee_id\}<\/hr:employeeId>/,
  );
  const doc = toJson(parseXml(sent.body)) as {
    Header: { Security: { UsernameToken: Record<string, unknown> } };
  };
  const token = doc.Header.Security.UsernameToken;
  assert.equal(token.Username, "svc");
  const password = token.Password as { Type: string; _text: string };
  assert.equal(password._text, "p&ss");
  assert.match(password.Type, /#PasswordText$/);
  assert.equal(sent.headers.authorization, undefined);
});

test("a digest UsernameToken hashes nonce, time and password", async () => {
  respond = xml(employee("IT"));
  seen.length = 0;
  await run(hr({ wsSecurity: { ...example.soap.wsSecurity, digest: true } }), "hr_get_employee", {
    employee_id: "E1",
  });
  const token = (
    toJson(parseXml(seen[0].body)) as never as {
      Header: {
        Security: {
          UsernameToken: { Password: { _text: string }; Nonce: { _text: string }; Created: string };
        };
      };
    }
  ).Header.Security.UsernameToken;
  const expected = createHash("sha1")
    .update(
      Buffer.concat([
        Buffer.from(token.Nonce._text, "base64"),
        Buffer.from(`${token.Created}p&ss`),
      ]),
    )
    .digest("base64");
  assert.equal(token.Password._text, expected);
  assert.doesNotMatch(seen[0].body, /p&amp;ss/);
});

test("SOAP 1.2 sends the action in the content type, and basic auth signs the call", async () => {
  respond = xml(
    `<e:Envelope xmlns:e="http://www.w3.org/2003/05/soap-envelope"><e:Body><GetEmployeeResponse><employee><id>E1</id><department>IT</department></employee></GetEmployeeResponse></e:Body></e:Envelope>`,
    200,
    "application/soap+xml; charset=utf-8",
  );
  seen.length = 0;
  const connector = hr({
    version: "1.2",
    wsSecurity: undefined,
    auth: { type: "basic", username: "username", password: "password" },
  });
  const { text } = await run(connector, "hr_get_employee", { employee_id: "E1" });
  assert.deepEqual(JSON.parse(text), { id: "E1", department: "IT" });
  assert.equal(
    seen[0].headers["content-type"],
    'application/soap+xml; charset=utf-8; action="urn:example:hr:v1/GetEmployee"',
  );
  assert.equal(seen[0].headers.soapaction, undefined);
  assert.match(
    seen[0].body,
    /<soap:Envelope xmlns:soap="http:\/\/www.w3.org\/2003\/05\/soap-envelope"/,
  );
  assert.doesNotMatch(seen[0].body, /Security/);
  assert.equal(
    seen[0].headers.authorization,
    `Basic ${Buffer.from("svc:p&ss").toString("base64")}`,
  );
});

test("answers are trimmed to the result path and fields; [] always gives a list", async () => {
  respond = xml(employee("IT"));
  const one = await run(hr(), "hr_get_employee", { employee_id: "E1" });
  assert.deepEqual(JSON.parse(one.text), { id: "E1", name: "Ada", department: "IT" });
  const list = (inner: string) =>
    xml(
      env11(
        `<h:SearchEmployeesResponse><h:employees>${inner}</h:employees></h:SearchEmployeesResponse>`,
      ),
    );
  respond = list(`<h:employee><h:id>E1</h:id><h:name>Ada</h:name><h:ssn>x</h:ssn></h:employee>`);
  seen.length = 0;
  const single = await run(hr(), "hr_search_employees", { department: "IT" });
  assert.deepEqual(JSON.parse(single.text), [{ id: "E1", name: "Ada" }]);
  // The optional prefix left out drops its element.
  assert.doesNotMatch(seen[0].body, /namePrefix/);
  respond = list(
    `<h:employee><h:id>E1</h:id></h:employee><h:employee><h:id>E2</h:id></h:employee>`,
  );
  const two = await run(hr(), "hr_search_employees", { department: "IT", name_prefix: "A" });
  assert.deepEqual(JSON.parse(two.text), [{ id: "E1" }, { id: "E2" }]);
  assert.match(seen[1].body, /<hr:namePrefix>A<\/hr:namePrefix>/);
  respond = list("");
  assert.deepEqual(
    JSON.parse((await run(hr(), "hr_search_employees", { department: "IT" })).text),
    [],
  );
});

test("a limited param must be one of the agent's values, and sent", async () => {
  respond = xml(env11("<h:UpdateWorkPhoneResponse><h:ok>true</h:ok></h:UpdateWorkPhoneResponse>"));
  seen.length = 0;
  const limits = { departments: ["IT"] };
  const args = { employee_id: "E1", phone: "+1 555 0100" };
  const ok = await run(hr(), "hr_update_work_phone", { ...args, department: "IT" }, { limits });
  assert.deepEqual(JSON.parse(ok.text), { ok: "true" });
  for (const department of ["FIN", undefined]) {
    const { result, audit } = await run(
      hr(),
      "hr_update_work_phone",
      { ...args, department },
      { limits },
    );
    assert.equal(result.isError, true);
    assert.equal(audit[0].decision, "denied");
  }
  assert.equal(seen.length, 1);
});

test("an answer limit refuses an answer outside the agent's values, failing closed", async () => {
  const limits = { departments: ["IT"] };
  respond = xml(employee("IT"));
  assert.equal(
    (await run(hr(), "hr_get_employee", { employee_id: "E1" }, { limits })).result.isError,
    undefined,
  );
  for (const body of [
    employee("FIN"),
    env11(
      "<h:GetEmployeeResponse><h:employee><h:id>E1</h:id></h:employee></h:GetEmployeeResponse>",
    ),
  ]) {
    respond = xml(body);
    const { result, text, audit } = await run(
      hr(),
      "hr_get_employee",
      { employee_id: "E1" },
      { limits },
    );
    assert.equal(result.isError, true);
    assert.doesNotMatch(text, /Ada|FIN/);
    assert.equal(audit[0].decision, "denied");
    assert.equal(audit[0].vendorStatus, 200);
  }
});

test("faults of either version become plain errors", async () => {
  const cases: [string, number, string, RegExp][] = [
    [
      env11(
        "<s:Fault><faultcode>s:Client</faultcode><faultstring>No employee E9</faultstring></s:Fault>",
      ),
      500,
      "text/xml",
      /^The service refused the call: No employee E9 \(s:Client\)$/,
    ],
    [
      `<e:Envelope xmlns:e="http://www.w3.org/2003/05/soap-envelope"><e:Body><e:Fault><e:Code><e:Value>e:Sender</e:Value></e:Code><e:Reason><e:Text xml:lang="en">Bad id</e:Text></e:Reason></e:Fault></e:Body></e:Envelope>`,
      400,
      "application/soap+xml",
      /^The service refused the call: Bad id \(e:Sender\)$/,
    ],
    // Some services answer a fault with 200.
    [
      env11("<s:Fault><faultstring>Locked</faultstring></s:Fault>"),
      200,
      "text/xml",
      /refused the call: Locked/,
    ],
    [
      "<html>Bad gateway</html>",
      502,
      "text/html",
      /^The service answered 502: <html>Bad gateway<\/html>$/,
    ],
  ];
  for (const [body, status, type, why] of cases) {
    respond = xml(body, status, type);
    const { result, text, audit } = await run(hr(), "hr_get_employee", { employee_id: "E9" });
    assert.equal(result.isError, true);
    assert.match(text, why);
    assert.equal(audit[0].decision, "failed");
    assert.equal(audit[0].vendorStatus, status === 200 ? 500 : status);
  }
});

test("answers with a DTD, unknown entities or no envelope are refused", async () => {
  for (const body of [
    `<?xml version="1.0"?><!DOCTYPE x [<!ENTITY e SYSTEM "file:///etc/passwd">]><x>&e;</x>`,
    `<!DOCTYPE l [<!ENTITY a "aaaa"><!ENTITY b "&a;&a;&a;">]>${env11("<x>&b;</x>")}`,
    env11("<h:GetEmployeeResponse>&e;</h:GetEmployeeResponse>"),
    "<employee><id>E1</id></employee>",
  ]) {
    respond = xml(body);
    const { result, text, audit } = await run(hr(), "hr_get_employee", { employee_id: "E1" });
    assert.equal(result.isError, true);
    assert.match(text, /not usable XML|without a SOAP envelope/);
    assert.equal(audit[0].decision, "failed");
  }
  respond = xml(employee("IT"), 200, "application/octet-stream");
  assert.match(
    (await run(hr(), "hr_get_employee", { employee_id: "E1" })).text,
    /application\/octet-stream; only XML, JSON or text is passed on/,
  );
});

test("a value XML cannot carry is refused before any call", async () => {
  seen.length = 0;
  const { result, audit } = await run(hr(), "hr_get_employee", { employee_id: "E1\u0000" });
  assert.equal(result.isError, true);
  assert.equal(audit[0].decision, "denied");
  assert.equal(seen.length, 0);
});

test("the verify check calls the configured operation and shows its answer", async () => {
  respond = xml(
    env11(
      "<h:SearchEmployeesResponse><h:employees><h:employee><h:name>Ada</h:name></h:employee><h:employee><h:name>Grace</h:name></h:employee></h:employees></h:SearchEmployeesResponse>",
    ),
  );
  seen.length = 0;
  const vendor = soapVendor(hr());
  const conn = await vendor.connect({ username: "svc", password: "pw" });
  assert.equal(conn.accountLabel, new URL(origin).host);
  const { values } = await vendor.checks!.operation(conn, { expected: [] });
  assert.deepEqual(values, ["Ada", "Grace"]);
  assert.match(seen[0].body, /<hr:department>IT<\/hr:department>/);
  respond = xml(env11("<s:Fault><faultstring>Bad user</faultstring></s:Fault>"), 500);
  await assert.rejects(
    vendor.checks!.operation(conn, { expected: [] }),
    /refused the call: Bad user/,
  );
});

test("the XML reader", () => {
  const doc = parseXml(
    `<?xml version="1.0"?><!-- hi --><a xmlns:x="urn:x" id="1" x:type="t"><b>1 &lt; 2 &#x41;&#66;</b><b><![CDATA[<raw>]]></b><c x:nil="true"/><__proto__><polluted>yes</polluted></__proto__><d k='v'>text</d></a>`,
  );
  assert.deepEqual(toJson(doc), {
    b: ["1 < 2 AB", "<raw>"],
    c: null,
    ["__proto__"]: { polluted: "yes" },
    d: { k: "v", _text: "text" },
    id: "1",
  });
  assert.equal(({} as Record<string, unknown>).polluted, undefined);
  for (const bad of [
    "<a><b></a>",
    "<a>&nbsp;</a>",
    "<a>& b</a>",
    "<a></a><b></b>",
    "<a>&#0;</a>",
    `<a x="1" x="2"/>`,
    "<a>",
    `${"<a>".repeat(101)}${"</a>".repeat(101)}`,
  ])
    assert.throws(() => parseXml(bad), XmlError, bad);
});
