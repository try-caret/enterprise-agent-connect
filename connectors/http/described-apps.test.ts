import assert from "node:assert/strict";
import { generateKeyPairSync, randomBytes } from "node:crypto";
import { test } from "node:test";
import type { McpServer } from "@modelcontextprotocol/server";
import type { ZodType } from "zod";
import type { ToolContext } from "../../packages/connector-kit/src/common.ts";
import { loadConnectors, type Connector } from "../../packages/connector-kit/src/connector.ts";
import type { Limits, VendorConnection } from "../../packages/connector-kit/src/types.ts";
import { httpTools, httpVendor } from "./index.ts";
import { builtinConnectors } from "../../packages/server/src/builtin.ts";
import { CredentialKey } from "../../packages/server/src/secrets.ts";
import { assertToolsImplemented, createService } from "../../packages/server/src/service.ts";
import { fakeStore, ORG } from "../../packages/server/test/helpers.ts";

// The built-in apps described only in their connector.json, each run through one representative
// tool against a fake app: the request it makes (method, address, auth) and what the agent gets.

const GUID = "11111111-2222-3333-4444-555555555555";
const connectors = loadConnectors();
const described = connectors.filter((c) => c.http);

type Seen = { method: string; url: string; headers: Record<string, string>; body: string };

// Answers a token request with a token and every other call with `answer`.
function fakeApp(answer: unknown, headers: Record<string, string> = {}) {
  const seen: Seen[] = [];
  const fetchImpl = (async (input: string | URL, init?: RequestInit) => {
    const url = String(input);
    seen.push({
      method: init?.method ?? "GET",
      url,
      headers: Object.fromEntries(new Headers(init?.headers)),
      body: init?.body === undefined ? "" : String(init.body),
    });
    if (/\/oauth\/(v1\/)?token$/.test(new URL(url).pathname))
      return Response.json({ access_token: "minted", expires_in: 3600 });
    return Response.json(answer, { headers });
  }) as typeof fetch;
  return { seen, fetchImpl };
}

async function run(
  connector: Connector,
  conn: VendorConnection,
  tool: string,
  args: Record<string, unknown>,
  app: ReturnType<typeof fakeApp>,
  limits: Limits = {},
) {
  const vendor = httpVendor(connector, app.fetchImpl);
  const handlers = new Map<string, (a: Record<string, unknown>) => Promise<unknown>>();
  const recorder = {
    registerTool: (name: string, c: { inputSchema: ZodType }, h: (a: unknown) => unknown) =>
      handlers.set(name, async (a) => h(c.inputSchema.parse(a))),
  } as unknown as McpServer;
  const ctx = {
    limits,
    config: conn.config,
    call: (r) => vendor.call(conn, r),
    audit: async () => {},
  } as ToolContext;
  const toolset = httpTools(connector.http!);
  toolset.register(recorder, ctx, new Set(toolset.names));
  const result = (await handlers.get(tool)!(args)) as {
    isError?: boolean;
    content: { text: string }[];
  };
  return { isError: result.isError ?? false, text: result.content[0].text };
}

const app = (id: string) => described.find((c) => c.id === id)!;

test("each described built-in app loads, is offered and runs on the generic module", () => {
  for (const id of ["gitlab", "argo-cd", "sonarqube", "datadog", "pagerduty", "grafana", "zephyr", "confluence"])
    assert.ok(described.some((c) => c.id === id), id);
  for (const connector of described) {
    assert.ok(builtinConnectors.some((m) => m.id === connector.id), `${connector.id} is offered`);
    assertToolsImplemented([
      { connector, toolset: httpTools(connector.http!), vendor: httpVendor(connector) },
    ]);
  }
});

test("devops apps: one tool each, against a fake app", async () => {
  // GitLab: a namespaced project path is one encoded segment, and X-Next-Page pages.
  let fake = fakeApp([{ id: 7, status: "failed", ref: "main", extra: "x" }], { "x-next-page": "2" });
  let out = await run(
    app("gitlab"),
    { config: { host: "gitlab.example.com" }, secret: { token: "glpat" } },
    "gitlab_list_pipelines",
    { project: "platform/payments-api", status: "failed" },
    fake,
  );
  assert.equal(fake.seen[0].method, "GET");
  assert.equal(
    fake.seen[0].url,
    "https://gitlab.example.com/api/v4/projects/platform%2Fpayments-api/pipelines?status=failed&per_page=20",
  );
  assert.equal(fake.seen[0].headers["private-token"], "glpat");
  const page = JSON.parse(out.text);
  assert.equal(page.next, 2);
  assert.equal(page.items[0].status, "failed");
  assert.equal(page.items[0].extra, undefined);

  // Argo CD: a key kept to a project sends it, and Argo CD checks the application is in it.
  fake = fakeApp({ nodes: [{ kind: "Pod", name: "web-1", health: { status: "Degraded" } }] });
  out = await run(
    app("argo-cd"),
    { config: { host: "argocd.example.com" }, secret: { token: "argo" } },
    "argocd_get_resource_tree",
    { name: "payments", project: "team-a" },
    fake,
    { projects: ["team-a"] },
  );
  assert.equal(
    fake.seen[0].url,
    "https://argocd.example.com/api/v1/applications/payments/resource-tree?project=team-a",
  );
  assert.equal(fake.seen[0].headers.authorization, "Bearer argo");
  assert.equal(JSON.parse(out.text)[0]["health.status"], "Degraded");
  const refused = await run(
    app("argo-cd"),
    { config: { host: "argocd.example.com" }, secret: { token: "argo" } },
    "argocd_get_resource_tree",
    { name: "payments" },
    fakeApp({}),
    { projects: ["team-a"] },
  );
  assert.equal(refused.isError, true);

  // SonarQube: the agent's names become SonarQube's.
  fake = fakeApp({ total: 0, p: 1, issues: [] });
  await run(
    app("sonarqube"),
    { config: { host: "sonarqube.example.com" }, secret: { token: "squ" } },
    "sonarqube_search_issues",
    { project: "acme:payments", pull_request: "42" },
    fake,
  );
  assert.equal(
    fake.seen[0].url,
    "https://sonarqube.example.com/api/issues/search?componentKeys=acme%3Apayments&pullRequest=42&resolved=false&p=1&ps=50",
  );
  assert.equal(fake.seen[0].headers.authorization, "Bearer squ");

  // Datadog: both keys, and only Datadog's own hosts.
  fake = fakeApp({ data: [], meta: { page: { after: "c2" } } });
  const datadog = { config: { host: "api.datadoghq.eu" }, secret: { api_key: "k", application_key: "a" } };
  out = await run(
    app("datadog"),
    datadog,
    "datadog_search_logs",
    { query: "service:checkout status:error", from: "2026-10-01T08:00:00Z", to: "2026-10-01T09:00:00Z" },
    fake,
  );
  assert.equal(
    fake.seen[0].url,
    "https://api.datadoghq.eu/api/v2/logs/events?filter%5Bquery%5D=service%3Acheckout+status%3Aerror&filter%5Bfrom%5D=2026-10-01T08%3A00%3A00Z&filter%5Bto%5D=2026-10-01T09%3A00%3A00Z&sort=-timestamp&page%5Blimit%5D=50",
  );
  assert.equal(fake.seen[0].headers["dd-api-key"], "k");
  assert.equal(fake.seen[0].headers["dd-application-key"], "a");
  assert.equal(JSON.parse(out.text).next, "c2");
  fake = fakeApp({});
  out = await run(
    app("datadog"),
    { ...datadog, config: { host: "datadog.attacker.example" } },
    "datadog_query_metrics",
    { query: "avg:system.cpu.user{*}", from: 1, to: 2 },
    fake,
  );
  assert.equal(out.isError, true);
  assert.equal(fake.seen.length, 0);

  // PagerDuty: Token auth, the v2 Accept, array params and a service limit.
  fake = fakeApp({ incidents: [], more: false });
  await run(
    app("pagerduty"),
    { config: { host: "api.pagerduty.com" }, secret: { api_key: "pd" } },
    "pagerduty_list_incidents",
    { statuses: ["triggered", "acknowledged"], service_ids: ["PABC123"] },
    fake,
    { services: ["PABC123"] },
  );
  assert.equal(
    fake.seen[0].url,
    "https://api.pagerduty.com/incidents?statuses%5B%5D=triggered&statuses%5B%5D=acknowledged&service_ids%5B%5D=PABC123&sort_by=created_at%3Adesc&limit=25&offset=0",
  );
  assert.equal(fake.seen[0].headers.authorization, "Token token=pd");
  assert.equal(fake.seen[0].headers.accept, "application/vnd.pagerduty+json;version=2");

  // Grafana: a dashboard outside the key's folders is refused after the read.
  fake = fakeApp({ dashboard: { uid: "d1", title: "API" }, meta: { folderUid: "other" } });
  out = await run(
    app("grafana"),
    { config: { host: "grafana.example.com" }, secret: { token: "glsa" } },
    "grafana_get_dashboard",
    { uid: "d1" },
    fake,
    { folders: ["ops"] },
  );
  assert.equal(fake.seen[0].url, "https://grafana.example.com/api/dashboards/uid/d1");
  assert.equal(fake.seen[0].headers.authorization, "Bearer glsa");
  assert.equal(out.isError, true);

  // Zephyr: a result recorded with Zephyr's field names, in the EU region.
  fake = fakeApp({ id: 1, self: "x" });
  await run(
    app("zephyr"),
    { config: { host: "eu.api.zephyrscale.smartbear.com" }, secret: { token: "zs" } },
    "zephyr_create_test_execution",
    { project_key: "QA", test_case_key: "QA-T40", test_cycle_key: "QA-R12", status_name: "Fail", comment: "Login times out" },
    fake,
  );
  assert.equal(fake.seen[0].method, "POST");
  assert.equal(fake.seen[0].url, "https://eu.api.zephyrscale.smartbear.com/v2/testexecutions");
  assert.equal(fake.seen[0].headers.authorization, "Bearer zs");
  assert.deepEqual(JSON.parse(fake.seen[0].body), {
    projectKey: "QA",
    testCaseKey: "QA-T40",
    testCycleKey: "QA-R12",
    statusName: "Fail",
    comment: "Login times out",
  });

  // Confluence: a service account token, then the site's gateway address by its cloud ID.
  fake = fakeApp(
    { results: [{ id: "9", title: "Release 4.2 checklist", spaceId: "123", body: {} }] },
    { link: '</wiki/api/v2/pages?space-id=123&cursor=abc>; rel="next"' },
  );
  out = await run(
    app("confluence"),
    { config: { cloud: GUID, client_id: "cid" }, secret: { client_secret: "cs" } },
    "confluence_list_pages",
    { space_ids: ["123", "456"] },
    fake,
  );
  assert.equal(fake.seen[0].url, "https://auth.atlassian.com/oauth/token");
  assert.match(fake.seen[0].body, /grant_type=client_credentials/);
  assert.equal(fake.seen[1].method, "GET");
  assert.equal(
    fake.seen[1].url,
    `https://api.atlassian.com/ex/confluence/${GUID}/wiki/api/v2/pages?space-id=123%2C456&sort=-modified-date&limit=50`,
  );
  assert.equal(fake.seen[1].headers.authorization, "Bearer minted");
  assert.equal(JSON.parse(out.text).next, "abc");
});

test("HR and procurement apps: one tool each, against a fake app", async () => {
  // Zip: the key header, a department limit on the search, and a request outside it refused.
  const zip = { config: {}, secret: { api_key: "zk" } };
  let fake = fakeApp({
    list: [{ id: "r1", request_number: 42, status: 1, department: { id: "d1", name: "IT" }, owner: {} }],
    next_page_token: "p2",
  });
  let out = await run(app("zip"), zip, "zip_search_requests", { status: 1, department_id: ["d1"] }, fake, {
    departments: ["d1"],
  });
  assert.equal(fake.seen[0].method, "GET");
  assert.equal(
    fake.seen[0].url,
    "https://api.ziphq.com/requests?status=1&department_id=d1&sort_by=create_time&page_size=20",
  );
  assert.equal(fake.seen[0].headers["zip-api-key"], "zk");
  let page = JSON.parse(out.text);
  assert.equal(page.next, "p2");
  assert.equal(page.items[0]["department.name"], "IT");
  assert.equal(page.items[0].owner, undefined);
  fake = fakeApp({});
  out = await run(app("zip"), zip, "zip_search_requests", {}, fake, { departments: ["d1"] });
  assert.equal(out.isError, true);
  assert.equal(fake.seen.length, 0);
  fake = fakeApp({ id: "r2", department: { id: "d9" } });
  out = await run(app("zip"), zip, "zip_get_request", { request_id: "r2" }, fake, { departments: ["d1"] });
  assert.equal(fake.seen[0].url, "https://api.ziphq.com/requests/r2");
  assert.equal(out.isError, true);

  // UKG Pro: Basic sign-in plus the customer key, a company limit, and only UKG's own hosts.
  const ukg = {
    config: { host: "service5.ultipro.com", username: "svc" },
    secret: { password: "pw", customer_api_key: "ck" },
  };
  fake = fakeApp([{ employeeID: "E1", companyID: "C1", annualSalary: 90000 }]);
  out = await run(app("ukg-pro"), ukg, "ukg_list_compensation", { company_id: "C1", date_in_job: ">01-01-2026" }, fake, {
    companies: ["C1"],
  });
  assert.equal(
    fake.seen[0].url,
    "https://service5.ultipro.com/personnel/v1/compensation-details?companyId=C1&dateInJob=%3E01-01-2026&page=1&per_Page=100",
  );
  assert.equal(fake.seen[0].headers.authorization, `Basic ${btoa("svc:pw")}`);
  assert.equal(fake.seen[0].headers["us-customer-api-key"], "ck");
  assert.equal(JSON.parse(out.text)[0].annualSalary, 90000);
  fake = fakeApp([{ employeeId: "E1", companyId: "C1" }, { employeeId: "E2", companyId: "C2" }]);
  out = await run(app("ukg-pro"), ukg, "ukg_list_time_off_plans", {}, fake, { companies: ["C1"] });
  assert.equal(fake.seen[0].url, "https://service5.ultipro.com/personnel/v1/pto-plans?page=1&per_page=100");
  assert.equal(out.isError, true);
  fake = fakeApp([{ employeeId: "E1", nationalId: "123-45-6789", firstName: "Ann" }]);
  out = await run(app("ukg-pro"), ukg, "ukg_list_people", { last_name: "Lee" }, fake);
  assert.equal(JSON.parse(out.text)[0].nationalId, undefined);
  assert.equal(JSON.parse(out.text)[0].firstName, "Ann");
  fake = fakeApp([]);
  out = await run(app("ukg-pro"), { ...ukg, config: { ...ukg.config, host: "ukg.attacker.example" } }, "ukg_list_jobs", {}, fake);
  assert.equal(out.isError, true);
  assert.equal(fake.seen.length, 0);

  // iCIMS: the customer id a path segment, the search JSON encoded, ids back; EU host allowed.
  const icims = { config: { host: "api-eu.icims.com", customer: "1234", username: "u" }, secret: { password: "p" } };
  fake = fakeApp({ searchResults: [{ id: "7", self: "x" }] });
  const search = '{"filters":[{"name":"job.folder","value":["Approved"]}]}';
  out = await run(app("icims"), icims, "icims_search_jobs", { search, staleness: 0 }, fake);
  assert.equal(
    fake.seen[0].url,
    `https://api-eu.icims.com/customers/1234/search/jobs?searchJson=${encodeURIComponent(search).replace(/%20/g, "+")}&staleness=0`,
  );
  assert.equal(fake.seen[0].headers.authorization, `Basic ${btoa("u:p")}`);
  assert.deepEqual(JSON.parse(out.text), [{ id: "7" }]);
  fake = fakeApp({ id: 5 });
  await run(app("icims"), icims, "icims_find_person_by_email", { email: "a@example.com", fields: "firstname,folder" }, fake);
  assert.equal(
    fake.seen[0].url,
    "https://api-eu.icims.com/customers/1234/people/by-email?email=a%40example.com&fields=firstname%2Cfolder",
  );
  fake = fakeApp({});
  out = await run(app("icims"), { ...icims, config: { ...icims.config, host: "api.attacker.example" } }, "icims_get_job", { job_id: "7" }, fake);
  assert.equal(out.isError, true);
  assert.equal(fake.seen.length, 0);
});

test("connecting proves each sign-in with one cheap read: Zip, UKG Pro, iCIMS", async () => {
  const key = new CredentialKey(randomBytes(32).toString("base64"));
  const fake = fakeApp({ total: 12 });
  const ukgFake = fakeApp([{ companyID: "C1", companyName: "Acme" }]);
  const icimsFake = fakeApp({ lists: ["lists.job.folders"] });
  const fakes: Record<string, ReturnType<typeof fakeApp>> = { zip: fake, "ukg-pro": ukgFake, icims: icimsFake };
  const service = createService({
    apps: Object.entries(fakes).map(([id, f]) => ({
      connector: app(id),
      vendor: httpVendor(app(id), f.fetchImpl),
      toolset: httpTools(app(id).http!),
    })),
    store: fakeStore(key).store,
    key,
  });
  const connect = (connectorId: string, capabilities: string[], fields: Record<string, string>) =>
    service.connect({ organizationId: ORG, connectorId, capabilities, fields, connectedBy: "admin@x" });

  let r = await connect("zip", ["requests"], { api_key: "zk" });
  assert.equal(fake.seen[0].url, "https://api.ziphq.com/requests?page_size=1");
  assert.deepEqual(r.checks[0].values, ["12"]);
  r = await connect("ukg-pro", ["read"], { host: "service5.ultipro.com", username: "svc", password: "pw", customer_api_key: "ck" });
  assert.equal(ukgFake.seen[0].url, "https://service5.ultipro.com/configuration/v1/company-details?per_Page=1");
  assert.deepEqual(r.checks[0].values, ["Acme"]);
  r = await connect("icims", ["jobs"], { host: "api.icims.com", customer: "1234", username: "u", password: "p" });
  assert.equal(icimsFake.seen[0].url, "https://api.icims.com/customers/1234/v2/lists");
  assert.deepEqual(r.checks[0].values, ["lists.job.folders"]);
});

test("ZoomInfo, DocuSign and Axonius: one tool each, against a fake app", async () => {
  // ZoomInfo: Basic client auth at the GTM token service, then a JSON:API search body.
  let fake = fakeApp({
    data: [{ id: "344589814", type: "Company", attributes: { name: "ZoomInfo", country: "United States", logo: "x" } }],
    meta: { totalResults: 1, page: { total: 1, number: 1 } },
  });
  let out = await run(
    app("zoominfo"),
    { config: { host: "api.zoominfo.com", client_id: "cid" }, secret: { client_secret: "cs" } },
    "zoominfo_search_companies",
    { company_name: "ZoomInfo", country: "United States" },
    fake,
  );
  assert.equal(fake.seen[0].url, "https://api.zoominfo.com/gtm/oauth/v1/token");
  assert.equal(fake.seen[0].headers.authorization, `Basic ${Buffer.from("cid:cs").toString("base64")}`);
  assert.equal(fake.seen[1].method, "POST");
  assert.equal(
    fake.seen[1].url,
    "https://api.zoominfo.com/gtm/data/v1/companies/search?page%5Bnumber%5D=1&page%5Bsize%5D=25",
  );
  assert.equal(fake.seen[1].headers["content-type"], "application/vnd.api+json");
  assert.equal(fake.seen[1].headers.accept, "application/vnd.api+json");
  assert.equal(fake.seen[1].headers.authorization, "Bearer minted");
  assert.deepEqual(JSON.parse(fake.seen[1].body), {
    data: {
      type: "CompanySearch",
      attributes: { companyName: "ZoomInfo", country: "United States", excludeDefunctCompanies: true },
    },
  });
  const company = JSON.parse(out.text)[0];
  assert.equal(company["attributes.name"], "ZoomInfo");
  assert.equal(company["attributes.logo"], undefined);
  // Enrich matches one company by id, in the list ZoomInfo wants.
  fake = fakeApp({ data: [] });
  await run(
    app("zoominfo"),
    { config: { host: "api.zoominfo.com", client_id: "cid" }, secret: { client_secret: "cs" } },
    "zoominfo_enrich_company",
    { company_id: 344589814, output_fields: ["id", "name", "parentName"] },
    fake,
  );
  assert.deepEqual(JSON.parse(fake.seen[1].body), {
    data: {
      type: "CompanyEnrich",
      attributes: { matchCompanyInput: [{ companyId: 344589814 }], outputFields: ["id", "name", "parentName"] },
    },
  });

  // DocuSign: a JWT Grant at the chosen sign-in host, then the account's own base address.
  const key = generateKeyPairSync("rsa", { modulusLength: 2048 }).privateKey.export({
    type: "pkcs8",
    format: "pem",
  }) as string;
  const docusign = {
    config: {
      auth_host: "account-d.docusign.com",
      host: "demo.docusign.net",
      account: GUID,
      integration_key: "ik",
      user_id: GUID,
    },
    secret: { private_key: key },
  };
  fake = fakeApp({
    resultSetSize: "1",
    envelopes: [{ envelopeId: "e1", status: "sent", emailSubject: "NDA", sender: { userName: "Ann", email: "a@x" }, documentsUri: "/d" }],
  });
  out = await run(
    app("docusign"),
    docusign,
    "docusign_list_envelopes",
    { from_date: "2026-10-01", status: "sent,delivered" },
    fake,
  );
  assert.equal(fake.seen[0].url, "https://account-d.docusign.com/oauth/token");
  const form = new URLSearchParams(fake.seen[0].body);
  assert.equal(form.get("grant_type"), "urn:ietf:params:oauth:grant-type:jwt-bearer");
  const claims = JSON.parse(Buffer.from(form.get("assertion")!.split(".")[1], "base64url").toString());
  assert.equal(claims.aud, "account-d.docusign.com");
  assert.equal(claims.iss, "ik");
  assert.equal(claims.sub, GUID);
  assert.equal(claims.scope, "signature impersonation");
  assert.equal(
    fake.seen[1].url,
    `https://demo.docusign.net/restapi/v2.1/accounts/${GUID}/envelopes?from_date=2026-10-01&status=sent%2Cdelivered&from_to_status=changed&order=desc&count=25&start_position=0`,
  );
  assert.equal(fake.seen[1].headers.authorization, "Bearer minted");
  const envelope = JSON.parse(out.text)[0];
  assert.equal(envelope["sender.userName"], "Ann");
  assert.equal(envelope.documentsUri, undefined);
  // Resending is a PUT with an empty body, nothing else.
  fake = fakeApp({ envelopeId: "e1" });
  await run(app("docusign"), docusign, "docusign_resend_envelope", { envelope_id: "e1" }, fake);
  assert.equal(fake.seen[1].method, "PUT");
  assert.equal(
    fake.seen[1].url,
    `https://demo.docusign.net/restapi/v2.1/accounts/${GUID}/envelopes/e1?resend_envelope=true`,
  );
  assert.equal(fake.seen[1].body, "{}");
  // The account address is held to DocuSign's hosts.
  fake = fakeApp({});
  out = await run(
    app("docusign"),
    { ...docusign, config: { ...docusign.config, host: "docusign.attacker.example" } },
    "docusign_get_envelope",
    { envelope_id: "e1" },
    fake,
  );
  assert.equal(out.isError, true);
  assert.equal(fake.seen.length, 0);

  // Axonius: both keys as headers, an AQL search with a nested page, and its cursor.
  fake = fakeApp({ assets: [{ internal_axon_id: "a1" }], meta: { page: { totalResources: 9 }, next_page: "c2" } });
  out = await run(
    app("axonius"),
    { config: { host: "acme.on.axonius.com" }, secret: { api_key: "k", api_secret: "s" } },
    "axonius_search_devices",
    { query: '("specific_data.data.hostname" == regex("IDC-", "i"))', fields: ["specific_data.data.hostname"], cursor: "c1" },
    fake,
  );
  assert.equal(fake.seen[0].method, "POST");
  assert.equal(fake.seen[0].url, "https://acme.on.axonius.com/api/v2/assets/devices");
  assert.equal(fake.seen[0].headers["api-key"], "k");
  assert.equal(fake.seen[0].headers["api-secret"], "s");
  assert.deepEqual(JSON.parse(fake.seen[0].body), {
    query: '("specific_data.data.hostname" == regex("IDC-", "i"))',
    fields: ["specific_data.data.hostname"],
    page: { limit: 50, cursor: "c1" },
  });
  assert.deepEqual(JSON.parse(out.text), { items: [{ internal_axon_id: "a1" }], next: "c2" });
});
