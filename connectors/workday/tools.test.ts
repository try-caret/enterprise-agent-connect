import assert from "node:assert/strict";
import { test } from "node:test";
import { z } from "zod";
import { workdayTools } from "./tools.ts";
import { parseEndpoint, workdayVendor } from "./vendor.ts";
import { loadConnectors } from "../../packages/connector-kit/src/connector.ts";
import type { ToolContext } from "../../packages/connector-kit/src/common.ts";
import type { VendorConnection, VendorRequest } from "../../packages/connector-kit/src/types.ts";
import { assertToolsImplemented } from "../../packages/server/src/service.ts";

// The Workday tools against a fake tenant `acme`, reached through the real vendor client.
// Supervisory organizations: TOP above IT and SALES; OTHER apart. Ada manages IT, Carol manages
// SALES and TOP. Dan and Erin are in OTHER; Erin also belongs to COMPANY.

const HOST = "wd5-services1.myworkday.com";
const id = (c: string) => c.repeat(32);
const [TOP, IT, SALES, OTHER, COMPANY] = ["1", "2", "3", "4", "5"].map(id);
const [ADA, BOB, CAROL, DAN, ERIN] = ["a", "b", "c", "d", "e"].map(id);
const VACATION = id("f");
const REPORT_LINK = `https://${HOST}/ccx/service/customreport2/acme/jsmith/Open_Reqs?Hiring_Org!WID=${TOP}&As_Of=2026-10-01-07:00&format=json`;

const ORGS: Record<string, { name: string; manager: string; superior?: string; subs: string[] }> = {
  [TOP]: { name: "Company", manager: CAROL, subs: [IT, SALES] },
  [IT]: { name: "IT", manager: ADA, superior: TOP, subs: [] },
  [SALES]: { name: "Sales", manager: CAROL, superior: TOP, subs: [] },
  [OTHER]: { name: "Other", manager: DAN, subs: [] },
};
const PEOPLE: Record<string, { name: string; org: string; company?: string }> = {
  [ADA]: { name: "Ada Lovelace", org: IT },
  [BOB]: { name: "Bob Ross", org: IT },
  [CAROL]: { name: "Carol Danvers", org: SALES },
  [DAN]: { name: "Dan Brown", org: OTHER },
  [ERIN]: { name: "Erin Hunt", org: OTHER, company: COMPANY },
};
const r = (wid: string, name: string) => ({ id: wid, descriptor: name, href: `https://x/${wid}` });

// What staffing answers, with fields the tools must never pass on.
const staffingWorker = (wid: string) => ({
  ...r(wid, PEOPLE[wid].name),
  workerId: `E${wid.slice(0, 3)}`,
  person: { email: `${wid.slice(0, 3)}@acme.example`, phone: "+1 555 0100" },
  primaryJob: {
    businessTitle: "Engineer",
    jobProfile: r(id("9"), "Software Engineer"),
    supervisoryOrganization: r(PEOPLE[wid].org, ORGS[PEOPLE[wid].org].name),
    location: r(id("8"), "Boston"),
  },
  workerType: r(id("7"), "Employee"),
  compensation: { annual: 250000 },
  nationalId: "123-45-6789",
});
const summary = (wid: string) => ({
  ...r(wid, PEOPLE[wid].name),
  businessTitle: "Engineer",
  isManager: wid === ADA,
  primaryWorkEmail: `${wid.slice(0, 3)}@acme.example`,
  primarySupervisoryOrganization: r(PEOPLE[wid].org, ORGS[PEOPLE[wid].org].name),
});
const list = (data: unknown[], total = data.length) => ({ body: { total, data } });

type Reply = { status?: number; body?: unknown };
type Sent = { url: URL; method: string; auth: string | null; body: unknown };

function fakeWorkday(overrides: Record<string, Reply | ((url: URL) => Reply)> = {}) {
  const sent: Sent[] = [];
  let mints = 0;
  const fetchImpl = (async (input: string, init: RequestInit) => {
    const url = new URL(input);
    const headers = new Headers(init.headers);
    const body =
      typeof init.body === "string" && headers.get("content-type") === "application/json"
        ? JSON.parse(init.body)
        : init.body instanceof URLSearchParams
          ? Object.fromEntries(init.body)
          : init.body;
    sent.push({ url, method: init.method ?? "GET", auth: headers.get("authorization"), body });
    const path = url.pathname.replace(/^\/ccx\/(api|service|oauth2)\//, "");
    const key = Object.keys(overrides).find((k) => new RegExp(`^${k}$`).test(path));
    if (key) {
      const o = overrides[key];
      const reply = typeof o === "function" ? o(url) : o;
      return new Response(
        typeof reply.body === "string" ? reply.body : JSON.stringify(reply.body ?? {}),
        { status: reply.status ?? 200 },
      );
    }
    if (path === "acme/token") mints++;
    const reply = answer(path, url, mints);
    return new Response(JSON.stringify(reply.body ?? {}), { status: reply.status ?? 200 });
  }) as unknown as typeof fetch;
  return { sent, fetchImpl, mints: () => mints };
}

function answer(path: string, url: URL, mints: number): Reply {
  const q = url.searchParams;
  let m: RegExpExecArray | null;
  if (path === "acme/token") return { body: { access_token: `tok${mints}`, token_type: "Bearer" } };
  if (path === "staffing/v7/acme/workers") {
    const words = (q.get("search") ?? "").toLowerCase().split(" ");
    const hits = Object.keys(PEOPLE).filter((w) =>
      words.some((x) => PEOPLE[w].name.toLowerCase().includes(x)),
    );
    const offset = Number(q.get("offset") ?? 0);
    const limit = Number(q.get("limit") ?? 20);
    return list(hits.slice(offset, offset + limit).map(staffingWorker), hits.length);
  }
  if ((m = /^staffing\/v7\/acme\/workers\/(\w+)$/.exec(path)))
    return PEOPLE[m[1]]
      ? { body: staffingWorker(m[1]) }
      : { status: 404, body: { error: "Not found" } };
  if ((m = /^staffing\/v7\/acme\/supervisoryOrganizations\/(\w+)(\/orgChart)?$/.exec(path))) {
    const o = ORGS[m[1]];
    if (!o) return { status: 404, body: { error: "Invalid ID" } };
    if (!m[2])
      return {
        body: {
          ...r(m[1], o.name),
          name: o.name,
          code: `SUP-${o.name}`,
          manager: r(o.manager, PEOPLE[o.manager].name),
        },
      };
    return list([
      {
        ...r(m[1], o.name),
        ...(o.superior
          ? { superior: { ...r(o.superior, ORGS[o.superior].name), name: ORGS[o.superior].name } }
          : {}),
        subordinates: o.subs.map((s) => ({
          ...r(s, ORGS[s].name),
          name: ORGS[s].name,
          code: `SUP-${ORGS[s].name}`,
          inactive: false,
          managers: [r(ORGS[s].manager, PEOPLE[ORGS[s].manager].name)],
        })),
      },
    ]);
  }
  if ((m = /^common\/v1\/acme\/workers\/(\w+)\/organizations$/.exec(path))) {
    const p = PEOPLE[m[1]];
    return list([
      { ...r(p.org, ORGS[p.org].name), organizationType: r(id("6"), "Supervisory") },
      ...(p.company
        ? [{ ...r(p.company, "Acme Inc"), organizationType: r(id("6"), "Company") }]
        : []),
    ]);
  }
  if ((m = /^common\/v1\/acme\/workers\/(\w+)\/directReports$/.exec(path)))
    return list(
      Object.keys(PEOPLE)
        .filter((w) => w !== m![1] && ORGS[PEOPLE[w].org].manager === m![1])
        .map(summary),
    );
  if ((m = /^common\/v1\/acme\/supervisoryOrganizations\/(\w+)\/workers$/.exec(path)))
    return list(
      Object.keys(PEOPLE)
        .filter((w) => PEOPLE[w].org === m![1])
        .map(summary),
    );
  if (path === "absenceManagement/v5/acme/balances")
    return list([
      {
        absencePlan: r(id("0"), "Vacation"),
        category: r(id("0"), "Time Off"),
        quantity: "12.5",
        unit: r(id("0"), "Days"),
        effectiveDate: "2026-10-27",
        worker: r(q.get("worker")!, "x"),
        position: r(id("0"), "P1"),
      },
    ]);
  if (/^absenceManagement\/v5\/acme\/workers\/\w+\/timeOffDetails$/.test(path))
    return list([
      {
        date: "2026-10-01",
        timeOffType: r(VACATION, "Vacation"),
        quantity: "1",
        unit: r(id("0"), "Days"),
        status: r(id("0"), "Approved"),
        comment: "Trip",
        position: r(id("0"), "P1"),
      },
    ]);
  if (/^absenceManagement\/v5\/acme\/workers\/\w+\/eligibleAbsenceTypes$/.test(path))
    return list([
      {
        id: VACATION,
        descriptor: "Vacation",
        absenceTypeGroup: r(id("0"), "Paid"),
        unitOfTime: r(id("0"), "Hours"),
        dailyDefaultQuantity: "8",
        reasonRequired: false,
        startAndEndTimeRequired: false,
      },
    ]);
  if (/^absenceManagement\/v5\/acme\/workers\/\w+\/requestTimeOff$/.test(path))
    return {
      status: 201,
      body: {
        businessProcessParameters: {
          overallStatus: "Submitted",
          overallBusinessProcess: r(id("0"), "Request Time Off: Bob Ross"),
        },
      },
    };
  if (path === "customreport2/acme/jsmith/Open_Reqs")
    return {
      body: {
        Report_Entry: [
          { Requisition: "R-1 Engineer", Org_ID: IT, Openings: 2, Detail: { a: 1 } },
          { Requisition: "R-2 Seller", Org_ID: SALES, Openings: 1 },
          { Requisition: "R-3 Other", Org_ID: OTHER, Openings: 1 },
          { Requisition: "R-4 Company", Org_ID: [COMPANY], Openings: 1 },
          { Requisition: "R-5 No org", Openings: 1 },
        ],
      },
    };
  return { status: 404, body: { error: "No such resource" } };
}

const connection: VendorConnection = {
  config: {
    host: HOST,
    tenant: "acme",
    account: `${HOST}/acme`,
    reports: JSON.stringify([
      {
        name: "Open_Reqs",
        owner: "jsmith",
        report: "Open_Reqs",
        prompts: [
          { name: "Hiring_Org!WID", type: "text" },
          { name: "As_Of", type: "date" },
        ],
      },
      { name: "Headcount", owner: "jsmith", report: "Headcount", prompts: [] },
    ]),
    reportOrgColumn: "Org_ID",
  },
  secret: { client_id: "cid", client_secret: "csecret", refresh_token: "rtoken" },
};

function setup(
  limits: Record<string, unknown> = {},
  overrides: Parameters<typeof fakeWorkday>[0] = {},
  config: Record<string, string> = connection.config,
) {
  const wd = fakeWorkday(overrides);
  const vendor = workdayVendor(wd.fetchImpl);
  const conn = { ...connection, config };
  const audit: { tool: string; target: string | null; decision: string; reason: string | null }[] =
    [];
  const ctx = {
    limits,
    config,
    call: (req: VendorRequest) => vendor.call(conn, req),
    forbidden: { workday_get_worker: "Give View on Worker Data: Public Worker Reports." },
    async audit(e) {
      audit.push(e);
    },
  } as ToolContext;
  const tools = new Map<string, (args: unknown) => Promise<any>>();
  // Arguments are parsed with each tool's schema, defaults filled, as the MCP server does.
  const server = {
    registerTool: (name: string, def: { inputSchema: z.ZodType }, fn: (a: unknown) => any) =>
      tools.set(name, (args) => fn(def.inputSchema.parse(args))),
  };
  workdayTools.register(server as never, ctx, new Set(workdayTools.names));
  const run = async (name: string, args: object) => {
    const res = await tools.get(name)!(args);
    const text = res.content[0].text as string;
    return { isError: Boolean(res.isError), text, body: res.isError ? null : JSON.parse(text) };
  };
  const calls = () =>
    wd.sent.filter((s) => !s.url.pathname.endsWith("/token")).map((s) => s.url.pathname);
  return { run, sent: wd.sent, calls, audit, vendor, mints: wd.mints };
}

test("the connector's capabilities, tools and checks line up; nothing is destructive", () => {
  const connector = loadConnectors().find((c) => c.id === "workday")!;
  assertToolsImplemented([{ connector, toolset: workdayTools, vendor: workdayVendor() }]);
  assert.deepEqual(
    connector.capabilities.filter((c) => c.destructive),
    [],
  );
  assert.equal(connector.capabilities.find((c) => c.required)?.id, "workers");
});

test("connect reads the endpoint, signs in with the refresh token and keeps the reports", async () => {
  const { vendor, sent } = setup();
  const conn = await vendor.connect({
    endpoint: ` https://${HOST.toUpperCase()}/ccx/api/v1/acme `,
    client_id: "cid",
    client_secret: "csecret",
    refresh_token: "rtoken",
    reports: `${REPORT_LINK}\nhttps://${HOST}/ccx/service/customreport2/acme/jsmith/Headcount?format=json`,
    report_org_column: "Org_ID",
  });
  assert.equal(conn.accountLabel, `acme (${HOST})`);
  assert.deepEqual(conn.secret, {
    client_id: "cid",
    client_secret: "csecret",
    refresh_token: "rtoken",
  });
  assert.equal(conn.config.account, `${HOST}/acme`);
  assert.equal(conn.config.reportOrgColumn, "Org_ID");
  assert.deepEqual(JSON.parse(conn.config.reports), JSON.parse(connection.config.reports));
  assert.equal(sent[0].url.href, `https://${HOST}/ccx/oauth2/acme/token`);
  assert.equal(sent[0].auth, `Basic ${Buffer.from("cid:csecret").toString("base64")}`);
  assert.deepEqual(sent[0].body, { grant_type: "refresh_token", refresh_token: "rtoken" });

  const fields = { client_id: "c", client_secret: "s", refresh_token: "t" };
  await assert.rejects(
    vendor.connect({ ...fields, endpoint: "https://evil.example.com/ccx/api/v1/acme" }),
    /not a Workday address/,
  );
  await assert.rejects(
    vendor.connect({ ...fields, endpoint: `https://${HOST}/somewhere` }),
    /should end with/,
  );
  await assert.rejects(
    vendor.connect({
      ...fields,
      endpoint: `https://${HOST}/ccx/api/v1/acme`,
      reports: REPORT_LINK.replace("/acme/", "/other/"),
    }),
    /belongs to tenant other/,
  );
  await assert.rejects(
    vendor.connect({
      ...fields,
      endpoint: `https://${HOST}/ccx/api/v1/acme`,
      reports: REPORT_LINK.replace(HOST, "wd1.myworkday.com"),
    }),
    /not a custom report link/,
  );
  const refused = setup({}, { "acme/token": { status: 401, body: { error: "invalid_grant" } } });
  await assert.rejects(
    refused.vendor.connect({ ...fields, endpoint: `https://${HOST}/ccx/api/v1/acme` }),
    /refused the sign-in \(401: invalid_grant\)/,
  );
  assert.deepEqual(parseEndpoint(`https://${HOST}/ccx/api/staffing/v7/acme`), {
    host: HOST,
    tenant: "acme",
  });
});

test("a token is reused, and minted again once after a 401", async () => {
  let first = true;
  const { run, mints } = setup(
    {},
    {
      "staffing/v7/acme/workers/a+": () => {
        if (!first) return { body: staffingWorker(ADA) };
        first = false;
        return { status: 401, body: { error: "expired" } };
      },
    },
  );
  assert.equal((await run("workday_get_worker", { worker_id: ADA })).isError, false);
  assert.equal((await run("workday_get_worker", { worker_id: ADA })).isError, false);
  assert.equal(mints(), 2);
});

test("workday_search_workers trims workers and pages", async () => {
  const { run, sent } = setup();
  const res = await run("workday_search_workers", { query: "ada bob", max_results: 1 });
  assert.deepEqual(res.body.workers, [
    {
      id: ADA,
      name: "Ada Lovelace",
      worker_id: "Eaaa",
      title: "Engineer",
      job_profile: "Software Engineer",
      organization: { id: IT, name: "IT" },
      location: "Boston",
      worker_type: "Employee",
      email: "aaa@acme.example",
    },
  ]);
  assert.equal(res.body.next, "1");
  assert.ok(
    !res.text.includes("250000") && !res.text.includes("123-45-6789") && !res.text.includes("555"),
  );
  assert.equal(sent[1].auth, "Bearer tok1");
  assert.equal(sent[1].url.searchParams.get("search"), "ada bob");
  const next = await run("workday_search_workers", {
    query: "ada bob",
    max_results: 1,
    cursor: "1",
  });
  assert.equal(next.body.workers[0].id, BOB);
  assert.equal(next.body.next, null);
  assert.equal((await run("workday_search_workers", { query: "x", cursor: "abc" })).isError, true);
});

test("an organizations limit keeps searches to its organizations, sub-organizations and companies", async () => {
  const top = setup({ organizations: [TOP] });
  const names = async (h: ReturnType<typeof setup>) =>
    (
      await h.run("workday_search_workers", { query: "a b c d e", max_results: 50 })
    ).body.workers.map((w: { name: string }) => w.name);
  assert.deepEqual(await names(top), ["Ada Lovelace", "Bob Ross", "Carol Danvers"]);
  const company = setup({ organizations: [COMPANY] });
  assert.deepEqual(await names(company), ["Erin Hunt"]);
});

test("workday_get_worker checks the answer against the limit, and refuses unknown and bad ids alike", async () => {
  const { run, calls, audit } = setup({ organizations: [IT] });
  assert.equal((await run("workday_get_worker", { worker_id: BOB })).body.name, "Bob Ross");
  const outside = await run("workday_get_worker", { worker_id: DAN });
  assert.equal(outside.isError, true);
  assert.equal(outside.text, "This agent may not look up that worker");
  assert.equal(audit.at(-1)!.decision, "denied");
  const unknown = await run("workday_get_worker", { worker_id: id("9") });
  assert.equal(unknown.text, "This agent may not look up that worker");
  const before = calls().length;
  const bad = await run("workday_get_worker", { worker_id: "../../workers" });
  assert.match(bad.text, /Workday ID/);
  assert.equal(calls().length, before);
});

test("direct reports and the management chain stay inside the limit", async () => {
  const all = setup();
  const reports = await all.run("workday_list_direct_reports", { worker_id: CAROL });
  assert.deepEqual(
    reports.body.direct_reports.map((w: { id: string; is_manager: boolean }) => w.id),
    [],
  );
  const ada = await all.run("workday_list_direct_reports", { worker_id: ADA });
  assert.deepEqual(ada.body.direct_reports, [
    {
      id: BOB,
      name: "Bob Ross",
      worker_id: null,
      title: "Engineer",
      job_profile: null,
      organization: { id: IT, name: "IT" },
      location: null,
      worker_type: null,
      email: "bbb@acme.example",
      is_manager: false,
    },
  ]);
  const chain = await all.run("workday_get_management_chain", { worker_id: BOB });
  assert.deepEqual(chain.body, {
    worker: { id: BOB, name: "Bob Ross" },
    managers: [
      { id: ADA, name: "Ada Lovelace", organization: { id: IT, name: "IT" } },
      { id: CAROL, name: "Carol Danvers", organization: { id: TOP, name: "Company" } },
    ],
  });
  // Ada manages her own organization: her manager is the next one up.
  const adaChain = await all.run("workday_get_management_chain", { worker_id: ADA });
  assert.deepEqual(
    adaChain.body.managers.map((m: { id: string }) => m.id),
    [CAROL],
  );
  const it = setup({ organizations: [IT] });
  const limited = await it.run("workday_get_management_chain", { worker_id: BOB });
  assert.deepEqual(
    limited.body.managers.map((m: { id: string }) => m.id),
    [ADA],
  );
  assert.equal((await it.run("workday_list_direct_reports", { worker_id: DAN })).isError, true);
});

test("organizations are read only inside the limit; the superior outside it is left out", async () => {
  const { run } = setup({ organizations: [IT] });
  const it = await run("workday_get_organization", { org_id: IT });
  assert.deepEqual(it.body, {
    id: IT,
    name: "IT",
    code: "SUP-IT",
    manager: { id: ADA, name: "Ada Lovelace" },
    superior: null,
    sub_organizations: [],
  });
  assert.equal((await run("workday_get_organization", { org_id: TOP })).isError, true);
  assert.equal((await run("workday_list_organization_workers", { org_id: SALES })).isError, true);
  const workers = await run("workday_list_organization_workers", { org_id: IT });
  assert.deepEqual(
    workers.body.workers.map((w: { id: string }) => w.id),
    [ADA, BOB],
  );
  const top = await setup().run("workday_get_organization", { org_id: TOP });
  assert.deepEqual(
    top.body.sub_organizations.map((o: { name: string; managers: unknown[] }) => [
      o.name,
      o.managers,
    ]),
    [
      ["IT", [{ id: ADA, name: "Ada Lovelace" }]],
      ["Sales", [{ id: CAROL, name: "Carol Danvers" }]],
    ],
  );
});

test("time off: balances, entries and their checks", async () => {
  const { run, sent } = setup({ organizations: [IT] });
  const balances = await run("workday_get_time_off_balances", {
    worker_id: BOB,
    as_of: "2026-10-27",
  });
  assert.deepEqual(balances.body.balances, [
    { plan: "Vacation", category: "Time Off", balance: "12.5", unit: "Days", as_of: "2026-10-27" },
  ]);
  const q = sent.at(-1)!.url.searchParams;
  assert.equal(q.get("worker"), BOB);
  assert.equal(q.get("effective"), "2026-10-27");
  assert.equal((await run("workday_get_time_off_balances", { worker_id: DAN })).isError, true);
  assert.equal(
    (await run("workday_get_time_off_balances", { worker_id: BOB, as_of: "today" })).isError,
    true,
  );

  const entries = await run("workday_list_time_off", {
    worker_id: BOB,
    from: "2026-01-01",
    to: "2026-12-31",
  });
  assert.deepEqual(entries.body.entries, [
    {
      date: "2026-10-01",
      type: "Vacation",
      quantity: "1",
      unit: "Days",
      status: "Approved",
      comment: "Trip",
    },
  ]);
  const long = await run("workday_list_time_off", {
    worker_id: BOB,
    from: "2025-01-01",
    to: "2026-12-31",
  });
  assert.match(long.text, /at most a year/);
});

test("an answer about another worker's balances is refused", async () => {
  const { run } = setup(
    {},
    {
      "absenceManagement/v5/acme/balances": list([{ quantity: "3", worker: r(DAN, "Dan Brown") }]),
    },
  );
  const res = await run("workday_get_time_off_balances", { worker_id: BOB });
  assert.equal(res.isError, true);
  assert.match(res.text, /another worker/);
  assert.ok(!res.text.includes("3"));
});

test("workday_request_time_off submits the request, only for a worker inside the limit", async () => {
  const { run, sent, audit } = setup({ organizations: [IT] });
  const types = await run("workday_list_time_off_types", { worker_id: BOB });
  assert.deepEqual(types.body.types[0], {
    id: VACATION,
    name: "Vacation",
    group: "Paid",
    unit: "Hours",
    daily_default: "8",
    reason_required: false,
    start_and_end_time_required: false,
  });
  const res = await run("workday_request_time_off", {
    worker_id: BOB,
    days: [{ date: "2026-11-02", time_off_type_id: VACATION, quantity: 8 }],
    comment: "Family",
  });
  assert.deepEqual(res.body, {
    submitted: true,
    status: "Submitted",
    process: { id: id("0"), name: "Request Time Off: Bob Ross" },
    days: 1,
  });
  const post = sent.at(-1)!;
  assert.equal(post.method, "POST");
  assert.equal(
    post.url.pathname,
    `/ccx/api/absenceManagement/v5/acme/workers/${BOB}/requestTimeOff`,
  );
  assert.deepEqual(post.body, {
    businessProcessParameters: {
      action: { id: "d9e4223e446c11de98360015c5e6daf6" },
      comment: "Family",
    },
    days: [{ date: "2026-11-02", dailyQuantity: 8, timeOffType: { id: VACATION } }],
  });
  assert.equal(audit.at(-1)!.target, `${BOB} 2026-11-02`);

  const posts = () =>
    sent.filter((s) => s.method === "POST" && !s.url.pathname.endsWith("/token")).length;
  const outside = await run("workday_request_time_off", {
    worker_id: DAN,
    days: [{ date: "2026-11-02", time_off_type_id: VACATION, quantity: 8 }],
  });
  assert.equal(outside.isError, true);
  const badType = await run("workday_request_time_off", {
    worker_id: BOB,
    days: [{ date: "2026-11-02", time_off_type_id: "Vacation", quantity: 8 }],
  });
  assert.match(badType.text, /time off type id/);
  assert.equal(posts(), 1);
});

test("reports: listed and run within the key's reports, prompts checked, rows trimmed", async () => {
  const { run, sent, calls } = setup({ reports: ["Open_Reqs"] });
  const listed = await run("workday_list_reports", {});
  assert.deepEqual(listed.body.reports, [
    {
      report: "Open_Reqs",
      prompts: [
        { name: "Hiring_Org!WID", type: "text" },
        { name: "As_Of", type: "date" },
      ],
    },
  ]);
  const res = await run("workday_run_report", {
    report: "Open_Reqs",
    prompts: { "Hiring_Org!WID": TOP, As_Of: "2026-10-27" },
    max_rows: 2,
  });
  assert.deepEqual(res.body, {
    report: "Open_Reqs",
    rows: [
      { Requisition: "R-1 Engineer", Org_ID: IT, Openings: 2, Detail: '{"a":1}' },
      { Requisition: "R-2 Seller", Org_ID: SALES, Openings: 1 },
    ],
    truncated: true,
  });
  const url = sent.at(-1)!.url;
  assert.equal(url.pathname, "/ccx/service/customreport2/acme/jsmith/Open_Reqs");
  assert.equal(url.searchParams.get("Hiring_Org!WID"), TOP);
  assert.equal(url.searchParams.get("format"), "json");

  const before = calls().length;
  assert.match(
    (await run("workday_run_report", { report: "Headcount" })).text,
    /not a report this agent may run/,
  );
  assert.match(
    (await run("workday_run_report", { report: "Open_Reqs", prompts: { Salary: "1" } })).text,
    /no prompt Salary\. Its prompts: Hiring_Org!WID, As_Of/,
  );
  assert.match(
    (await run("workday_run_report", { report: "Open_Reqs", prompts: { As_Of: "Oct 1" } })).text,
    /As_Of as YYYY-MM-DD/,
  );
  assert.equal(calls().length, before);
});

test("reports under an organizations limit keep only rows of its organizations", async () => {
  const { run } = setup({ organizations: [TOP, COMPANY] });
  const res = await run("workday_run_report", { report: "Open_Reqs" });
  assert.deepEqual(
    res.body.rows.map((row: { Requisition: string }) => row.Requisition),
    ["R-1 Engineer", "R-2 Seller", "R-4 Company"],
  );
  const without = setup(
    { organizations: [TOP] },
    {},
    { ...connection.config, reportOrgColumn: "" },
  );
  const refused = await without.run("workday_run_report", { report: "Open_Reqs" });
  assert.match(refused.text, /no organization column set/);
  assert.deepEqual(without.calls(), []);
});

test("Workday's refusals reach the agent, with the capability's hint on a 403", async () => {
  const forbidden = setup(
    {},
    { "staffing/v7/acme/workers/b+": { status: 403, body: { error: "Access denied" } } },
  );
  const res = await forbidden.run("workday_list_time_off", {
    worker_id: BOB,
    from: "2026-01-01",
    to: "2026-02-01",
  });
  assert.equal(
    res.text,
    "Workday answered 403: Access denied. Give View on Worker Data: Public Worker Reports.",
  );
  assert.equal(forbidden.audit.at(-1)!.decision, "failed");
  const html = setup(
    {},
    { "customreport2/acme/jsmith/Open_Reqs": { body: "<html>sign in</html>" } },
  );
  assert.match(
    (await html.run("workday_run_report", { report: "Open_Reqs" })).text,
    /not a JSON report/,
  );
  const errors = setup(
    {},
    {
      "staffing/v7/acme/workers": {
        status: 400,
        body: { error: "Invalid request", errors: [{ error: "limit too big" }] },
      },
    },
  );
  assert.equal(
    (await errors.run("workday_search_workers", { query: "a" })).text,
    "Workday answered 400: Invalid request; limit too big",
  );
});

test("every call goes to the connected Workday host", async () => {
  const { run, sent } = setup({ organizations: [TOP] });
  await run("workday_get_management_chain", { worker_id: BOB });
  await run("workday_run_report", { report: "Open_Reqs" });
  assert.ok(sent.length > 3);
  assert.ok(sent.every((s) => s.url.host === HOST && s.url.pathname.startsWith("/ccx/")));
});

test("the tenant check lists one worker and names the tenant; discover lists the reports", async () => {
  const { vendor, sent } = setup();
  assert.deepEqual(await vendor.checks!.tenant(connection, { expected: [] }), {
    values: [`acme (${HOST})`],
  });
  assert.equal(sent.at(-1)!.url.searchParams.get("limit"), "1");
  assert.deepEqual((await vendor.checks!.reports(connection, { expected: [] })).values, [
    "Open_Reqs",
    "Headcount",
  ]);
  const refused = setup({}, { "staffing/v7/acme/workers": { status: 403, body: {} } });
  await assert.rejects(
    refused.vendor.checks!.tenant(connection, { expected: [] }),
    /Public Worker Reports/,
  );
});
