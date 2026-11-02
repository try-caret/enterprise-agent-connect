import { z } from "zod";
import { connectedReports } from "./vendor.js";
import {
  audited,
  Denied,
  json,
  listLimit,
  VendorFailed,
  type ToolContext,
  type Toolset,
} from "../../packages/connector-kit/src/common.js";

// Workday REST (staffing v7, common v1, absenceManagement v5) and custom reports. Paths are
// Workday's REST directory; answer fields follow CData's tables generated from Workday's specs
// (Workers, WorkersDirectReports, WorkersOrganizations, SupervisoryOrganizations,
// SupervisoryOrganizationsOrgChart, Balances, WorkersTimeOffDetails,
// WorkersEligibleAbsenceTypes, WorkersRequestTimeOff; read October 27, 2026).

const WID = /^[0-9a-f]{32}$/;
const DATE = /^\d{4}-\d{2}-\d{2}$/;
// Workday's submit action for a business process (CData's WorkersRequestTimeOff).
const SUBMIT = "d9e4223e446c11de98360015c5e6daf6";
// Supervisory levels walked up for a management chain or an organization limit.
const MAX_LEVELS = 12;
const MAX_ROWS = 200;
const OUTSIDE = "This agent may not look up that worker";

type Mark = (status: number) => void;
type Ref = { id?: string; descriptor?: string } | undefined;

const api = (ctx: ToolContext, service: string, rest: string) =>
  `/ccx/api/${service}/${ctx.config.tenant}${rest}`;

export function workdayError(res: { status: number; data: unknown }): string {
  const d = res.data as { error?: unknown; errors?: { error?: unknown }[] } | string | null;
  const detail =
    typeof d === "string"
      ? d.slice(0, 200)
      : [d?.error, ...(Array.isArray(d?.errors) ? d.errors.map((e) => e?.error) : [])]
          .filter((e) => typeof e === "string")
          .join("; ")
          .slice(0, 500);
  return `Workday answered ${res.status}${detail ? `: ${detail}` : ""}`;
}

async function ok(ctx: ToolContext, mark: Mark, path: string, body?: unknown, hintTool?: string) {
  const res = await ctx.call(
    body === undefined ? { method: "GET", path } : { method: "POST", path, body },
  );
  mark(res.status);
  if (res.status < 200 || res.status >= 300)
    throw new VendorFailed(res.status, workdayError(res), hintTool);
  return res.data as any;
}

const ref = (r: Ref) => (r?.id ? { id: r.id, name: r.descriptor ?? null } : null);

// Staffing's worker (primaryJob, person) and common's summary (direct reports) share these.
export function worker(w: any) {
  return {
    id: w.id,
    name: w.descriptor ?? null,
    worker_id: w.workerId ?? null,
    title: w.primaryJob?.businessTitle ?? w.businessTitle ?? null,
    job_profile: w.primaryJob?.jobProfile?.descriptor ?? null,
    organization: ref(w.primaryJob?.supervisoryOrganization ?? w.primarySupervisoryOrganization),
    location: w.primaryJob?.location?.descriptor ?? null,
    worker_type: w.workerType?.descriptor ?? null,
    email: w.person?.email ?? w.primaryWorkEmail ?? null,
    ...(typeof w.isManager === "boolean" ? { is_manager: w.isManager } : {}),
  };
}

const page = (data: any, offset: number, limit: number) => {
  const total = typeof data?.total === "number" ? data.total : 0;
  return offset + limit < total ? String(offset + limit) : null;
};

const offsetOf = (cursor: string | undefined) => {
  if (cursor === undefined) return 0;
  if (!/^\d{1,6}$/.test(cursor)) throw new Denied("Pass back the cursor a listing returned");
  return Number(cursor);
};

// The key's organizations limit: Workday IDs of supervisory organizations (their
// sub-organizations included) or of any other organization a worker belongs to, e.g. a company.
// Memoized per call.
export function scope(ctx: ToolContext, mark: Mark) {
  const limit = listLimit(ctx.limits, "organizations");
  const charts = new Map<string, Promise<any>>();

  // Staffing's org chart of a supervisory organization: its superior and subordinates. null
  // when the id is not a supervisory organization.
  const chart = (id: string) => {
    if (!charts.has(id))
      charts.set(
        id,
        (async () => {
          const res = await ctx.call({
            method: "GET",
            path: api(ctx, "staffing/v7", `/supervisoryOrganizations/${id}/orgChart`),
          });
          mark(res.status);
          if (res.status === 400 || res.status === 404) return null;
          if (res.status < 200 || res.status >= 300)
            throw new VendorFailed(res.status, workdayError(res));
          return (res.data as any)?.data?.[0] ?? null;
        })(),
      );
    return charts.get(id)!;
  };

  async function org(id: string | undefined): Promise<boolean> {
    if (!limit) return true;
    for (let level = 0; id && WID.test(id) && level < MAX_LEVELS; level++) {
      if (limit.includes(id)) return true;
      id = (await chart(id))?.superior?.id;
    }
    return false;
  }

  async function worker(w: any): Promise<boolean> {
    if (!limit) return true;
    const supervisory = (w.primaryJob?.supervisoryOrganization ?? w.primarySupervisoryOrganization)
      ?.id;
    if (await org(supervisory)) return true;
    const orgs = await ok(
      ctx,
      mark,
      api(ctx, "common/v1", `/workers/${w.id}/organizations?limit=100`),
    );
    return (orgs?.data ?? []).some((o: Ref) => o?.id && limit.includes(o.id));
  }

  // A worker read by id, refused (as if not found) when outside the key's organizations. It is
  // the worker lookup capability's permission, whichever tool reads.
  async function read(id: string) {
    if (!WID.test(id)) throw new Denied("Give a worker's Workday ID (32 letters and digits)");
    const res = await ctx.call({ method: "GET", path: api(ctx, "staffing/v7", `/workers/${id}`) });
    mark(res.status);
    if (limit && res.status === 404) throw new Denied(OUTSIDE);
    if (res.status < 200 || res.status >= 300)
      throw new VendorFailed(res.status, workdayError(res), "workday_get_worker");
    if (!(await worker(res.data))) throw new Denied(OUTSIDE);
    return res.data as any;
  }

  return { limited: limit !== null, org, worker, read, chart };
}

const wid = (what: string) =>
  z.string().describe(`The ${what}'s Workday ID (32 letters and digits).`);
const cursor = z.string().optional().describe("The `next` value of the previous page.");

export const workdayTools: Toolset = {
  names: [
    "workday_search_workers",
    "workday_get_worker",
    "workday_list_direct_reports",
    "workday_get_management_chain",
    "workday_get_organization",
    "workday_list_organization_workers",
    "workday_get_time_off_balances",
    "workday_list_time_off",
    "workday_list_time_off_types",
    "workday_request_time_off",
    "workday_list_reports",
    "workday_run_report",
  ],

  register(server, ctx, allowed) {
    const readOnly = { readOnlyHint: true, openWorldHint: true } as const;
    const tool = (
      name: string,
      description: string,
      inputSchema: z.ZodObject,
      target: (args: any) => string | null,
      run: (args: any, mark: Mark, retarget: (s: string) => void) => Promise<unknown>,
      annotations: object = readOnly,
    ) => {
      if (allowed.has(name))
        server.registerTool(name, { description, inputSchema, annotations }, (args: any) =>
          audited(ctx, name, target(args), async (mark, retarget) =>
            json(await run(args, mark, retarget)),
          ),
        );
    };

    tool(
      "workday_search_workers",
      "Find workers by name or worker ID (Workday's search: case-insensitive, space-separated words match any). Returns name, title, organization, location and work email; never pay or personal data.",
      z.object({
        query: z.string().min(1).max(100),
        max_results: z.number().int().min(1).max(50).default(10),
        cursor,
      }),
      (a) => a.query,
      async ({ query, max_results, cursor: c }, mark) => {
        const offset = offsetOf(c);
        const s = scope(ctx, mark);
        const q = new URLSearchParams({
          search: query,
          limit: String(max_results),
          offset: String(offset),
        });
        const data = await ok(ctx, mark, api(ctx, "staffing/v7", `/workers?${q}`));
        const found = [];
        for (const w of data?.data ?? []) if (await s.worker(w)) found.push(worker(w));
        return { workers: found, next: page(data, offset, max_results) };
      },
    );

    tool(
      "workday_get_worker",
      "Read one worker: name, worker ID, title, job profile, supervisory organization, location, worker type and work email.",
      z.object({ worker_id: wid("worker") }),
      (a) => a.worker_id,
      async ({ worker_id }, mark) => worker(await scope(ctx, mark).read(worker_id)),
    );

    tool(
      "workday_list_direct_reports",
      "List a worker's direct reports.",
      z.object({ worker_id: wid("manager"), cursor }),
      (a) => a.worker_id,
      async ({ worker_id, cursor: c }, mark) => {
        const offset = offsetOf(c);
        const s = scope(ctx, mark);
        await s.read(worker_id);
        const data = await ok(
          ctx,
          mark,
          api(ctx, "common/v1", `/workers/${worker_id}/directReports?limit=100&offset=${offset}`),
        );
        const reports = [];
        for (const w of data?.data ?? []) if (await s.worker(w)) reports.push(worker(w));
        return { direct_reports: reports, next: page(data, offset, 100) };
      },
    );

    tool(
      "workday_get_management_chain",
      "A worker's managers, nearest first, up the supervisory organizations. Stops at the top, or where this agent's organizations end.",
      z.object({ worker_id: wid("worker") }),
      (a) => a.worker_id,
      async ({ worker_id }, mark) => {
        const s = scope(ctx, mark);
        const w = await s.read(worker_id);
        const chain: { id: string; name: string | null; organization: unknown }[] = [];
        let orgId: string | undefined = w.primaryJob?.supervisoryOrganization?.id;
        for (let level = 0; orgId && WID.test(orgId) && level < MAX_LEVELS; level++) {
          if (!(await s.org(orgId))) break;
          const o = await ok(
            ctx,
            mark,
            api(ctx, "staffing/v7", `/supervisoryOrganizations/${orgId}`),
          );
          const m: Ref = o?.manager;
          const last = chain.length ? chain[chain.length - 1].id : worker_id;
          if (m?.id && m.id !== last)
            chain.push({ id: m.id, name: m.descriptor ?? null, organization: ref(o) });
          orgId = (await s.chart(orgId))?.superior?.id;
        }
        return { worker: ref(w), managers: chain };
      },
    );

    tool(
      "workday_get_organization",
      "Read a supervisory organization: its name, code, manager, superior organization and sub-organizations.",
      z.object({ org_id: wid("supervisory organization") }),
      (a) => a.org_id,
      async ({ org_id }, mark) => {
        if (!WID.test(org_id)) throw new Denied("Give an organization's Workday ID");
        const s = scope(ctx, mark);
        if (!(await s.org(org_id)))
          throw new Denied("This agent may not look up that organization");
        const o = await ok(
          ctx,
          mark,
          api(ctx, "staffing/v7", `/supervisoryOrganizations/${org_id}`),
        );
        const c = await s.chart(org_id);
        const superior = c?.superior?.id && (await s.org(c.superior.id)) ? ref(c.superior) : null;
        return {
          id: o.id,
          name: o.name ?? o.descriptor ?? null,
          code: o.code ?? null,
          manager: ref(o.manager),
          superior,
          sub_organizations: (c?.subordinates ?? []).map((x: any) => ({
            id: x.id,
            name: x.name ?? x.descriptor ?? null,
            code: x.code ?? null,
            inactive: x.inactive ?? false,
            managers: (x.managers ?? []).map(ref).filter(Boolean),
          })),
        };
      },
    );

    tool(
      "workday_list_organization_workers",
      "List the workers in a supervisory organization.",
      z.object({ org_id: wid("supervisory organization"), cursor }),
      (a) => a.org_id,
      async ({ org_id, cursor: c }, mark) => {
        if (!WID.test(org_id)) throw new Denied("Give an organization's Workday ID");
        const offset = offsetOf(c);
        if (!(await scope(ctx, mark).org(org_id)))
          throw new Denied("This agent may not look up that organization");
        const data = await ok(
          ctx,
          mark,
          api(
            ctx,
            "common/v1",
            `/supervisoryOrganizations/${org_id}/workers?limit=100&offset=${offset}`,
          ),
        );
        return { workers: (data?.data ?? []).map(worker), next: page(data, offset, 100) };
      },
    );

    tool(
      "workday_get_time_off_balances",
      "A worker's time off and leave balances, per plan, as of today or a given date.",
      z.object({
        worker_id: wid("worker"),
        as_of: z.string().optional().describe("YYYY-MM-DD; today when left out."),
      }),
      (a) => a.worker_id,
      async ({ worker_id, as_of }, mark) => {
        if (as_of !== undefined && !DATE.test(as_of)) throw new Denied("Give as_of as YYYY-MM-DD");
        await scope(ctx, mark).read(worker_id);
        const q = new URLSearchParams({
          worker: worker_id,
          limit: "100",
          ...(as_of ? { effective: as_of } : {}),
        });
        const data = await ok(ctx, mark, api(ctx, "absenceManagement/v5", `/balances?${q}`));
        const rows = (data?.data ?? []) as any[];
        // The filter is Workday's; an answer about anyone else is never passed on.
        if (rows.some((b) => b?.worker?.id !== worker_id))
          throw new Denied("Workday answered with another worker's balances");
        return {
          balances: rows.map((b) => ({
            plan: b.absencePlan?.descriptor ?? null,
            category: b.category?.descriptor ?? null,
            balance: b.quantity ?? null,
            unit: b.unit?.descriptor ?? null,
            as_of: b.effectiveDate ?? null,
          })),
        };
      },
    );

    tool(
      "workday_list_time_off",
      "A worker's time off entries (requested, approved and taken) between two dates, at most a year apart.",
      z.object({
        worker_id: wid("worker"),
        from: z.string().describe("YYYY-MM-DD"),
        to: z.string().describe("YYYY-MM-DD"),
        cursor,
      }),
      (a) => `${a.worker_id} ${a.from}..${a.to}`,
      async ({ worker_id, from, to, cursor: c }, mark) => {
        if (!DATE.test(from) || !DATE.test(to)) throw new Denied("Give from and to as YYYY-MM-DD");
        const days = (Date.parse(to) - Date.parse(from)) / 86_400_000;
        if (!(days >= 0 && days <= 366))
          throw new Denied("from must be before to, at most a year apart");
        const offset = offsetOf(c);
        await scope(ctx, mark).read(worker_id);
        const q = new URLSearchParams({
          fromDate: from,
          toDate: to,
          limit: "100",
          offset: String(offset),
        });
        const data = await ok(
          ctx,
          mark,
          api(ctx, "absenceManagement/v5", `/workers/${worker_id}/timeOffDetails?${q}`),
        );
        return {
          entries: (data?.data ?? []).map((e: any) => ({
            date: e.date ?? null,
            type: e.timeOffType?.descriptor ?? null,
            quantity: e.quantity ?? null,
            unit: e.unit?.descriptor ?? null,
            status: e.status?.descriptor ?? null,
            comment: e.comment ?? null,
          })),
          next: page(data, offset, 100),
        };
      },
    );

    tool(
      "workday_list_time_off_types",
      "The time off types a worker may request, with their ids for workday_request_time_off, unit and default daily quantity.",
      z.object({ worker_id: wid("worker") }),
      (a) => a.worker_id,
      async ({ worker_id }, mark) => {
        await scope(ctx, mark).read(worker_id);
        const data = await ok(
          ctx,
          mark,
          api(ctx, "absenceManagement/v5", `/workers/${worker_id}/eligibleAbsenceTypes?limit=100`),
        );
        return {
          types: (data?.data ?? []).map((t: any) => ({
            id: t.id,
            name: t.descriptor ?? null,
            group: t.absenceTypeGroup?.descriptor ?? null,
            unit: t.unitOfTime?.descriptor ?? null,
            daily_default: t.dailyDefaultQuantity ?? null,
            reason_required: t.reasonRequired ?? false,
            start_and_end_time_required: t.startAndEndTimeRequired ?? false,
          })),
        };
      },
    );

    tool(
      "workday_request_time_off",
      "Submit a time off request for a worker. It starts Workday's Request Time Off approval, as the worker's request; the worker and their manager are notified as Workday is set up to do. Get type ids from workday_list_time_off_types.",
      z.object({
        worker_id: wid("worker"),
        days: z
          .array(
            z.object({
              date: z.string().describe("YYYY-MM-DD"),
              time_off_type_id: z.string(),
              quantity: z
                .number()
                .positive()
                .max(24)
                .describe("Hours or days, in the type's unit."),
              comment: z.string().max(500).optional(),
            }),
          )
          .min(1)
          .max(31),
        comment: z.string().max(1000).optional().describe("A comment on the whole request."),
      }),
      (a) => a.worker_id,
      async ({ worker_id, days, comment }, mark, retarget) => {
        for (const d of days) {
          if (!DATE.test(d.date)) throw new Denied("Give each date as YYYY-MM-DD");
          if (!WID.test(d.time_off_type_id))
            throw new Denied("Give a time off type id from workday_list_time_off_types");
        }
        await scope(ctx, mark).read(worker_id);
        retarget(`${worker_id} ${days.map((d: any) => d.date).join(",")}`.slice(0, 300));
        const data = await ok(
          ctx,
          mark,
          api(ctx, "absenceManagement/v5", `/workers/${worker_id}/requestTimeOff`),
          {
            businessProcessParameters: {
              action: { id: SUBMIT },
              ...(comment ? { comment } : {}),
            },
            days: days.map((d: any) => ({
              date: d.date,
              dailyQuantity: d.quantity,
              timeOffType: { id: d.time_off_type_id },
              ...(d.comment ? { comment: d.comment } : {}),
            })),
          },
        );
        const bp = data?.businessProcessParameters;
        return {
          submitted: true,
          status: bp?.overallStatus ?? bp?.transactionStatus?.descriptor ?? null,
          process: ref(bp?.overallBusinessProcess),
          days: days.length,
        };
      },
      { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
    );

    const reports = () => {
      const limit = listLimit(ctx.limits, "reports");
      return connectedReports(ctx.config).filter((r) => !limit || limit.includes(r.name));
    };

    tool(
      "workday_list_reports",
      "The Workday reports this agent may run, with each report's prompts (filters) and their types.",
      z.object({}),
      () => null,
      async () => ({
        reports: reports().map((r) => ({ report: r.name, prompts: r.prompts })),
      }),
    );

    tool(
      "workday_run_report",
      "Run one of the Workday reports from workday_list_reports, with values for any of its prompts (dates as YYYY-MM-DD). Returns its rows, at most max_rows.",
      z.object({
        report: z.string(),
        prompts: z.record(z.string(), z.string().max(200)).default({}),
        max_rows: z.number().int().min(1).max(MAX_ROWS).default(50),
      }),
      (a) => a.report,
      async ({ report, prompts, max_rows }, mark, retarget) => {
        const r = reports().find((x) => x.name === report);
        if (!r) throw new Denied(`${report} is not a report this agent may run`);
        const q = new URLSearchParams();
        for (const [name, value] of Object.entries(prompts as Record<string, string>)) {
          const p = r.prompts.find((x) => x.name === name);
          if (!p)
            throw new Denied(
              `${report} has no prompt ${name}. Its prompts: ${r.prompts.map((x) => x.name).join(", ") || "none"}`,
            );
          if (p.type === "date" && !DATE.test(value))
            throw new Denied(`Give ${name} as YYYY-MM-DD`);
          q.append(name, value);
        }
        q.set("format", "json");
        const s = scope(ctx, mark);
        const column = ctx.config.reportOrgColumn;
        if (s.limited && !column)
          throw new Denied(
            "This agent is limited to some organizations, and reports have no organization column set, so it may not run reports",
          );
        retarget(`${report}?${q}`.slice(0, 300));
        const enc = encodeURIComponent;
        const data = await ok(
          ctx,
          mark,
          `/ccx/service/customreport2/${ctx.config.tenant}/${enc(r.owner)}/${enc(r.report)}?${q}`,
        );
        if (!Array.isArray(data?.Report_Entry))
          throw new VendorFailed(200, "Workday's answer was not a JSON report");
        const rows = [];
        for (const row of data.Report_Entry as Record<string, unknown>[]) {
          if (s.limited) {
            const v = row[column as string];
            const ids = (Array.isArray(v) ? v : [v]).filter(
              (x): x is string => typeof x === "string",
            );
            let inside = false;
            for (const id of ids) if ((inside = await s.org(id))) break;
            if (!inside) continue;
          }
          rows.push(trimRow(row));
          if (rows.length === max_rows) break;
        }
        return { report, rows, truncated: rows.length === max_rows };
      },
    );
  },
};

export function trimRow(row: Record<string, unknown>) {
  return Object.fromEntries(
    Object.entries(row).map(([k, v]) => [
      k,
      typeof v === "string"
        ? v.slice(0, 500)
        : typeof v === "number" || typeof v === "boolean" || v === null
          ? v
          : JSON.stringify(v).slice(0, 500),
    ]),
  );
}
