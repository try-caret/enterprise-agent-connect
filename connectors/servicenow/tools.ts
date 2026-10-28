import { z } from "zod";
import {
  audited,
  Denied,
  json,
  listLimit,
  VendorFailed,
  type ToolContext,
  type Toolset,
} from "../../packages/connector-kit/src/common.js";

// ServiceNow's Table API (/api/now/table/<table>) and Service Catalog API
// (/api/sn_sc/servicecatalog). Agents never write an encoded query: each tool builds one from
// its own parameters and the key's limits, and every value is refused if it carries `^` (which
// starts a new condition, or with NQ a whole new query) or `javascript:` (which ServiceNow
// runs when it reads a query value).

const TABLE = "/api/now/table";
const CATALOG = "/api/sn_sc/servicecatalog";
const NUMBER = /^[A-Z]{2,8}[0-9]{4,12}$/;
const SYS_ID = /^[0-9a-f]{32}$/;
const CURSOR = /^[0-9]{1,6}$/;
const JOURNAL_CHARS = 8000;
const ARTICLE_CHARS = 20000;

type Field = { value?: unknown; display_value?: unknown } | undefined;
type Rec = Record<string, Field>;
type Mark = (s: number) => void;

// With sysparm_display_value=all every field comes as { value, display_value }: references
// and choices are read by their display value (a group's name, "In Progress"), times by their
// value, which the Table API gives in UTC.
const text = (f: Field, key: "value" | "display_value") =>
  typeof f?.[key] === "string" && f[key] !== "" ? (f[key] as string) : null;
const shown = (r: Rec, field: string) => text(r[field], "display_value");
const raw = (r: Rec, field: string) => text(r[field], "value");
const utc = (r: Rec, field: string) => raw(r, field)?.replace(" ", "T").concat("Z") ?? null;

export function safe(value: string, what = "A value"): string {
  if (!value.trim() || /[\^\r\n]/.test(value) || /javascript:/i.test(value))
    throw new Denied(`${what} cannot be empty or contain ^, line breaks or javascript:`);
  return value;
}

// field=a^ORfield=b: ^OR binds to the condition before it, so this stays one condition when
// joined to others with ^.
const anyOf = (field: string, values: string[]) =>
  values.map((v, i) => `${i ? "^OR" : ""}${field}=${safe(v)}`).join("");

const userIs = (field: string, who: string) =>
  `${field}.email=${safe(who, "A person")}^OR${field}.user_name=${who}`;

// What the key's limits mean on one table: the field holding the assignment group and the
// field holding the company, where that limit applies.
type Scoped = { group?: string; company?: string };

function scopeQuery(ctx: ToolContext, on: Scoped): string[] {
  const groups = listLimit(ctx.limits, "groups");
  const companies = listLimit(ctx.limits, "companies");
  return [
    ...(groups && on.group ? [anyOf(`${on.group}.name`, groups)] : []),
    ...(companies && on.company ? [anyOf(`${on.company}.name`, companies)] : []),
  ];
}

const within = (values: string[] | null, value: string | null) =>
  !values || (value !== null && values.some((v) => v.toLowerCase() === value.toLowerCase()));

// The same limits, checked again on each record answered.
function inScope(ctx: ToolContext, on: Scoped, r: Rec): boolean {
  return (
    (!on.group || within(listLimit(ctx.limits, "groups"), shown(r, on.group))) &&
    (!on.company || within(listLimit(ctx.limits, "companies"), shown(r, on.company)))
  );
}

async function ok(ctx: ToolContext, mark: Mark, req: Parameters<ToolContext["call"]>[0]) {
  const res = await ctx.call(req);
  mark(res.status);
  if (res.status < 200 || res.status >= 300) {
    const e = (res.data as { error?: { message?: unknown; detail?: unknown } } | null)?.error;
    const detail = [e?.message, e?.detail].filter((x) => typeof x === "string" && x).join(": ");
    throw new VendorFailed(
      res.status,
      `ServiceNow answered ${res.status}${detail ? `: ${detail.slice(0, 400)}` : ""}`,
    );
  }
  return res.data as { result?: any };
}

function tablePath(
  table: string,
  query: string[],
  fields: string[],
  extra: Record<string, string> = {},
): string {
  const params = new URLSearchParams({
    ...(query.length ? { sysparm_query: query.join("^") } : {}),
    sysparm_fields: fields.join(","),
    sysparm_display_value: "all",
    sysparm_exclude_reference_link: "true",
    ...extra,
  });
  return `${TABLE}/${table}?${params}`;
}

function offset(cursor: string | undefined): number {
  if (cursor === undefined) return 0;
  if (!CURSOR.test(cursor)) throw new Denied("Give the cursor exactly as the last answer's next");
  return Number(cursor);
}

// One page: asks for one more row than wanted, so `next` is only given when there is more.
async function page(
  ctx: ToolContext,
  mark: Mark,
  table: string,
  query: string[],
  fields: string[],
  max: number,
  cursor: string | undefined,
  on: Scoped,
) {
  const start = offset(cursor);
  const data = await ok(ctx, mark, {
    method: "GET",
    path: tablePath(table, query, ["sys_id", ...fields], {
      sysparm_limit: String(max + 1),
      sysparm_offset: String(start),
    }),
  });
  const rows = (data.result ?? []) as Rec[];
  return {
    rows: rows.slice(0, max).filter((r) => inScope(ctx, on, r)),
    next: rows.length > max ? String(start + max) : null,
  };
}

// A record by its number. One outside the key's limits is answered like one that does not
// exist, so a key cannot learn which numbers exist elsewhere.
async function byNumber(
  ctx: ToolContext,
  mark: Mark,
  table: string,
  kind: string,
  number: string,
  fields: string[],
  on: Scoped,
): Promise<Rec> {
  if (!NUMBER.test(number)) throw new Denied(`Give the ${kind} number, such as ${example[table]}`);
  const data = await ok(ctx, mark, {
    method: "GET",
    path: tablePath(table, [`number=${number}`, ...scopeQuery(ctx, on)], ["sys_id", ...fields], {
      sysparm_limit: "1",
    }),
  });
  const r = (data.result ?? [])[0] as Rec | undefined;
  if (!r || !inScope(ctx, on, r) || !SYS_ID.test(raw(r, "sys_id") ?? ""))
    throw new Denied(`${number} was not found, or this agent may not use it`);
  return r;
}

const example: Record<string, string> = {
  incident: "INC0010001",
  change_request: "CHG0030001",
  kb_knowledge: "KB0010001",
};

// One active user by email or user name. `scoped`: under a companies limit, only a person of
// those companies (a caller or someone ordered for, not the staff member assigned).
async function person(ctx: ToolContext, mark: Mark, who: string, scoped: boolean) {
  const query = ["active=true", `email=${safe(who, "A person")}^ORuser_name=${who}`];
  if (scoped) query.push(...scopeQuery(ctx, { company: "company" }));
  const data = await ok(ctx, mark, {
    method: "GET",
    path: tablePath("sys_user", query, ["sys_id", "name", "company"], {
      sysparm_limit: "2",
    }),
  });
  const rows = (data.result ?? []) as Rec[];
  const r = rows[0];
  if (rows.length !== 1 || (scoped && !inScope(ctx, { company: "company" }, r)))
    throw new Denied(`No single active person ${who} that this agent may act for`);
  return { id: raw(r, "sys_id")!, company: raw(r, "company") };
}

async function group(ctx: ToolContext, mark: Mark, name: string) {
  if (!within(listLimit(ctx.limits, "groups"), safe(name, "A group name")))
    throw new Denied(`This agent may not use the assignment group ${name}`);
  const data = await ok(ctx, mark, {
    method: "GET",
    path: tablePath("sys_user_group", ["active=true", `name=${name}`], ["sys_id"], {
      sysparm_limit: "2",
    }),
  });
  const rows = (data.result ?? []) as Rec[];
  if (rows.length !== 1) throw new Denied(`No single active assignment group is named ${name}`);
  return raw(rows[0], "sys_id")!;
}

const clip = (s: string | null, max: number) =>
  s && s.length > max ? `${s.slice(0, max)}\n[cut at ${max} characters]` : s;

// Knowledge articles and catalog descriptions are HTML; agents get the text.
export function htmlText(html: string | null): string {
  if (!html) return "";
  return html
    .replace(/<(script|style)[\s\S]*?<\/\1>/gi, "")
    .replace(/<br\s*\/?>|<\/(p|div|li|h[1-6]|tr)>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&amp;/g, "&")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

// Incidents.
const INCIDENT: Scoped = { group: "assignment_group", company: "company" };
const INCIDENT_LIST = [
  "number",
  "short_description",
  "state",
  "priority",
  "assignment_group",
  "assigned_to",
  "caller_id",
  "company",
  "opened_at",
  "sys_updated_on",
];
const INCIDENT_FULL = [
  ...INCIDENT_LIST,
  "description",
  "category",
  "impact",
  "urgency",
  "cmdb_ci",
  "close_code",
  "close_notes",
  "resolved_at",
  "comments",
  "work_notes",
];
// The incident table's state values (out of the box).
const INCIDENT_STATE = {
  new: "1",
  in_progress: "2",
  on_hold: "3",
  resolved: "6",
  closed: "7",
};
const HOLD_REASON = {
  awaiting_caller: "1",
  awaiting_problem: "3",
  awaiting_vendor: "4",
  awaiting_change: "5",
};

const incidentSummary = (r: Rec) => ({
  number: shown(r, "number"),
  short_description: shown(r, "short_description"),
  state: shown(r, "state"),
  priority: shown(r, "priority"),
  assignment_group: shown(r, "assignment_group"),
  assigned_to: shown(r, "assigned_to"),
  caller: shown(r, "caller_id"),
  company: shown(r, "company"),
  opened: utc(r, "opened_at"),
  updated: utc(r, "sys_updated_on"),
});

const incidentDetail = (r: Rec) => ({
  ...incidentSummary(r),
  description: shown(r, "description"),
  category: shown(r, "category"),
  impact: shown(r, "impact"),
  urgency: shown(r, "urgency"),
  configuration_item: shown(r, "cmdb_ci"),
  close_code: shown(r, "close_code"),
  close_notes: shown(r, "close_notes"),
  resolved: utc(r, "resolved_at"),
  comments: clip(shown(r, "comments"), JOURNAL_CHARS),
  work_notes: clip(shown(r, "work_notes"), JOURNAL_CHARS),
});

// Changes.
const CHANGE: Scoped = { group: "assignment_group", company: "company" };
const CHANGE_LIST = [
  "number",
  "short_description",
  "type",
  "state",
  "risk",
  "assignment_group",
  "assigned_to",
  "company",
  "start_date",
  "end_date",
  "sys_updated_on",
];
const CHANGE_FULL = [
  ...CHANGE_LIST,
  "description",
  "impact",
  "approval",
  "cmdb_ci",
  "justification",
  "implementation_plan",
  "backout_plan",
  "test_plan",
  "close_code",
  "close_notes",
];
const changeSummary = (r: Rec) => ({
  number: shown(r, "number"),
  short_description: shown(r, "short_description"),
  type: shown(r, "type"),
  state: shown(r, "state"),
  risk: shown(r, "risk"),
  assignment_group: shown(r, "assignment_group"),
  assigned_to: shown(r, "assigned_to"),
  company: shown(r, "company"),
  planned_start: utc(r, "start_date"),
  planned_end: utc(r, "end_date"),
  updated: utc(r, "sys_updated_on"),
});

// Configuration items and requested items: only the companies limit applies.
const CI: Scoped = { company: "company" };
const CI_FIELDS = [
  "name",
  "sys_class_name",
  "serial_number",
  "asset_tag",
  "install_status",
  "operational_status",
  "assigned_to",
  "support_group",
  "company",
  "location",
  "ip_address",
  "sys_updated_on",
];
const RITM: Scoped = { company: "requested_for.company" };
const RITM_FIELDS = [
  "number",
  "request",
  "cat_item",
  "stage",
  "state",
  "requested_for",
  "requested_for.company",
  "assignment_group",
  "opened_at",
  "due_date",
  "sys_updated_on",
];

const who = z.string().max(100);
const max = (top: number, byDefault: number) => z.number().int().min(1).max(top).default(byDefault);
const cursor = z
  .string()
  .optional()
  .describe("The `next` value of the previous answer, to read the following page.");
const status = z
  .enum(["open", "closed", "any"])
  .default("open")
  .describe("open: still active; closed: resolved, closed or cancelled.");

const statusQuery = (s: "open" | "closed" | "any") =>
  s === "open" ? ["active=true"] : s === "closed" ? ["active=false"] : [];
const textQuery = (t: string | undefined) =>
  t ? [`short_descriptionLIKE${safe(t, "Search text")}^ORdescriptionLIKE${t}`] : [];

export const servicenowTools: Toolset = {
  names: [
    "servicenow_search_incidents",
    "servicenow_get_incident",
    "servicenow_create_incident",
    "servicenow_update_incident",
    "servicenow_add_incident_note",
    "servicenow_resolve_incident",
    "servicenow_search_changes",
    "servicenow_get_change",
    "servicenow_find_configuration_items",
    "servicenow_search_knowledge",
    "servicenow_get_article",
    "servicenow_search_catalog_items",
    "servicenow_get_catalog_item",
    "servicenow_order_catalog_item",
    "servicenow_search_requests",
  ],

  register(server, ctx, allowed) {
    const readOnly = { readOnlyHint: true, openWorldHint: true } as const;
    const write = {
      readOnlyHint: false,
      destructiveHint: false,
      openWorldHint: true,
    } as const;
    const groupNote = listLimit(ctx.limits, "groups")
      ? ` This agent may only use these assignment groups: ${listLimit(ctx.limits, "groups")!.join(", ")}.`
      : "";

    if (allowed.has("servicenow_search_incidents"))
      server.registerTool(
        "servicenow_search_incidents",
        {
          description:
            "Search ServiceNow incidents, newest change first. All filters are optional and combine. " +
            "When `next` is set, call again with `cursor` to read more.",
          inputSchema: z.object({
            text: z
              .string()
              .max(200)
              .optional()
              .describe("Words in the short description or description."),
            status,
            priority: z
              .number()
              .int()
              .min(1)
              .max(5)
              .optional()
              .describe("1 Critical to 5 Planning."),
            assignment_group: z
              .string()
              .max(100)
              .optional()
              .describe(`Assignment group name.${groupNote}`),
            caller: who.optional().describe("The caller's email or user name."),
            assigned_to: who.optional().describe("The assignee's email or user name."),
            max_results: max(50, 20),
            cursor,
          }),
          annotations: readOnly,
        },
        (args) =>
          audited(
            ctx,
            "servicenow_search_incidents",
            JSON.stringify(args).slice(0, 300),
            async (mark, retarget) => {
              if (
                args.assignment_group &&
                !within(listLimit(ctx.limits, "groups"), args.assignment_group)
              )
                throw new Denied(
                  `This agent may not use the assignment group ${args.assignment_group}`,
                );
              const query = [
                ...statusQuery(args.status),
                ...textQuery(args.text),
                ...(args.priority ? [`priority=${args.priority}`] : []),
                ...(args.assignment_group
                  ? [`assignment_group.name=${safe(args.assignment_group)}`]
                  : []),
                ...(args.caller ? [userIs("caller_id", args.caller)] : []),
                ...(args.assigned_to ? [userIs("assigned_to", args.assigned_to)] : []),
                ...scopeQuery(ctx, INCIDENT),
                "ORDERBYDESCsys_updated_on",
              ];
              retarget(query.join("^").slice(0, 500));
              const { rows, next } = await page(
                ctx,
                mark,
                "incident",
                query,
                INCIDENT_LIST,
                args.max_results,
                args.cursor,
                INCIDENT,
              );
              return json({ incidents: rows.map(incidentSummary), next });
            },
          ),
      );

    if (allowed.has("servicenow_get_incident"))
      server.registerTool(
        "servicenow_get_incident",
        {
          description:
            "Read one incident by number: its fields, the customer-visible comments and the internal work notes " +
            "(each as ServiceNow's activity text, newest first).",
          inputSchema: z.object({
            number: z.string().describe("Incident number, e.g. INC0010001."),
          }),
          annotations: readOnly,
        },
        ({ number }) =>
          audited(ctx, "servicenow_get_incident", number, async (mark) =>
            json(
              incidentDetail(
                await byNumber(ctx, mark, "incident", "incident", number, INCIDENT_FULL, INCIDENT),
              ),
            ),
          ),
      );

    if (allowed.has("servicenow_create_incident"))
      server.registerTool(
        "servicenow_create_incident",
        {
          description:
            "Open a new incident." +
            (listLimit(ctx.limits, "groups")
              ? " assignment_group is required for this agent."
              : "") +
            (listLimit(ctx.limits, "companies") ? " caller is required for this agent." : ""),
          inputSchema: z.object({
            short_description: z.string().min(1).max(160),
            description: z.string().max(10000).optional(),
            caller: who.optional().describe("Who the incident is for: email or user name."),
            assignment_group: z
              .string()
              .max(100)
              .optional()
              .describe(`Assignment group name.${groupNote}`),
            impact: z.number().int().min(1).max(3).optional().describe("1 High, 2 Medium, 3 Low."),
            urgency: z.number().int().min(1).max(3).optional().describe("1 High, 2 Medium, 3 Low."),
            category: z
              .string()
              .max(40)
              .optional()
              .describe("Category value, e.g. software, hardware, network."),
          }),
          annotations: write,
        },
        (args) =>
          audited(
            ctx,
            "servicenow_create_incident",
            args.short_description.slice(0, 200),
            async (mark, retarget) => {
              if (listLimit(ctx.limits, "groups") && !args.assignment_group)
                throw new Denied(
                  "Give an assignment_group: this agent may only open incidents for its groups",
                );
              if (listLimit(ctx.limits, "companies") && !args.caller)
                throw new Denied(
                  "Give a caller: this agent may only open incidents for people of its companies",
                );
              const caller = args.caller ? await person(ctx, mark, args.caller, true) : null;
              const body: Record<string, string> = {
                short_description: args.short_description,
              };
              if (args.description) body.description = args.description;
              if (args.assignment_group)
                body.assignment_group = await group(ctx, mark, args.assignment_group);
              if (caller) {
                body.caller_id = caller.id;
                if (caller.company) body.company = caller.company;
              }
              if (args.impact) body.impact = String(args.impact);
              if (args.urgency) body.urgency = String(args.urgency);
              if (args.category) body.category = args.category;
              const data = await ok(ctx, mark, {
                method: "POST",
                path: tablePath("incident", [], INCIDENT_LIST),
                body,
              });
              const r = data.result as Rec;
              retarget(`${shown(r, "number")}: ${args.short_description}`.slice(0, 300));
              return json(incidentSummary(r));
            },
          ),
      );

    // The incident read back within the key's limits, then changed by its sys_id.
    const patchIncident = async (mark: Mark, number: string, body: Record<string, string>) => {
      const current = await byNumber(
        ctx,
        mark,
        "incident",
        "incident",
        number,
        ["number"],
        INCIDENT,
      );
      const data = await ok(ctx, mark, {
        method: "PATCH",
        path: `${TABLE}/incident/${raw(current, "sys_id")}?${new URLSearchParams({
          sysparm_fields: INCIDENT_LIST.join(","),
          sysparm_display_value: "all",
          sysparm_exclude_reference_link: "true",
        })}`,
        body,
      });
      return incidentSummary(data.result as Rec);
    };

    if (allowed.has("servicenow_update_incident"))
      server.registerTool(
        "servicenow_update_incident",
        {
          description:
            "Change an incident's fields, reassign it or set it to new, in progress or on hold. " +
            "To resolve it use servicenow_resolve_incident; to add a note, servicenow_add_incident_note.",
          inputSchema: z.object({
            number: z.string().describe("Incident number, e.g. INC0010001."),
            short_description: z.string().min(1).max(160).optional(),
            description: z.string().max(10000).optional(),
            impact: z.number().int().min(1).max(3).optional(),
            urgency: z.number().int().min(1).max(3).optional(),
            category: z.string().max(40).optional(),
            assignment_group: z
              .string()
              .max(100)
              .optional()
              .describe(`New assignment group name.${groupNote}`),
            assigned_to: who.optional().describe("New assignee: email or user name."),
            state: z.enum(["new", "in_progress", "on_hold"]).optional(),
            on_hold_reason: z
              .enum(["awaiting_caller", "awaiting_problem", "awaiting_vendor", "awaiting_change"])
              .optional()
              .describe("Required with state on_hold."),
          }),
          annotations: write,
        },
        ({ number, ...args }) =>
          audited(ctx, "servicenow_update_incident", number, async (mark) => {
            if ((args.state === "on_hold") !== Boolean(args.on_hold_reason))
              throw new Denied("Give on_hold_reason exactly when state is on_hold");
            const body: Record<string, string> = {};
            for (const f of ["short_description", "description", "category"] as const)
              if (args[f] !== undefined) body[f] = args[f];
            if (args.impact) body.impact = String(args.impact);
            if (args.urgency) body.urgency = String(args.urgency);
            if (args.state) body.state = INCIDENT_STATE[args.state];
            if (args.on_hold_reason) body.hold_reason = HOLD_REASON[args.on_hold_reason];
            if (args.assignment_group)
              body.assignment_group = await group(ctx, mark, args.assignment_group);
            if (args.assigned_to)
              body.assigned_to = (await person(ctx, mark, args.assigned_to, false)).id;
            if (!Object.keys(body).length) throw new Denied("Give at least one field to change");
            return json(await patchIncident(mark, number, body));
          }),
      );

    if (allowed.has("servicenow_add_incident_note"))
      server.registerTool(
        "servicenow_add_incident_note",
        {
          description:
            "Add a note to an incident. internal (default) is a work note only staff see; customer is a " +
            "comment the caller sees and is usually emailed.",
          inputSchema: z.object({
            number: z.string().describe("Incident number, e.g. INC0010001."),
            text: z.string().min(1).max(10000),
            visibility: z.enum(["internal", "customer"]).default("internal"),
          }),
          annotations: write,
        },
        ({ number, text: note, visibility }) =>
          audited(
            ctx,
            "servicenow_add_incident_note",
            `${number} (${visibility})`,
            async (mark) => {
              const field = visibility === "customer" ? "comments" : "work_notes";
              const incident = await patchIncident(mark, number, {
                [field]: note,
              });
              return json({
                number: incident.number,
                added: visibility === "customer" ? "comment" : "work note",
              });
            },
          ),
      );

    if (allowed.has("servicenow_resolve_incident"))
      server.registerTool(
        "servicenow_resolve_incident",
        {
          description:
            "Resolve an incident with a resolution code and notes. The caller is usually told and can reopen it.",
          inputSchema: z.object({
            number: z.string().describe("Incident number, e.g. INC0010001."),
            close_code: z
              .string()
              .min(1)
              .max(80)
              .describe("Resolution code as your instance names it, e.g. Solution provided."),
            close_notes: z.string().min(1).max(4000),
          }),
          annotations: write,
        },
        ({ number, close_code, close_notes }) =>
          audited(ctx, "servicenow_resolve_incident", number, async (mark) =>
            json(
              await patchIncident(mark, number, {
                state: INCIDENT_STATE.resolved,
                close_code,
                close_notes,
              }),
            ),
          ),
      );

    if (allowed.has("servicenow_search_changes"))
      server.registerTool(
        "servicenow_search_changes",
        {
          description:
            "Search change requests, newest change first. When `next` is set, call again with `cursor` to read more.",
          inputSchema: z.object({
            text: z
              .string()
              .max(200)
              .optional()
              .describe("Words in the short description or description."),
            status,
            assignment_group: z
              .string()
              .max(100)
              .optional()
              .describe(`Assignment group name.${groupNote}`),
            max_results: max(50, 20),
            cursor,
          }),
          annotations: readOnly,
        },
        (args) =>
          audited(
            ctx,
            "servicenow_search_changes",
            JSON.stringify(args).slice(0, 300),
            async (mark, retarget) => {
              if (
                args.assignment_group &&
                !within(listLimit(ctx.limits, "groups"), args.assignment_group)
              )
                throw new Denied(
                  `This agent may not use the assignment group ${args.assignment_group}`,
                );
              const query = [
                ...statusQuery(args.status),
                ...textQuery(args.text),
                ...(args.assignment_group
                  ? [`assignment_group.name=${safe(args.assignment_group)}`]
                  : []),
                ...scopeQuery(ctx, CHANGE),
                "ORDERBYDESCsys_updated_on",
              ];
              retarget(query.join("^").slice(0, 500));
              const { rows, next } = await page(
                ctx,
                mark,
                "change_request",
                query,
                CHANGE_LIST,
                args.max_results,
                args.cursor,
                CHANGE,
              );
              return json({ changes: rows.map(changeSummary), next });
            },
          ),
      );

    if (allowed.has("servicenow_get_change"))
      server.registerTool(
        "servicenow_get_change",
        {
          description: "Read one change request by number, with its plans and approval state.",
          inputSchema: z.object({
            number: z.string().describe("Change number, e.g. CHG0030001."),
          }),
          annotations: readOnly,
        },
        ({ number }) =>
          audited(ctx, "servicenow_get_change", number, async (mark) => {
            const r = await byNumber(
              ctx,
              mark,
              "change_request",
              "change",
              number,
              CHANGE_FULL,
              CHANGE,
            );
            return json({
              ...changeSummary(r),
              description: shown(r, "description"),
              impact: shown(r, "impact"),
              approval: shown(r, "approval"),
              configuration_item: shown(r, "cmdb_ci"),
              justification: shown(r, "justification"),
              implementation_plan: shown(r, "implementation_plan"),
              backout_plan: shown(r, "backout_plan"),
              test_plan: shown(r, "test_plan"),
              close_code: shown(r, "close_code"),
              close_notes: shown(r, "close_notes"),
            });
          }),
      );

    if (allowed.has("servicenow_find_configuration_items"))
      server.registerTool(
        "servicenow_find_configuration_items",
        {
          description:
            "Look up configuration items (computers, servers, applications and so on) in the CMDB by name or serial number.",
          inputSchema: z.object({
            name: z.string().max(100).optional().describe("All or part of the item's name."),
            serial_number: z.string().max(100).optional().describe("The exact serial number."),
            max_results: max(20, 10),
          }),
          annotations: readOnly,
        },
        (args) =>
          audited(
            ctx,
            "servicenow_find_configuration_items",
            JSON.stringify(args).slice(0, 300),
            async (mark) => {
              if (!args.name && !args.serial_number)
                throw new Denied("Give a name or a serial number");
              const query = [
                ...(args.name ? [`nameLIKE${safe(args.name, "The name")}`] : []),
                ...(args.serial_number
                  ? [`serial_number=${safe(args.serial_number, "The serial number")}`]
                  : []),
                ...scopeQuery(ctx, CI),
                "ORDERBYname",
              ];
              const { rows, next } = await page(
                ctx,
                mark,
                "cmdb_ci",
                query,
                CI_FIELDS,
                args.max_results,
                undefined,
                CI,
              );
              return json({
                items: rows.map((r) => ({
                  name: shown(r, "name"),
                  class: shown(r, "sys_class_name"),
                  serial_number: shown(r, "serial_number"),
                  asset_tag: shown(r, "asset_tag"),
                  install_status: shown(r, "install_status"),
                  operational_status: shown(r, "operational_status"),
                  assigned_to: shown(r, "assigned_to"),
                  support_group: shown(r, "support_group"),
                  company: shown(r, "company"),
                  location: shown(r, "location"),
                  ip_address: shown(r, "ip_address"),
                  updated: utc(r, "sys_updated_on"),
                })),
                more: next !== null,
              });
            },
          ),
      );

    if (allowed.has("servicenow_search_knowledge"))
      server.registerTool(
        "servicenow_search_knowledge",
        {
          description:
            "Search published knowledge articles by keywords. Read one with servicenow_get_article. " +
            "When `next` is set, call again with `cursor` to read more.",
          inputSchema: z.object({
            text: z.string().min(1).max(200).describe("Keywords, e.g. vpn will not connect."),
            max_results: max(20, 10),
            cursor,
          }),
          annotations: readOnly,
        },
        ({ text: words, max_results, cursor: at }) =>
          audited(ctx, "servicenow_search_knowledge", words, async (mark) => {
            const query = [
              "workflow_state=published",
              `123TEXTQUERY321=${safe(words, "Search text")}`,
            ];
            const fields = [
              "number",
              "short_description",
              "kb_knowledge_base",
              "kb_category",
              "sys_updated_on",
            ];
            const { rows, next } = await page(
              ctx,
              mark,
              "kb_knowledge",
              query,
              fields,
              max_results,
              at,
              {},
            );
            return json({
              articles: rows.map((r) => ({
                number: shown(r, "number"),
                title: shown(r, "short_description"),
                knowledge_base: shown(r, "kb_knowledge_base"),
                category: shown(r, "kb_category"),
                updated: utc(r, "sys_updated_on"),
              })),
              next,
            });
          }),
      );

    if (allowed.has("servicenow_get_article"))
      server.registerTool(
        "servicenow_get_article",
        {
          description: "Read one published knowledge article by number, as plain text.",
          inputSchema: z.object({
            number: z.string().describe("Article number, e.g. KB0010001."),
          }),
          annotations: readOnly,
        },
        ({ number }) =>
          audited(ctx, "servicenow_get_article", number, async (mark) => {
            if (!NUMBER.test(number)) throw new Denied("Give an article number such as KB0010001");
            const data = await ok(ctx, mark, {
              method: "GET",
              path: tablePath(
                "kb_knowledge",
                [`number=${number}`, "workflow_state=published"],
                [
                  "number",
                  "short_description",
                  "text",
                  "wiki",
                  "kb_knowledge_base",
                  "kb_category",
                  "sys_updated_on",
                ],
                { sysparm_limit: "1" },
              ),
            });
            const r = (data.result ?? [])[0] as Rec | undefined;
            if (!r) throw new Denied(`${number} was not found, or is not published`);
            const body = htmlText(raw(r, "text")) || (raw(r, "wiki") ?? "");
            return json({
              number: shown(r, "number"),
              title: shown(r, "short_description"),
              knowledge_base: shown(r, "kb_knowledge_base"),
              category: shown(r, "kb_category"),
              updated: utc(r, "sys_updated_on"),
              text: clip(body, ARTICLE_CHARS),
            });
          }),
      );

    if (allowed.has("servicenow_search_catalog_items"))
      server.registerTool(
        "servicenow_search_catalog_items",
        {
          description:
            "Search the service catalog for things people can request (laptops, software, access). " +
            "Read one with servicenow_get_catalog_item before ordering it.",
          inputSchema: z.object({
            text: z
              .string()
              .max(200)
              .optional()
              .describe("Words in the item's name or description."),
            max_results: max(50, 20),
            cursor,
          }),
          annotations: readOnly,
        },
        ({ text: words, max_results, cursor: at }) =>
          audited(ctx, "servicenow_search_catalog_items", words ?? null, async (mark) => {
            const start = offset(at);
            const params = new URLSearchParams({
              ...(words ? { sysparm_text: safe(words, "Search text") } : {}),
              sysparm_limit: String(max_results + 1),
              sysparm_offset: String(start),
            });
            const data = await ok(ctx, mark, {
              method: "GET",
              path: `${CATALOG}/items?${params}`,
            });
            const items = (data.result ?? []) as any[];
            return json({
              items: items.slice(0, max_results).map((i) => ({
                id: i.sys_id,
                name: i.name,
                short_description: i.short_description || null,
                category: i.category?.title ?? null,
                price: i.price || null,
              })),
              next: items.length > max_results ? String(start + max_results) : null,
            });
          }),
      );

    if (allowed.has("servicenow_get_catalog_item"))
      server.registerTool(
        "servicenow_get_catalog_item",
        {
          description:
            "Read one catalog item: what it is and the variables (questions) an order must answer.",
          inputSchema: z.object({
            id: z.string().describe("The item's id from servicenow_search_catalog_items."),
          }),
          annotations: readOnly,
        },
        ({ id }) =>
          audited(ctx, "servicenow_get_catalog_item", id, async (mark) => {
            if (!SYS_ID.test(id)) throw new Denied("Give the item's 32 character id");
            const data = await ok(ctx, mark, {
              method: "GET",
              path: `${CATALOG}/items/${id}`,
            });
            const i = data.result ?? {};
            return json({
              id: i.sys_id,
              name: i.name,
              short_description: i.short_description || null,
              description: clip(htmlText(i.description ?? null), 4000),
              price: i.price || null,
              variables: ((i.variables ?? []) as any[])
                .filter((v) => v.name)
                .map((v) => ({
                  name: v.name,
                  label: v.label ?? null,
                  type: v.friendly_type ?? v.type ?? null,
                  mandatory: Boolean(v.mandatory),
                  ...(Array.isArray(v.choices) && v.choices.length
                    ? {
                        choices: v.choices.map((c: any) => ({
                          value: c.value,
                          label: c.label,
                        })),
                      }
                    : {}),
                })),
            });
          }),
      );

    if (allowed.has("servicenow_order_catalog_item"))
      server.registerTool(
        "servicenow_order_catalog_item",
        {
          description:
            "Order a catalog item, answering its variables by name. This starts the item's approvals and fulfillment." +
            (listLimit(ctx.limits, "companies")
              ? " requested_for is required for this agent."
              : ""),
          inputSchema: z.object({
            id: z.string().describe("The item's id."),
            variables: z
              .record(z.string(), z.string())
              .default({})
              .describe("Variable names and answers."),
            quantity: z.number().int().min(1).max(10).default(1),
            requested_for: who
              .optional()
              .describe("Who it is for: email or user name. Default: the integration user."),
          }),
          annotations: write,
        },
        ({ id, variables, quantity, requested_for }) =>
          audited(ctx, "servicenow_order_catalog_item", id, async (mark, retarget) => {
            if (!SYS_ID.test(id)) throw new Denied("Give the item's 32 character id");
            if (Object.keys(variables).length > 100) throw new Denied("Give at most 100 variables");
            if (listLimit(ctx.limits, "companies") && !requested_for)
              throw new Denied(
                "Give requested_for: this agent may only order for people of its companies",
              );
            const forWhom = requested_for ? await person(ctx, mark, requested_for, true) : null;
            const data = await ok(ctx, mark, {
              method: "POST",
              path: `${CATALOG}/items/${id}/order_now`,
              body: {
                sysparm_quantity: String(quantity),
                variables,
                ...(forWhom ? { sysparm_requested_for: forWhom.id } : {}),
              },
            });
            const request = data.result?.request_number ?? data.result?.number ?? null;
            retarget(`${id} -> ${request}`);
            return json({ request_number: request });
          }),
      );

    if (allowed.has("servicenow_search_requests"))
      server.registerTool(
        "servicenow_search_requests",
        {
          description:
            "Track requested items (RITM): their stage and state. Filter by a RITM or REQ number or by who they are for. " +
            "When `next` is set, call again with `cursor` to read more.",
          inputSchema: z.object({
            number: z
              .string()
              .optional()
              .describe("A requested item (RITM...) or request (REQ...) number."),
            requested_for: who.optional().describe("Email or user name."),
            status,
            max_results: max(50, 20),
            cursor,
          }),
          annotations: readOnly,
        },
        (args) =>
          audited(
            ctx,
            "servicenow_search_requests",
            JSON.stringify(args).slice(0, 300),
            async (mark) => {
              if (args.number && !NUMBER.test(args.number))
                throw new Denied("Give a number such as RITM0010001");
              const query = [
                ...statusQuery(args.status),
                ...(args.number
                  ? [
                      args.number.startsWith("REQ")
                        ? `request.number=${args.number}`
                        : `number=${args.number}`,
                    ]
                  : []),
                ...(args.requested_for ? [userIs("requested_for", args.requested_for)] : []),
                ...scopeQuery(ctx, RITM),
                "ORDERBYDESCsys_updated_on",
              ];
              const { rows, next } = await page(
                ctx,
                mark,
                "sc_req_item",
                query,
                RITM_FIELDS,
                args.max_results,
                args.cursor,
                RITM,
              );
              return json({
                requests: rows.map((r) => ({
                  number: shown(r, "number"),
                  request: shown(r, "request"),
                  item: shown(r, "cat_item"),
                  stage: shown(r, "stage"),
                  state: shown(r, "state"),
                  requested_for: shown(r, "requested_for"),
                  assignment_group: shown(r, "assignment_group"),
                  opened: utc(r, "opened_at"),
                  due: utc(r, "due_date"),
                  updated: utc(r, "sys_updated_on"),
                })),
                next,
              });
            },
          ),
      );
  },
};
