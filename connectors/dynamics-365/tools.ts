import { z } from "zod";
import { API, dataverseError } from "./vendor.js";
import {
  audited,
  Denied,
  json,
  listLimit,
  VendorFailed,
  type ToolContext,
  type Toolset,
} from "../../packages/connector-kit/src/common.js";
import type { VendorRequest } from "../../packages/connector-kit/src/types.js";

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const COLUMN = /^_?[a-z][a-z0-9_]{0,99}$/;
const TABLE = /^[a-z][a-z0-9_]{0,63}$/;
const ISO_DATE = /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}(:\d{2}(\.\d{1,7})?)?(Z|[+-]\d{2}:\d{2}))?$/;
const FORMATTED = "@OData.Community.Display.V1.FormattedValue";
const LOGICAL = "@Microsoft.Dynamics.CRM.lookuplogicalname";
// Formatted values (option labels, money, lookup names) and lookups' table names.
const PREFER = 'odata.include-annotations="*"';
const SCOPE_COLUMNS = ["_owningbusinessunit_value", "_owningteam_value"];
const MAX_TEXT = 2000;

type Row = Record<string, unknown>;

// The tables the built-in tools use: their entity set, id column, the columns a search answers
// and the ones a single read adds, and what a text search looks in.
const TABLES = {
  account: {
    set: "accounts",
    id: "accountid",
    text: ["name", "accountnumber"],
    list: [
      "name",
      "accountnumber",
      "telephone1",
      "emailaddress1",
      "websiteurl",
      "address1_city",
      "address1_country",
      "_primarycontactid_value",
      "_parentaccountid_value",
      "_ownerid_value",
      "statecode",
      "modifiedon",
    ],
    detail: ["industrycode", "numberofemployees", "revenue", "description", "createdon"],
  },
  contact: {
    set: "contacts",
    id: "contactid",
    text: ["fullname", "emailaddress1"],
    list: [
      "fullname",
      "jobtitle",
      "emailaddress1",
      "telephone1",
      "mobilephone",
      "_parentcustomerid_value",
      "address1_city",
      "_ownerid_value",
      "statecode",
      "modifiedon",
    ],
    detail: ["firstname", "lastname", "address1_country", "description", "createdon"],
  },
  opportunity: {
    set: "opportunities",
    id: "opportunityid",
    text: ["name"],
    list: [
      "name",
      "_customerid_value",
      "estimatedvalue",
      "estimatedclosedate",
      "closeprobability",
      "stepname",
      "_ownerid_value",
      "statecode",
      "statuscode",
      "modifiedon",
    ],
    detail: ["actualvalue", "actualclosedate", "description", "createdon"],
  },
  incident: {
    set: "incidents",
    id: "incidentid",
    text: ["title", "ticketnumber"],
    list: [
      "ticketnumber",
      "title",
      "_customerid_value",
      "prioritycode",
      "casetypecode",
      "_ownerid_value",
      "statecode",
      "statuscode",
      "createdon",
      "modifiedon",
    ],
    detail: ["description", "_primarycontactid_value", "caseorigincode"],
  },
} as const;
type TableName = keyof typeof TABLES;
const PARENTS = Object.keys(TABLES) as [TableName, ...TableName[]];

type Scope = { units: string[]; teams: string[] } | null;

function scopeOf(ctx: ToolContext): Scope {
  const units = listLimit(ctx.limits, "business_units") ?? [];
  const teams = listLimit(ctx.limits, "teams") ?? [];
  return units.length || teams.length
    ? {
        units: units.map((u) => u.toLowerCase()),
        teams: teams.map((t) => t.toLowerCase()),
      }
    : null;
}

export function scopeFilter(scope: NonNullable<Scope>): string {
  return `(${[
    ...scope.units.map((u) => `_owningbusinessunit_value eq ${u}`),
    ...scope.teams.map((t) => `_owningteam_value eq ${t}`),
  ].join(" or ")})`;
}

export function inScope(scope: Scope, row: Row): boolean {
  if (!scope) return true;
  const unit = String(row._owningbusinessunit_value ?? "").toLowerCase();
  const team = String(row._owningteam_value ?? "").toLowerCase();
  return scope.units.includes(unit) || (team !== "" && scope.teams.includes(team));
}

const OUTSIDE = "This agent may not use that record";

/** One answer value as an OData literal: GUIDs, dates, numbers and booleans as they are, any
 *  other text quoted. */
export function literal(value: string | number | boolean): string {
  if (typeof value === "boolean") return String(value);
  if (typeof value === "number") {
    if (!Number.isFinite(value)) throw new Denied("Numbers must be finite");
    return String(value);
  }
  if (GUID.test(value) || ISO_DATE.test(value)) return value;
  return `'${value.replaceAll("'", "''")}'`;
}

const filterSchema = z.object({
  column: z
    .string()
    .describe(
      "Column logical name, e.g. statecode or address1_city. A lookup is _<name>_value, e.g. _parentcustomerid_value (compare it to a record id).",
    ),
  op: z.enum(["eq", "ne", "gt", "ge", "lt", "le", "contains", "startswith", "null", "not_null"]),
  value: z
    .union([z.string().max(400), z.number(), z.boolean()])
    .optional()
    .describe(
      "Not used with null and not_null. Ids, dates (2026-10-27 or 2026-10-27T09:00:00Z) and numbers are compared as such. Option set columns such as statecode compare their number, e.g. 0 for active.",
    ),
});
type Filter = z.infer<typeof filterSchema>;
const filtersArg = z
  .array(filterSchema)
  .max(10)
  .default([])
  .describe("Conditions every record must meet. No free-form queries.");

function checkColumn(column: string): string {
  if (!COLUMN.test(column)) throw new Denied(`${column} is not a column logical name`);
  return column;
}

export function filterExpression(filters: Filter[]): string[] {
  return filters.map(({ column, op, value }) => {
    checkColumn(column);
    if (op === "null" || op === "not_null") return `${column} ${op === "null" ? "eq" : "ne"} null`;
    if (value === undefined) throw new Denied(`Give a value to compare ${column} with`);
    if (op === "contains" || op === "startswith") {
      if (typeof value !== "string") throw new Denied(`${op} compares text`);
      return `${op}(${column},${literal(value)})`;
    }
    return `${column} ${op} ${literal(value)}`;
  });
}

async function ok(ctx: ToolContext, mark: (s: number) => void, req: VendorRequest) {
  const res = await ctx.call({
    ...req,
    headers: { Prefer: PREFER, ...req.headers },
  });
  mark(res.status);
  if (res.status < 200 || res.status >= 300)
    throw new VendorFailed(res.status, dataverseError(res.status, res.data));
  return res.data as Row;
}

/** A row as agents get it: lookups as {id, name, table}, option sets and money as their
 *  labels, text and dates as they are (long text shortened). */
export function trimRow(row: Row, idColumn: string): Row {
  const out: Row = { id: row[idColumn] ?? null };
  for (const [key, value] of Object.entries(row)) {
    if (key.includes("@") || key === idColumn || SCOPE_COLUMNS.includes(key)) continue;
    const label = row[`${key}${FORMATTED}`];
    const lookup = /^_(.+)_value$/.exec(key);
    if (lookup)
      out[lookup[1]] =
        value === null
          ? null
          : {
              id: value,
              name: label ?? null,
              table: row[`${key}${LOGICAL}`] ?? null,
            };
    else if (typeof value === "string")
      out[key] = value.length > MAX_TEXT ? `${value.slice(0, MAX_TEXT)}…` : value;
    else out[key] = value === null ? null : (label ?? value);
  }
  return out;
}

const select = (columns: readonly string[], scope: Scope) =>
  [...new Set([...columns, ...(scope ? SCOPE_COLUMNS : [])])].join(",");

function idArg(id: string): string {
  if (!GUID.test(id)) throw new Denied("Give a record id (a GUID)");
  return id.toLowerCase();
}

type Table = { set: string; id: string };

// One record, read only when the key may use it. Under a scope, a record not found gets the
// same refusal as one outside it, so a key cannot learn which ids exist.
async function readRecord(
  ctx: ToolContext,
  mark: (s: number) => void,
  table: Table,
  id: string,
  columns: readonly string[],
): Promise<Row> {
  const scope = scopeOf(ctx);
  const res = await ctx.call({
    method: "GET",
    path: `${API}/${table.set}(${idArg(id)})?$select=${select(columns, scope)}`,
    headers: { Prefer: PREFER },
  });
  mark(res.status);
  if (scope && res.status === 404) throw new Denied(OUTSIDE);
  if (res.status !== 200) throw new VendorFailed(res.status, dataverseError(res.status, res.data));
  if (!inScope(scope, res.data as Row)) throw new Denied(OUTSIDE);
  return res.data as Row;
}

const nextOf = (data: Row): string | null => {
  const link = data["@odata.nextLink"];
  return typeof link === "string" ? new URL(link).searchParams.get("$skiptoken") : null;
};

// A page of rows. Dataverse pages with `odata.maxpagesize` and answers a skip token, which comes
// back as `cursor` with the same arguments. `unscoped`: rows of a record already checked against
// the key's scope (its activities and notes are the record's, whoever owns them).
async function listRows(
  ctx: ToolContext,
  mark: (s: number) => void,
  input: {
    table: Table;
    columns: readonly string[];
    conditions: string[];
    orderBy: string;
    top: number;
    cursor?: string;
    unscoped?: true;
  },
) {
  const scope = input.unscoped ? null : scopeOf(ctx);
  const conditions = [...input.conditions, ...(scope ? [scopeFilter(scope)] : [])];
  if (input.cursor !== undefined && (!input.cursor || input.cursor.length > 4000))
    throw new Denied("Pass the cursor exactly as it was answered");
  const query = [
    `$select=${select(input.columns, scope)}`,
    ...(conditions.length ? [`$filter=${encodeURIComponent(conditions.join(" and "))}`] : []),
    `$orderby=${encodeURIComponent(input.orderBy)}`,
    ...(input.cursor ? [`$skiptoken=${encodeURIComponent(input.cursor)}`] : []),
  ].join("&");
  const data = await ok(ctx, mark, {
    method: "GET",
    path: `${API}/${input.table.set}?${query}`,
    headers: { Prefer: `${PREFER},odata.maxpagesize=${input.top}` },
  });
  const rows = ((data.value as Row[] | undefined) ?? []).filter((r) => inScope(scope, r));
  const next = nextOf(data);
  return {
    records: rows.map((r) => trimRow(r, input.table.id)),
    more: Boolean(next),
    next,
  };
}

const textSearch = (columns: readonly string[], text: string) =>
  `(${columns.map((c) => `contains(${c},${literal(text)})`).join(" or ")})`;

const topArg = z.number().int().min(1).max(50).default(20);
const cursorArg = z
  .string()
  .optional()
  .describe("The `next` of the previous answer, with the same other arguments.");
const fieldValues = z
  .record(z.string(), z.union([z.string().max(4000), z.number(), z.boolean(), z.null()]))
  .describe(
    'Column logical names and their values, e.g. {"telephone1": "+1 555 0100"}. Option sets take their number. Lookups cannot be set here.',
  );

function checkFields(fields: Record<string, unknown>): void {
  const names = Object.keys(fields);
  if (names.length === 0 || names.length > 50) throw new Denied("Give between 1 and 50 columns");
  // Plain columns only: a lookup (`…@odata.bind`) could reassign the record's owner or parent.
  names.forEach(checkColumn);
  const lookup = names.find((n) => n.startsWith("_"));
  if (lookup) throw new Denied(`${lookup} is a lookup, which cannot be set here`);
}

// Records an agent creates are owned by the application user (its own business unit) unless a
// team is named. Under a scope, the new record must land inside it.
async function owner(
  ctx: ToolContext,
  mark: (s: number) => void,
  team: string | undefined,
): Promise<Row> {
  const scope = scopeOf(ctx);
  if (team) {
    const id = idArg(team);
    if (scope && !scope.teams.includes(id)) {
      const found = await ctx.call({
        method: "GET",
        path: `${API}/teams(${id})?$select=_businessunitid_value`,
      });
      mark(found.status);
      const unit = String((found.data as Row | null)?._businessunitid_value ?? "").toLowerCase();
      if (found.status !== 200 || !scope.units.includes(unit))
        throw new Denied("This agent may not give records to that team");
    }
    return { "ownerid@odata.bind": `/teams(${id})` };
  }
  if (!scope) return {};
  if (scope.units.length) {
    const who = await ok(ctx, mark, { method: "GET", path: `${API}/WhoAmI` });
    if (scope.units.includes(String(who.BusinessUnitId).toLowerCase())) return {};
  }
  if (scope.teams.length === 1) return { "ownerid@odata.bind": `/teams(${scope.teams[0]})` };
  throw new Denied(
    "A record created here would be outside this agent's business units and teams. Name one of its teams as owner_team.",
  );
}

const annotations = {
  readOnlyHint: false,
  destructiveHint: false,
  openWorldHint: true,
} as const;
const readOnly = { readOnlyHint: true, openWorldHint: true } as const;

const regarding = (kind: "task" | "phonecall", table: TableName, id: string) => ({
  [`regardingobjectid_${table}_${kind}@odata.bind`]: `/${TABLES[table].set}(${id})`,
});

const ACTIVITY_COLUMNS = [
  "activityid",
  "activitytypecode",
  "subject",
  "description",
  "scheduledend",
  "actualend",
  "statecode",
  "_ownerid_value",
  "createdon",
];
const NOTE_COLUMNS = [
  "annotationid",
  "subject",
  "notetext",
  "filename",
  "_createdby_value",
  "createdon",
];

// Tables an agent may delete from. Deleting a customer, opportunity or case also deletes the
// records Dynamics cascades to, which may be outside a scoped key, so a scoped key deletes only
// activities and notes, checked through the record they belong to.
const DELETABLE = {
  ...Object.fromEntries(PARENTS.map((t) => [t, { set: TABLES[t].set, parent: null }])),
  task: { set: "tasks", parent: "_regardingobjectid_value" },
  phonecall: { set: "phonecalls", parent: "_regardingobjectid_value" },
  annotation: { set: "annotations", parent: "_objectid_value" },
} as Record<string, { set: string; parent: string | null }>;

type EntityDefinition = {
  EntitySetName?: string;
  PrimaryIdAttribute?: string;
  PrimaryNameAttribute?: string | null;
  OwnershipType?: string;
};

async function listedTable(ctx: ToolContext, mark: (s: number) => void, table: string) {
  const tables = listLimit(ctx.limits, "tables");
  if (!tables)
    throw new Denied("This agent has no tables to read. An administrator lists them on its key.");
  if (!TABLE.test(table) || !tables.includes(table))
    throw new Denied(`This agent may not read ${table}. It may read: ${tables.join(", ")}`);
  const def = (await ok(ctx, mark, {
    method: "GET",
    path: `${API}/EntityDefinitions(LogicalName='${table}')?$select=EntitySetName,PrimaryIdAttribute,PrimaryNameAttribute,OwnershipType`,
  })) as EntityDefinition;
  if (!def.EntitySetName || !def.PrimaryIdAttribute)
    throw new Error(`Dynamics 365 metadata for ${table} has no entity set`);
  if (scopeOf(ctx) && def.OwnershipType !== "UserOwned")
    throw new Denied(
      `${table} records have no owner, so an agent kept to business units or teams cannot read them`,
    );
  return {
    set: def.EntitySetName,
    id: def.PrimaryIdAttribute,
    defaults: [
      def.PrimaryIdAttribute,
      ...(def.PrimaryNameAttribute ? [def.PrimaryNameAttribute] : []),
    ],
  };
}

const columnsArg = z
  .array(z.string())
  .max(30)
  .optional()
  .describe("Column logical names to return. The id and primary name if omitted.");

const tableArg = (tables: readonly TableName[]) =>
  z
    .enum(tables as [TableName, ...TableName[]])
    .default(tables[0])
    .describe(
      tables.length > 1 ? "account (a company) or contact (a person)." : `Always ${tables[0]}.`,
    );

export const dynamicsTools: Toolset = {
  names: [
    "dynamics_search_customers",
    "dynamics_get_customer",
    "dynamics_create_customer",
    "dynamics_update_customer",
    "dynamics_search_opportunities",
    "dynamics_get_opportunity",
    "dynamics_close_opportunity",
    "dynamics_search_cases",
    "dynamics_get_case",
    "dynamics_close_case",
    "dynamics_list_activities",
    "dynamics_add_activity",
    "dynamics_list_rows",
    "dynamics_get_row",
    "dynamics_delete_record",
  ],

  register(server, ctx, allowed) {
    const search = (name: string, tables: readonly TableName[], what: string) => {
      if (!allowed.has(name)) return;
      server.registerTool(
        name,
        {
          description: `Find ${what}, newest change first, by text and column conditions. When \`more\` is true, call again with \`cursor\` set to \`next\`.`,
          inputSchema: z.object({
            table: tableArg(tables),
            text: z
              .string()
              .max(200)
              .optional()
              .describe(`Part of ${tables.map((t) => TABLES[t].text.join(" or ")).join("; or ")}.`),
            filters: filtersArg,
            top: topArg,
            cursor: cursorArg,
          }),
          annotations: readOnly,
        },
        (args) =>
          audited(ctx, name, JSON.stringify(args).slice(0, 300), async (mark) => {
            const def = TABLES[args.table];
            const conditions = [
              ...(args.text ? [textSearch(def.text, args.text)] : []),
              ...filterExpression(args.filters),
            ];
            return json(
              await listRows(ctx, mark, {
                table: def,
                columns: [def.id, ...def.list],
                conditions,
                orderBy: "modifiedon desc",
                top: args.top,
                cursor: args.cursor,
              }),
            );
          }),
      );
    };

    const get = (name: string, tables: readonly TableName[], what: string) => {
      if (!allowed.has(name)) return;
      server.registerTool(
        name,
        {
          description: `Read one ${what} by id, with its description and status.`,
          inputSchema: z.object({
            table: tableArg(tables),
            id: z.string().describe("The record id (a GUID)."),
          }),
          annotations: readOnly,
        },
        (args) =>
          audited(ctx, name, `${args.table}/${args.id}`, async (mark) => {
            const def = TABLES[args.table];
            const row = await readRecord(ctx, mark, def, args.id, [
              def.id,
              ...def.list,
              ...def.detail,
            ]);
            return json(trimRow(row, def.id));
          }),
      );
    };

    search("dynamics_search_customers", ["account", "contact"], "accounts or contacts");
    get("dynamics_get_customer", ["account", "contact"], "account or contact");
    search("dynamics_search_opportunities", ["opportunity"], "opportunities");
    get("dynamics_get_opportunity", ["opportunity"], "opportunity");
    search("dynamics_search_cases", ["incident"], "cases");
    get("dynamics_get_case", ["incident"], "case");

    const customerArgs = {
      fields: fieldValues,
      account_id: z
        .string()
        .optional()
        .describe("Contacts only: the account the contact works for."),
    };
    async function accountBind(
      mark: (s: number) => void,
      table: "account" | "contact",
      accountId: string | undefined,
    ): Promise<Row> {
      if (accountId === undefined) return {};
      if (table !== "contact") throw new Denied("account_id is for contacts");
      await readRecord(ctx, mark, TABLES.account, accountId, [TABLES.account.id]);
      return {
        "parentcustomerid_account@odata.bind": `/accounts(${idArg(accountId)})`,
      };
    }

    if (allowed.has("dynamics_create_customer"))
      server.registerTool(
        "dynamics_create_customer",
        {
          description:
            "Create an account (a company) or a contact (a person). Accounts need `name`; contacts need `lastname`. Returns the new record.",
          inputSchema: z.object({
            table: z.enum(["account", "contact"]),
            ...customerArgs,
            owner_team: z
              .string()
              .optional()
              .describe(
                "A team id to own the record. Needed when this agent is kept to several teams.",
              ),
          }),
          annotations,
        },
        ({ table, fields, account_id, owner_team }) =>
          audited(ctx, "dynamics_create_customer", table, async (mark, retarget) => {
            checkFields(fields);
            const def = TABLES[table];
            const body = {
              ...fields,
              ...(await accountBind(mark, table, account_id)),
              ...(await owner(ctx, mark, owner_team)),
            };
            const created = await ok(ctx, mark, {
              method: "POST",
              path: `${API}/${def.set}?$select=${[def.id, ...def.list].join(",")}`,
              body,
              headers: { Prefer: `${PREFER},return=representation` },
            });
            retarget(`${table}/${created[def.id]}`);
            return json(trimRow(created, def.id));
          }),
      );

    if (allowed.has("dynamics_update_customer"))
      server.registerTool(
        "dynamics_update_customer",
        {
          description:
            "Change columns of an existing account or contact, e.g. a phone number or address. Never creates a record.",
          inputSchema: z.object({
            table: z.enum(["account", "contact"]),
            id: z.string().describe("The record id (a GUID)."),
            ...customerArgs,
          }),
          annotations,
        },
        ({ table, id, fields, account_id }) =>
          audited(ctx, "dynamics_update_customer", `${table}/${id}`, async (mark) => {
            checkFields(fields);
            const def = TABLES[table];
            await readRecord(ctx, mark, def, id, [def.id]);
            const body = {
              ...fields,
              ...(await accountBind(mark, table, account_id)),
            };
            // If-Match: an update to an id that does not exist would otherwise create it.
            await ok(ctx, mark, {
              method: "PATCH",
              path: `${API}/${def.set}(${idArg(id)})`,
              body,
              headers: { "If-Match": "*" },
            });
            return json({
              updated: true,
              id: idArg(id),
              columns: Object.keys(fields),
            });
          }),
      );

    if (allowed.has("dynamics_close_opportunity"))
      server.registerTool(
        "dynamics_close_opportunity",
        {
          description:
            "Close an open opportunity as won or lost. Dynamics records the close with its subject and, for a win, the actual revenue.",
          inputSchema: z.object({
            id: z.string().describe("The opportunity id (a GUID)."),
            outcome: z.enum(["won", "lost"]),
            subject: z.string().min(1).max(200).describe("Why it closed, in a few words."),
            description: z.string().max(2000).optional(),
            actual_revenue: z.number().nonnegative().optional().describe("Won only."),
          }),
          annotations,
        },
        ({ id, outcome, subject, description, actual_revenue }) =>
          audited(
            ctx,
            "dynamics_close_opportunity",
            `opportunity/${id} ${outcome}`,
            async (mark) => {
              if (outcome === "lost" && actual_revenue !== undefined)
                throw new Denied("actual_revenue is for a won opportunity");
              await readRecord(ctx, mark, TABLES.opportunity, id, [TABLES.opportunity.id]);
              // Won is status 3, lost 4: Dynamics' default status reasons for those states.
              await ok(ctx, mark, {
                method: "POST",
                path: `${API}/${outcome === "won" ? "WinOpportunity" : "LoseOpportunity"}`,
                body: {
                  Status: outcome === "won" ? 3 : 4,
                  OpportunityClose: {
                    subject,
                    ...(description ? { description } : {}),
                    ...(actual_revenue !== undefined ? { actualrevenue: actual_revenue } : {}),
                    "opportunityid@odata.bind": `/opportunities(${idArg(id)})`,
                  },
                },
              });
              return json({ closed: true, id: idArg(id), outcome });
            },
          ),
      );

    if (allowed.has("dynamics_close_case"))
      server.registerTool(
        "dynamics_close_case",
        {
          description:
            "Resolve a case with a resolution, or cancel it. Resolving records the resolution on the case.",
          inputSchema: z.object({
            id: z.string().describe("The case id (a GUID)."),
            outcome: z.enum(["resolved", "cancelled"]),
            resolution: z
              .string()
              .min(1)
              .max(200)
              .optional()
              .describe("Resolved only: the resolution."),
            description: z.string().max(2000).optional(),
          }),
          annotations,
        },
        ({ id, outcome, resolution, description }) =>
          audited(ctx, "dynamics_close_case", `incident/${id} ${outcome}`, async (mark) => {
            if (outcome === "resolved" && !resolution)
              throw new Denied("Give the resolution of a resolved case");
            await readRecord(ctx, mark, TABLES.incident, id, [TABLES.incident.id]);
            const case_ = idArg(id);
            if (outcome === "resolved")
              // Status 5 is Problem Solved, the default resolved reason.
              await ok(ctx, mark, {
                method: "POST",
                path: `${API}/CloseIncident`,
                body: {
                  Status: 5,
                  IncidentResolution: {
                    subject: resolution,
                    ...(description ? { description } : {}),
                    "incidentid@odata.bind": `/incidents(${case_})`,
                  },
                },
              });
            else
              await ok(ctx, mark, {
                method: "PATCH",
                path: `${API}/incidents(${case_})`,
                body: { statecode: 2, statuscode: 6 },
                headers: { "If-Match": "*" },
              });
            return json({ closed: true, id: case_, outcome });
          }),
      );

    const parentArgs = {
      table: z
        .enum(PARENTS)
        .describe("The table of the record: account, contact, opportunity or incident (a case)."),
      id: z.string().describe("The record id (a GUID)."),
    };

    if (allowed.has("dynamics_list_activities"))
      server.registerTool(
        "dynamics_list_activities",
        {
          description:
            "List the activities (tasks, phone calls, emails, appointments) or the notes of an account, contact, opportunity or case, newest first. When `more` is true, call again with `cursor` set to `next`.",
          inputSchema: z.object({
            ...parentArgs,
            kind: z.enum(["activities", "notes"]).default("activities"),
            top: topArg,
            cursor: cursorArg,
          }),
          annotations: readOnly,
        },
        ({ table, id, kind, top, cursor }) =>
          audited(ctx, "dynamics_list_activities", `${table}/${id} ${kind}`, async (mark) => {
            await readRecord(ctx, mark, TABLES[table], id, [TABLES[table].id]);
            const notes = kind === "notes";
            const page = await listRows(ctx, mark, {
              table: notes
                ? { set: "annotations", id: "annotationid" }
                : { set: "activitypointers", id: "activityid" },
              columns: notes ? NOTE_COLUMNS : ACTIVITY_COLUMNS,
              conditions: [
                `${notes ? "_objectid_value" : "_regardingobjectid_value"} eq ${idArg(id)}`,
              ],
              orderBy: "createdon desc",
              top,
              cursor,
              unscoped: true,
            });
            return json(page);
          }),
      );

    if (allowed.has("dynamics_add_activity"))
      server.registerTool(
        "dynamics_add_activity",
        {
          description:
            "Add a note, a task or a logged phone call to an account, contact, opportunity or case.",
          inputSchema: z.object({
            ...parentArgs,
            kind: z.enum(["note", "task", "phone_call"]),
            subject: z.string().min(1).max(200),
            text: z
              .string()
              .max(10000)
              .optional()
              .describe("The note, or the activity's description."),
            due: z
              .string()
              .optional()
              .describe("Tasks only: when it is due, e.g. 2026-11-03T17:00:00Z."),
            phone_number: z.string().max(50).optional().describe("Phone calls only."),
            direction: z
              .enum(["outgoing", "incoming"])
              .default("outgoing")
              .describe("Phone calls only."),
            completed: z
              .boolean()
              .default(true)
              .describe("Phone calls only: false keeps it open as a call still to make."),
          }),
          annotations,
        },
        (args) =>
          audited(
            ctx,
            "dynamics_add_activity",
            `${args.table}/${args.id} ${args.kind}`,
            async (mark, retarget) => {
              const { table, id, kind, subject, text } = args;
              if (args.due !== undefined && (kind !== "task" || !ISO_DATE.test(args.due)))
                throw new Denied("due is a date or time such as 2026-11-03T17:00:00Z, for tasks");
              if (args.phone_number !== undefined && kind !== "phone_call")
                throw new Denied("phone_number is for phone calls");
              await readRecord(ctx, mark, TABLES[table], id, [TABLES[table].id]);
              const record = idArg(id);
              const prefer = { Prefer: `${PREFER},return=representation` };
              let created: string;
              if (kind === "note") {
                const note = await ok(ctx, mark, {
                  method: "POST",
                  path: `${API}/annotations?$select=annotationid`,
                  body: {
                    subject,
                    ...(text ? { notetext: text } : {}),
                    [`objectid_${table}@odata.bind`]: `/${TABLES[table].set}(${record})`,
                  },
                  headers: prefer,
                });
                created = String(note.annotationid);
              } else {
                const set = kind === "task" ? "tasks" : "phonecalls";
                const activity = await ok(ctx, mark, {
                  method: "POST",
                  path: `${API}/${set}?$select=activityid`,
                  body: {
                    subject,
                    ...(text ? { description: text } : {}),
                    ...(kind === "task"
                      ? { ...(args.due ? { scheduledend: args.due } : {}) }
                      : {
                          directioncode: args.direction === "outgoing",
                          ...(args.phone_number ? { phonenumber: args.phone_number } : {}),
                        }),
                    ...regarding(kind === "task" ? "task" : "phonecall", table, record),
                  },
                  headers: prefer,
                });
                created = String(activity.activityid);
                // A logged call is a call that happened: Completed (state 1, status 2).
                if (kind === "phone_call" && args.completed)
                  await ok(ctx, mark, {
                    method: "PATCH",
                    path: `${API}/phonecalls(${created})`,
                    body: { statecode: 1, statuscode: 2 },
                    headers: { "If-Match": "*" },
                  });
              }
              retarget(`${table}/${record} ${kind} ${created}`);
              return json({ id: created, kind });
            },
          ),
      );

    if (allowed.has("dynamics_list_rows"))
      server.registerTool(
        "dynamics_list_rows",
        {
          description:
            "List rows of a table this agent may read (its key lists them), with column conditions. When `more` is true, call again with `cursor` set to `next`.",
          inputSchema: z.object({
            table: z.string().describe("Table logical name, e.g. product or new_project."),
            columns: columnsArg,
            filters: filtersArg,
            order_by: z
              .string()
              .optional()
              .describe("A column to sort by. The table's own order if omitted."),
            descending: z.boolean().default(false),
            top: topArg,
            cursor: cursorArg,
          }),
          annotations: readOnly,
        },
        (args) =>
          audited(
            ctx,
            "dynamics_list_rows",
            `${args.table}: ${JSON.stringify(args).slice(0, 250)}`,
            async (mark) => {
              const columns = (args.columns ?? []).map(checkColumn);
              const conditions = filterExpression(args.filters);
              const order = args.order_by ? checkColumn(args.order_by) : null;
              const table = await listedTable(ctx, mark, args.table);
              return json(
                await listRows(ctx, mark, {
                  table,
                  columns: columns.length ? [table.id, ...columns] : table.defaults,
                  conditions,
                  orderBy: `${order ?? table.id} ${args.descending ? "desc" : "asc"}`,
                  top: args.top,
                  cursor: args.cursor,
                }),
              );
            },
          ),
      );

    if (allowed.has("dynamics_get_row"))
      server.registerTool(
        "dynamics_get_row",
        {
          description: "Read one row of a table this agent may read (its key lists them) by id.",
          inputSchema: z.object({
            table: z.string().describe("Table logical name, e.g. product or new_project."),
            id: z.string().describe("The row id (a GUID)."),
            columns: columnsArg,
          }),
          annotations: readOnly,
        },
        ({ table, id, columns }) =>
          audited(ctx, "dynamics_get_row", `${table}/${id}`, async (mark) => {
            const picked = (columns ?? []).map(checkColumn);
            idArg(id);
            const def = await listedTable(ctx, mark, table);
            const row = await readRecord(
              ctx,
              mark,
              def,
              id,
              picked.length ? [def.id, ...picked] : def.defaults,
            );
            return json(trimRow(row, def.id));
          }),
      );

    if (allowed.has("dynamics_delete_record"))
      server.registerTool(
        "dynamics_delete_record",
        {
          description:
            "Delete an account, contact, opportunity, case (incident), task, phone call or note (annotation). Dynamics also deletes the records that belong to it, such as its activities and notes. It cannot be undone.",
          inputSchema: z.object({
            table: z.enum(Object.keys(DELETABLE) as [string, ...string[]]),
            id: z.string().describe("The record id (a GUID)."),
          }),
          annotations: {
            readOnlyHint: false,
            destructiveHint: true,
            openWorldHint: true,
          },
        },
        ({ table, id }) =>
          audited(ctx, "dynamics_delete_record", `${table}/${id}`, async (mark) => {
            const target = DELETABLE[table];
            const record = idArg(id);
            if (scopeOf(ctx)) {
              if (!target.parent)
                throw new Denied(
                  `An agent kept to business units or teams cannot delete ${table} records: Dynamics would also delete records that belong to it, which may be outside them`,
                );
              const found = await ctx.call({
                method: "GET",
                path: `${API}/${target.set}(${record})?$select=${target.parent}`,
                headers: { Prefer: PREFER },
              });
              mark(found.status);
              if (found.status === 404) throw new Denied(OUTSIDE);
              if (found.status !== 200)
                throw new VendorFailed(found.status, dataverseError(found.status, found.data));
              const row = found.data as Row;
              const parentTable = row[`${target.parent}${LOGICAL}`] as string | undefined;
              const parentId = row[target.parent];
              if (!parentTable || !(parentTable in TABLES) || typeof parentId !== "string")
                throw new Denied(OUTSIDE);
              const def = TABLES[parentTable as TableName];
              await readRecord(ctx, mark, def, parentId, [def.id]);
            }
            await ok(ctx, mark, {
              method: "DELETE",
              path: `${API}/${target.set}(${record})`,
            });
            return json({ deleted: true, table, id: record });
          }),
      );
  },
};
