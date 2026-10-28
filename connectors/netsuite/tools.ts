import { z } from "zod";
import { netsuiteError, RESTLET_PATH, SUITEQL_PATH } from "./vendor.js";
import {
  audited,
  Denied,
  json,
  listLimit,
  VendorFailed,
  type ToolContext,
  type Toolset,
} from "../../packages/connector-kit/src/common.js";

// Agents never write SuiteQL: every query is built here from validated parameters. SuiteQL
// over REST takes no bound parameters, so ids and dates are checked against strict patterns
// and text is quoted with single quotes doubled (SuiteQL, like Oracle, has no backslash
// escapes). The key's subsidiaries and locations go into every query's WHERE and are checked
// again on each row returned.

const RECORD = "/services/rest/record/v1";
const SAVED_SEARCH_SCRIPT =
  "script=customscript_agent_saved_search&deploy=customdeploy_agent_saved_search";
const ID = /^[0-9]{1,12}$/;
const DATE = /^\d{4}-\d{2}-\d{2}$/;
const SEARCH_ID = /^(customsearch[a-z0-9_]{1,80}|[0-9]{1,12})$/;

type Row = Record<string, unknown>;
type Mark = (status: number) => void;
type Request = Parameters<ToolContext["call"]>[0];

async function ok(ctx: ToolContext, mark: Mark, req: Request, hintTool?: string) {
  const res = await ctx.call(req);
  mark(res.status);
  if (res.status >= 200 && res.status < 300) return res.data as any;
  const why = netsuiteError(res.data);
  // NetSuite answers a record type the role may not read as not found, so the hint of what
  // the role needs is given as for a refusal.
  const status = res.status === 400 && /Record '[^']+' was not found/i.test(why) ? 403 : res.status;
  throw new VendorFailed(
    status,
    `NetSuite answered ${res.status}${why ? `: ${why}` : ""}`,
    hintTool,
  );
}

const quote = (v: string) => {
  if (/[\u0000-\u001f]/.test(v)) throw new Denied("Text may not contain control characters");
  return `'${v.replaceAll("'", "''")}'`;
};
const like = (v: string) => quote(`%${v.toUpperCase()}%`);

function id(value: string, what: string): string {
  if (!ID.test(value)) throw new Denied(`${what} must be a NetSuite internal id (digits)`);
  return value;
}

function date(value: string, what: string): string {
  if (!DATE.test(value) || Number.isNaN(Date.parse(value)))
    throw new Denied(`${what} must be a date as YYYY-MM-DD`);
  return `TO_DATE('${value}', 'YYYY-MM-DD')`;
}

function scope(ctx: ToolContext) {
  const read = (limit: string) => {
    const values = listLimit(ctx.limits, limit);
    return values && values.map((v) => id(v, "A limit value"));
  };
  return { subsidiaries: read("subsidiaries"), locations: read("locations") };
}

const within = (list: string[] | null, value: unknown) =>
  !list || (value !== null && value !== undefined && list.includes(String(value)));

function checkLocation(ctx: ToolContext, location: string | undefined): string | undefined {
  const { locations } = scope(ctx);
  if (location !== undefined) id(location, "location");
  if (locations && location === undefined)
    throw new Denied(`This agent must name a location: one of ${locations.join(", ")}`);
  if (location !== undefined && !within(locations, location))
    throw new Denied(`This agent may not use location ${location}`);
  return location;
}

// Paging: the cursor is "<offset>:<page size>" as an earlier answer gave it, since NetSuite
// pages by offset in steps of the page size.
function page(cursor: string | undefined, size: number) {
  if (cursor === undefined) return { offset: 0, limit: size };
  const m = /^(\d{1,7}):(\d{1,4})$/.exec(cursor);
  if (!m || Number(m[1]) % Number(m[2]))
    throw new Denied("Pass the cursor exactly as the previous answer gave it");
  return { offset: Number(m[1]), limit: Number(m[2]) };
}

async function suiteql(
  ctx: ToolContext,
  mark: Mark,
  q: string,
  at: { offset: number; limit: number },
  hintTool?: string,
) {
  const data = await ok(
    ctx,
    mark,
    {
      method: "POST",
      path: `${SUITEQL_PATH}?limit=${at.limit}&offset=${at.offset}`,
      body: { q },
    },
    hintTool,
  );
  const items = ((data?.items ?? []) as Row[]).map(({ links: _links, ...row }) => row);
  return { items, next: data?.hasMore ? `${at.offset + at.limit}:${at.limit}` : null };
}

const where = (parts: (string | null | false | undefined)[]) => {
  const kept = parts.filter(Boolean);
  return kept.length ? `WHERE ${kept.join(" AND ")}` : "";
};

const maxResults = (max: number, fallback: number) =>
  z.number().int().min(1).max(max).default(fallback).describe(`At most ${max}.`);
const cursor = z
  .string()
  .optional()
  .describe("The `next` value of a previous answer, to read on. Leave out for the first page.");

// Customers and vendors share the entity columns.
async function findParties(
  ctx: ToolContext,
  mark: Mark,
  table: "customer" | "vendor",
  args: { id?: string; name?: string; email?: string; include_inactive: boolean },
  at: { offset: number; limit: number },
) {
  const { subsidiaries } = scope(ctx);
  const q =
    "SELECT e.id, e.entityid AS number, e.altname AS name, e.companyname AS company_name, " +
    "e.email, e.phone, e.subsidiary AS subsidiary_id, BUILTIN.DF(e.subsidiary) AS subsidiary, " +
    `e.isinactive AS inactive FROM ${table} e ` +
    where([
      args.id && `e.id = ${id(args.id, "id")}`,
      args.name &&
        `(UPPER(e.entityid) LIKE ${like(args.name)} OR UPPER(e.altname) LIKE ${like(args.name)} OR UPPER(e.companyname) LIKE ${like(args.name)})`,
      args.email && `UPPER(e.email) = ${quote(args.email.toUpperCase())}`,
      !args.include_inactive && "e.isinactive = 'F'",
      subsidiaries && `e.subsidiary IN (${subsidiaries.join(", ")})`,
    ]) +
    " ORDER BY e.id";
  const { items, next } = await suiteql(ctx, mark, q, at);
  return { [`${table}s`]: items.filter((r) => within(subsidiaries, r.subsidiary_id)), next };
}

type TxType = "SalesOrd" | "CustInvc" | "PurchOrd";
// Status letters per type: sales and purchase orders C cancelled or rejected, G billed,
// H closed; invoices A open.
const OPEN: Record<TxType, string> = {
  SalesOrd: "t.status NOT IN ('C', 'G', 'H')",
  PurchOrd: "t.status NOT IN ('C', 'G', 'H')",
  CustInvc: "t.status = 'A'",
};

async function findTransactions(
  ctx: ToolContext,
  mark: Mark,
  type: TxType,
  args: {
    id?: string;
    number?: string;
    entity?: string;
    from?: string;
    to?: string;
    open_only: boolean;
    overdue_only?: boolean;
    include_lines: boolean;
  },
  at: { offset: number; limit: number },
) {
  const { subsidiaries, locations } = scope(ctx);
  const invoice = type === "CustInvc";
  const q =
    "SELECT t.id, t.tranid AS number, TO_CHAR(t.trandate, 'YYYY-MM-DD') AS date, " +
    "BUILTIN.DF(t.status) AS status, t.entity AS entity_id, BUILTIN.DF(t.entity) AS entity, " +
    "tl.subsidiary AS subsidiary_id, BUILTIN.DF(tl.subsidiary) AS subsidiary, " +
    "tl.location AS location_id, BUILTIN.DF(tl.location) AS location, " +
    "t.foreigntotal AS total, BUILTIN.DF(t.currency) AS currency, t.memo" +
    (invoice
      ? ", t.foreignamountremaining AS amount_due, TO_CHAR(t.duedate, 'YYYY-MM-DD') AS due_date"
      : "") +
    " FROM transaction t JOIN transactionline tl ON tl.transaction = t.id AND tl.mainline = 'T' " +
    where([
      `t.type = '${type}'`,
      args.id && `t.id = ${id(args.id, "id")}`,
      args.number && `UPPER(t.tranid) = ${quote(args.number.toUpperCase())}`,
      args.entity &&
        `t.entity = ${id(args.entity, invoice || type === "SalesOrd" ? "customer" : "vendor")}`,
      args.from && `t.trandate >= ${date(args.from, "from")}`,
      args.to && `t.trandate <= ${date(args.to, "to")}`,
      (args.open_only || args.overdue_only) && OPEN[type],
      args.overdue_only && "t.duedate < SYSDATE",
      subsidiaries && `tl.subsidiary IN (${subsidiaries.join(", ")})`,
      locations && `tl.location IN (${locations.join(", ")})`,
    ]) +
    " ORDER BY t.trandate DESC, t.id DESC";
  const found = await suiteql(ctx, mark, q, at);
  const items = found.items.filter(
    (r) => within(subsidiaries, r.subsidiary_id) && within(locations, r.location_id),
  );
  if (!args.include_lines || !items.length) return { transactions: items, next: found.next };

  // Sales lines are stored with the sign of a credit: turned so a sold line reads positive.
  const sign = type === "PurchOrd" ? "" : "-";
  const lines = await suiteql(
    ctx,
    mark,
    "SELECT tl.transaction AS transaction_id, tl.linesequencenumber AS line, tl.item AS item_id, " +
      `BUILTIN.DF(tl.item) AS item, ${sign}tl.quantity AS quantity, tl.rate, ${sign}tl.foreignamount AS amount, ` +
      "tl.location AS location_id, BUILTIN.DF(tl.location) AS location, tl.memo AS description " +
      `FROM transactionline tl WHERE tl.transaction IN (${items.map((r) => id(String(r.id), "id")).join(", ")}) ` +
      "AND tl.mainline = 'F' AND tl.taxline = 'F' AND tl.item IS NOT NULL " +
      "ORDER BY tl.transaction, tl.linesequencenumber",
    { offset: 0, limit: 1000 },
  );
  // A line at a location outside the key's (the header's is inside) is left out.
  const shown = lines.items.filter(
    (l) => l.location_id == null || within(locations, l.location_id),
  );
  return {
    transactions: items.map((t) => ({
      ...t,
      lines: shown
        .filter((l) => String(l.transaction_id) === String(t.id))
        .map(({ transaction_id: _t, ...l }) => l),
    })),
    lines_hidden: lines.items.length - shown.length,
    next: found.next,
  };
}

const createdId = (data: unknown): string => {
  const location = (data as { location?: unknown } | null)?.location;
  const made = typeof location === "string" ? location.split("/").pop() : undefined;
  if (!made || !ID.test(made)) throw new Error("NetSuite did not say which record it made");
  return made;
};

// The saved search script answers {columns: [{name, join, label}], rows: [[{value, text}]],
// pages, total}. Columns are named by label (else name); a key's subsidiaries or locations need
// the search to show that column (no join) so each row can be checked.
type Column = { name: string; join: string | null; label: string | null };
type Cell = { value: unknown; text: unknown };

export function searchRows(
  answer: { columns: Column[]; rows: Cell[][] },
  limits: { subsidiaries: string[] | null; locations: string[] | null },
) {
  const index = (name: string, list: string[] | null, label: string) => {
    if (!list) return -1;
    const at = answer.columns.findIndex((c) => c.name === name && !c.join);
    if (at < 0)
      throw new Denied(
        `This saved search has no ${label} column, so this agent, limited to some ${label.toLowerCase()} values, cannot run it. Add a ${label} column to the search in NetSuite`,
      );
    return at;
  };
  const sub = index("subsidiary", limits.subsidiaries, "Subsidiary");
  const loc = index("location", limits.locations, "Location");
  const keys: string[] = [];
  for (const c of answer.columns) {
    const base = c.label || (c.join ? `${c.join}.${c.name}` : c.name);
    let key = base;
    for (let n = 2; keys.includes(key); n++) key = `${base} (${n})`;
    keys.push(key);
  }
  const kept = answer.rows.filter(
    (r) =>
      (sub < 0 || within(limits.subsidiaries, r[sub]?.value)) &&
      (loc < 0 || within(limits.locations, r[loc]?.value)),
  );
  return {
    rows: kept.map((r) =>
      Object.fromEntries(keys.map((k, i) => [k, r[i]?.text || r[i]?.value || null])),
    ),
    hidden: answer.rows.length - kept.length,
  };
}

export const netsuiteTools: Toolset = {
  names: [
    "netsuite_find_customers",
    "netsuite_find_vendors",
    "netsuite_find_sales_orders",
    "netsuite_find_invoices",
    "netsuite_find_purchase_orders",
    "netsuite_item_availability",
    "netsuite_create_customer",
    "netsuite_create_sales_order",
    "netsuite_run_saved_search",
  ],

  register(server, ctx, allowed) {
    const readOnly = { readOnlyHint: true, openWorldHint: true } as const;
    const write = { readOnlyHint: false, destructiveHint: false, openWorldHint: true } as const;
    const text = (max: number) => z.string().min(1).max(max);

    for (const [tool, table] of [
      ["netsuite_find_customers", "customer"],
      ["netsuite_find_vendors", "vendor"],
    ] as const)
      if (allowed.has(tool))
        server.registerTool(
          tool,
          {
            description:
              `Find ${table}s by internal id, by name or number (part of it is enough) or by exact email. ` +
              "Leave every filter out to list them. Answers the internal id agents pass to other tools.",
            inputSchema: z.object({
              id: z.string().optional().describe("Internal id."),
              name: text(100).optional().describe("Part of the name, company name or number."),
              email: text(254).optional(),
              include_inactive: z.boolean().default(false),
              max_results: maxResults(100, 25),
              cursor,
            }),
            annotations: readOnly,
          },
          ({ max_results, cursor: c, ...args }) =>
            audited(ctx, tool, JSON.stringify(args).slice(0, 300), async (mark) =>
              json(await findParties(ctx, mark, table, args, page(c, max_results))),
            ),
        );

    const transactionTools = [
      ["netsuite_find_sales_orders", "SalesOrd", "sales orders", "customer"],
      ["netsuite_find_invoices", "CustInvc", "invoices", "customer"],
      ["netsuite_find_purchase_orders", "PurchOrd", "purchase orders", "vendor"],
    ] as const;
    for (const [tool, type, what, party] of transactionTools)
      if (allowed.has(tool))
        server.registerTool(
          tool,
          {
            description:
              `Find ${what}, newest first, with status, total` +
              (type === "CustInvc" ? ", amount due and due date" : "") +
              `. Filter by internal id, document number, ${party} or date range. ` +
              "Set include_lines to also get each one's item lines (then at most 10 are returned).",
            inputSchema: z.object({
              id: z.string().optional().describe("Internal id."),
              number: text(45).optional().describe("Document number, e.g. SO1234."),
              [party]: z.string().optional().describe(`The ${party}'s internal id.`),
              from: z.string().optional().describe("Earliest date, YYYY-MM-DD."),
              to: z.string().optional().describe("Latest date, YYYY-MM-DD."),
              open_only: z
                .boolean()
                .default(false)
                .describe(
                  type === "CustInvc"
                    ? "Only invoices not yet paid."
                    : "Leave out cancelled, fully billed and closed ones.",
                ),
              ...(type === "CustInvc"
                ? {
                    overdue_only: z
                      .boolean()
                      .default(false)
                      .describe("Only unpaid invoices past due."),
                  }
                : {}),
              include_lines: z.boolean().default(false),
              max_results: maxResults(100, 25),
              cursor,
            }),
            annotations: readOnly,
          },
          (input: Record<string, any>) => {
            const { max_results, cursor: c, ...args } = input;
            return audited(ctx, tool, JSON.stringify(args).slice(0, 300), async (mark) => {
              const size = args.include_lines ? Math.min(max_results, 10) : max_results;
              return json(
                await findTransactions(
                  ctx,
                  mark,
                  type,
                  { ...(args as any), entity: args[party] },
                  page(c, size),
                ),
              );
            });
          },
        );

    if (allowed.has("netsuite_item_availability"))
      server.registerTool(
        "netsuite_item_availability",
        {
          description:
            "How many of an item are available, on hand, committed, on order and back ordered, per location. " +
            "Give the item's internal id, or part of its number or name.",
          inputSchema: z.object({
            item: text(100).describe("Internal id, or part of the item number or name."),
            location: z.string().optional().describe("One location's internal id."),
            max_results: maxResults(200, 50),
            cursor,
          }),
          annotations: readOnly,
        },
        ({ item, location, max_results, cursor: c }) =>
          audited(
            ctx,
            "netsuite_item_availability",
            `${item} @ ${location ?? "any"}`,
            async (mark) => {
              const { subsidiaries, locations } = scope(ctx);
              if (location !== undefined) {
                id(location, "location");
                if (!within(locations, location))
                  throw new Denied(`This agent may not use location ${location}`);
              }
              const q =
                "SELECT a.item AS item_id, i.itemid AS item_number, i.displayname AS item_name, " +
                "a.location AS location_id, BUILTIN.DF(a.location) AS location, l.subsidiary AS subsidiary_id, " +
                "a.quantityavailable AS available, a.quantityonhand AS on_hand, a.quantitycommitted AS committed, " +
                "a.quantityonorder AS on_order, a.quantitybackordered AS back_ordered " +
                "FROM aggregateitemlocation a JOIN item i ON i.id = a.item JOIN location l ON l.id = a.location " +
                where([
                  ID.test(item)
                    ? `i.id = ${item}`
                    : `(UPPER(i.itemid) LIKE ${like(item)} OR UPPER(i.displayname) LIKE ${like(item)})`,
                  "i.isinactive = 'F'",
                  location && `a.location = ${location}`,
                  locations && `a.location IN (${locations.join(", ")})`,
                  subsidiaries && `l.subsidiary IN (${subsidiaries.join(", ")})`,
                ]) +
                " ORDER BY i.itemid, a.location";
              const { items, next } = await suiteql(ctx, mark, q, page(c, max_results));
              const shown = items.filter(
                (r) => within(locations, r.location_id) && within(subsidiaries, r.subsidiary_id),
              );
              return json({ availability: shown, next });
            },
          ),
      );

    if (allowed.has("netsuite_create_customer"))
      server.registerTool(
        "netsuite_create_customer",
        {
          description:
            "Create a company customer. Find it first with netsuite_find_customers so it is not made twice.",
          inputSchema: z.object({
            company_name: text(83),
            email: text(254).optional(),
            phone: text(32).optional(),
            subsidiary: z
              .string()
              .optional()
              .describe(
                "Internal id of its primary subsidiary. Needed where the account has subsidiaries.",
              ),
          }),
          annotations: write,
        },
        ({ company_name, email, phone, subsidiary }) =>
          audited(ctx, "netsuite_create_customer", company_name, async (mark, retarget) => {
            const { subsidiaries } = scope(ctx);
            if (subsidiary !== undefined) id(subsidiary, "subsidiary");
            if (subsidiaries && !within(subsidiaries, subsidiary))
              throw new Denied(
                `This agent must name a subsidiary: one of ${subsidiaries.join(", ")}`,
              );
            const data = await ok(ctx, mark, {
              method: "POST",
              path: `${RECORD}/customer`,
              body: {
                isPerson: false,
                companyName: company_name,
                ...(email ? { email } : {}),
                ...(phone ? { phone } : {}),
                ...(subsidiary ? { subsidiary: { id: subsidiary } } : {}),
              },
            });
            const made = createdId(data);
            retarget(`customer/${made} ${company_name}`);
            return json({ id: made });
          }),
      );

    if (allowed.has("netsuite_create_sales_order"))
      server.registerTool(
        "netsuite_create_sales_order",
        {
          description:
            "Create a sales order for a customer with item lines. Leave rate out to use the item's price for that customer. " +
            "Check the customer's open orders first so the order is not placed twice.",
          inputSchema: z.object({
            customer: z.string().describe("The customer's internal id."),
            lines: z
              .array(
                z.object({
                  item: z.string().describe("The item's internal id."),
                  quantity: z.number().positive().max(1e9),
                  rate: z.number().min(0).max(1e12).optional().describe("Price per unit."),
                }),
              )
              .min(1)
              .max(100),
            location: z.string().optional().describe("Internal id of the location it ships from."),
            ship_date: z.string().optional().describe("YYYY-MM-DD."),
            po_number: text(45).optional().describe("The customer's purchase order number."),
            memo: text(999).optional(),
          }),
          annotations: write,
        },
        ({ customer, lines, location, ship_date, po_number, memo }) =>
          audited(
            ctx,
            "netsuite_create_sales_order",
            `customer ${customer}`,
            async (mark, retarget) => {
              id(customer, "customer");
              lines.forEach((l) => id(l.item, "item"));
              if (ship_date) date(ship_date, "ship_date");
              checkLocation(ctx, location);
              const { subsidiaries } = scope(ctx);
              let subsidiary: string | undefined;
              if (subsidiaries) {
                // The order takes the customer's subsidiary, so the customer is read first.
                const { items } = await suiteql(
                  ctx,
                  mark,
                  `SELECT e.id, e.subsidiary AS subsidiary_id FROM customer e WHERE e.id = ${customer}`,
                  { offset: 0, limit: 1 },
                  "netsuite_find_customers",
                );
                const found = items[0]?.subsidiary_id;
                // Not found and not allowed read the same, so a key cannot probe other customers.
                if (!within(subsidiaries, found))
                  throw new Denied(`This agent may not use customer ${customer}`);
                subsidiary = String(found);
              }
              const data = await ok(ctx, mark, {
                method: "POST",
                path: `${RECORD}/salesOrder`,
                body: {
                  entity: { id: customer },
                  ...(subsidiary ? { subsidiary: { id: subsidiary } } : {}),
                  ...(location ? { location: { id: location } } : {}),
                  ...(ship_date ? { shipDate: ship_date } : {}),
                  ...(po_number ? { otherRefNum: po_number } : {}),
                  ...(memo ? { memo } : {}),
                  item: {
                    items: lines.map((l) => ({
                      item: { id: l.item },
                      quantity: l.quantity,
                      ...(l.rate === undefined ? {} : { rate: l.rate }),
                      ...(location ? { location: { id: location } } : {}),
                    })),
                  },
                },
              });
              const made = createdId(data);
              retarget(`salesOrder/${made} for customer ${customer}`);
              return json({ id: made });
            },
          ),
      );

    if (allowed.has("netsuite_run_saved_search"))
      server.registerTool(
        "netsuite_run_saved_search",
        {
          description:
            "Run one of the saved searches this agent may use, by its id (customsearch... or a number), a page at a time.",
          inputSchema: z.object({
            search: z.string().describe("The saved search's id, as your administrator gave it."),
            max_results: z
              .number()
              .int()
              .min(5)
              .max(1000)
              .default(100)
              .describe("Rows per page, 5 to 1000."),
            cursor,
          }),
          annotations: readOnly,
        },
        ({ search, max_results, cursor: c }) =>
          audited(ctx, "netsuite_run_saved_search", search, async (mark) => {
            const searches = listLimit(ctx.limits, "saved_searches");
            if (!searches)
              throw new Denied(
                "This agent has no saved searches. An administrator adds their ids to its key",
              );
            if (!SEARCH_ID.test(search) || !searches.includes(search))
              throw new Denied(
                `This agent may run only these saved searches: ${searches.join(", ")}`,
              );
            const at = page(c, max_results);
            const index = at.offset / at.limit;
            const data = await ok(ctx, mark, {
              method: "POST",
              path: `${RESTLET_PATH}?${SAVED_SEARCH_SCRIPT}`,
              body: { search, page: index, size: at.limit },
            });
            if (!Array.isArray(data?.columns) || !Array.isArray(data?.rows))
              throw new VendorFailed(
                502,
                "The saved search script answered something else. Check it is the script from the setup steps",
              );
            const { rows, hidden } = searchRows(data, scope(ctx));
            return json({
              total: data.total,
              rows,
              ...(hidden ? { rows_hidden: hidden } : {}),
              next: index + 1 < Number(data.pages) ? `${at.offset + at.limit}:${at.limit}` : null,
            });
          }),
      );
  },
};
