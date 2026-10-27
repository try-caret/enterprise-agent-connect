import { z } from "zod";
import {
  audited,
  Denied,
  json,
  listLimit,
  mapLimit,
  type ToolContext,
  type Toolset,
} from "../../packages/connector-kit/src/common.js";
import {
  checkItem,
  cursorArg,
  inSite,
  nextCursor,
  ok,
  openSite,
  pageQuery,
  targetFolder,
  type DriveItem,
} from "./tools.js";

// Excel workbooks (.xlsx) in the SharePoint sites of the agent's grant (its `sites` limit,
// and the folder its `folders` limit names there), through the Excel API of Microsoft Graph,
// with the files tools' checks. A workbook is read back by id and must sit in a granted site,
// in its document library and inside the key's folder. Reads are sessionless; each edit runs in its own persistent session.

export const MAX_CELLS = 10_000;
export const MAX_TABLE_ROWS = 500;
const XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const ITEM_ID = /^[A-Za-z0-9!_-]{1,256}$/;
const CELL = /^([A-Z]{1,3})([1-9][0-9]{0,6})$/;
const MAX_COLUMN = 16_384;
const MAX_ROW = 1_048_576;
// Excel forbids these in sheet names, and caps them at 31 characters.
const SHEET = /^[^[\]:*?/\\]{1,31}$/;
const TABLE = /^[A-Za-z0-9_.{}\\-]{1,255}$/;

type Cell = string | number | boolean | null;

function column(letters: string): number {
  return [...letters].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
}

/** A1 or A1:D20, upper-cased, with its size. Whole rows or columns are refused: their size
 *  is the whole sheet. */
export function parseRange(raw: string): { address: string; rows: number; columns: number } {
  const [first, last = first, extra] = raw.trim().toUpperCase().split(":");
  const a = CELL.exec(first ?? "");
  const b = CELL.exec(last);
  if (extra !== undefined || !a || !b) throw new Denied(`${raw} is not a range such as A1:D20`);
  const [c1, c2] = [column(a[1]), column(b[1])];
  const [r1, r2] = [Number(a[2]), Number(b[2])];
  if (Math.max(c1, c2) > MAX_COLUMN || Math.max(r1, r2) > MAX_ROW)
    throw new Denied(`${raw} is outside the sheet`);
  const rows = Math.abs(r2 - r1) + 1;
  const columns = Math.abs(c2 - c1) + 1;
  if (rows * columns > MAX_CELLS)
    throw new Denied(`Read or change at most ${MAX_CELLS} cells at a time`);
  return { address: first === last ? first : `${first}:${last}`, rows, columns };
}

function sheetSegment(name: string): string {
  if (!SHEET.test(name)) throw new Denied(`${name} is not a worksheet name`);
  return encodeURIComponent(name).replaceAll("'", "%27");
}

function tableSegment(name: string): string {
  if (!TABLE.test(name)) throw new Denied("Give a table name from excel_list_tables");
  return encodeURIComponent(name);
}

export function isWorkbook(item: DriveItem): boolean {
  const type = item.file?.mimeType;
  return (
    Boolean(item.file) &&
    /\.xlsx$/i.test(item.name ?? "") &&
    (type === undefined || type === XLSX_TYPE)
  );
}

/** The workbook's Excel API path, once its item is read back and found in a granted site. */
async function workbookPath(
  ctx: ToolContext,
  mark: (s: number) => void,
  site: string,
  itemId: string,
): Promise<string> {
  if (!ITEM_ID.test(itemId)) throw new Denied("Give a workbook id from excel_list_workbooks");
  const opened = await openSite(ctx, mark, site);
  const path = `/v1.0/sites/${encodeURIComponent(opened.id)}/drive/items/${encodeURIComponent(itemId)}`;
  const item: DriveItem = await ok(ctx, mark, {
    method: "GET",
    path: `${path}?$select=id,name,file,folder,parentReference,remoteItem`,
  });
  await checkItem(ctx, mark, opened, item);
  if (!isWorkbook(item)) throw new Denied("Only Excel workbooks (.xlsx) can be opened");
  return `${path}/workbook`;
}

// Runs edits in a session that saves them, and closes it whatever happens.
async function inSession<T>(
  ctx: ToolContext,
  mark: (s: number) => void,
  workbook: string,
  run: (headers: Record<string, string>) => Promise<T>,
): Promise<T> {
  const session = await ok(ctx, mark, {
    method: "POST",
    path: `${workbook}/createSession`,
    body: { persistChanges: true },
  });
  if (typeof session?.id !== "string") throw new Error("Excel returned no session id");
  const headers = { "workbook-session-id": session.id };
  try {
    return await run(headers);
  } finally {
    const closed = await ctx.call({ method: "POST", path: `${workbook}/closeSession`, headers });
    if (closed.status >= 300)
      console.error(`[agent-connect] closeSession answered ${closed.status}`);
  }
}

// Excel runs text starting with = (and +, - or @, which it turns into =) as a formula, and a
// formula can reach other workbooks or outside addresses (external links, WEBSERVICE,
// HYPERLINK) when a person opens the file. A leading apostrophe keeps it as the text written;
// tab and carriage return are on OWASP's list of formula starts too.
const FORMULA_START = /^[=+\-@\t\r]/;
export const literal = (rows: Cell[][]): Cell[][] =>
  rows.map((row) => row.map((v) => (typeof v === "string" && FORMULA_START.test(v) ? `'${v}` : v)));

function checkShape(values: Cell[][], rows: number, columns: number) {
  if (values.length !== rows || values.some((r) => r.length !== columns))
    throw new Denied(`values must be ${rows} row(s) of ${columns} cell(s), the size of the range`);
}

export const excelTools: Toolset = {
  names: [
    "excel_list_workbooks",
    "excel_list_worksheets",
    "excel_read_range",
    "excel_list_tables",
    "excel_read_table",
    "excel_update_range",
    "excel_add_table_rows",
  ],

  register(server, ctx, allowed) {
    const readOnly = { readOnlyHint: true, openWorldHint: true } as const;
    const edits = { readOnlyHint: false, destructiveHint: false, openWorldHint: true } as const;
    // The agent may have no files tool to list its sites, so they are named here.
    const sites = listLimit(ctx.limits, "sites") ?? [];
    const folders = mapLimit(ctx.limits, "folders") ?? {};
    const named = sites.map((s) => (folders[s] ? `${s} (only its ${folders[s]} folder)` : s));
    const site = z
      .string()
      .max(500)
      .describe(`Site address, one of: ${named.join(", ") || "none (ask an administrator)"}`);
    const workbook = z.string().max(300).describe("Workbook id from excel_list_workbooks.");
    const worksheet = z.string().max(100).describe("Worksheet name from excel_list_worksheets.");
    const table = z.string().max(255).describe("Table name from excel_list_tables.");
    const cell = z.union([z.string().max(32_767), z.number(), z.boolean(), z.null()]);
    const grid = z.array(z.array(cell).max(MAX_COLUMN)).max(MAX_CELLS);

    if (allowed.has("excel_list_workbooks"))
      server.registerTool(
        "excel_list_workbooks",
        {
          description:
            "List Excel workbooks (.xlsx) and folders in a site's document library, at its top or in a folder. A page can hold few or none while more follow: when `more` is true, call again with `cursor` set to `next`.",
          inputSchema: z.object({
            site,
            folder: z
              .string()
              .max(1000)
              .default("")
              .describe("Folder path, e.g. Reports/2026. Empty: the top, or this agent's folder."),
            max_results: z.number().int().min(1).max(200).default(100),
            cursor: cursorArg,
          }),
          annotations: readOnly,
        },
        ({ site: s, folder, max_results, cursor }) =>
          audited(ctx, "excel_list_workbooks", `${s} ${folder}`.trim(), async (mark) => {
            const path = targetFolder(ctx, s, folder).map(encodeURIComponent).join("/");
            const { id } = await openSite(ctx, mark, s);
            const base = `/v1.0/sites/${encodeURIComponent(id)}/drive/root`;
            const data = await ok(ctx, mark, {
              method: "GET",
              path: `${path ? `${base}:/${path}:` : base}/children?${pageQuery(max_results, cursor)}&$select=id,name,file,folder,lastModifiedDateTime,parentReference,remoteItem`,
            });
            const items = (data.value as DriveItem[]).filter(
              (i) => !i.remoteItem && inSite(i, id) && (i.folder || isWorkbook(i)),
            );
            return json({
              items: items.map((i) => ({
                id: i.id,
                name: i.name ?? "",
                kind: i.folder ? "folder" : "workbook",
                modified: i.lastModifiedDateTime ?? null,
              })),
              more: Boolean(data["@odata.nextLink"]),
              next: nextCursor(data),
            });
          }),
      );

    if (allowed.has("excel_list_worksheets"))
      server.registerTool(
        "excel_list_worksheets",
        {
          description: "List the worksheets of a workbook.",
          inputSchema: z.object({ site, workbook_id: workbook }),
          annotations: readOnly,
        },
        ({ site: s, workbook_id }) =>
          audited(ctx, "excel_list_worksheets", `${s} ${workbook_id}`, async (mark) => {
            const base = await workbookPath(ctx, mark, s, workbook_id);
            const data = await ok(ctx, mark, {
              method: "GET",
              path: `${base}/worksheets?$select=name,position,visibility`,
            });
            return json({
              worksheets: (data.value as { name: string; visibility?: string }[]).map((w) => ({
                name: w.name,
                visible: w.visibility === undefined || w.visibility === "Visible",
              })),
            });
          }),
      );

    if (allowed.has("excel_read_range"))
      server.registerTool(
        "excel_read_range",
        {
          description: `Read the values of a range of cells (at most ${MAX_CELLS}) in a worksheet.`,
          inputSchema: z.object({
            site,
            workbook_id: workbook,
            worksheet,
            range: z.string().max(30).describe("e.g. A1:D20"),
          }),
          annotations: readOnly,
        },
        ({ site: s, workbook_id, worksheet: sheet, range }) =>
          audited(
            ctx,
            "excel_read_range",
            `${s} ${workbook_id} ${sheet}!${range}`,
            async (mark) => {
              const { address } = parseRange(range);
              const segment = sheetSegment(sheet);
              const base = await workbookPath(ctx, mark, s, workbook_id);
              const data = await ok(ctx, mark, {
                method: "GET",
                path: `${base}/worksheets/${segment}/range(address='${address}')?$select=address,values`,
              });
              return json({ address: data.address ?? address, values: data.values ?? [] });
            },
          ),
      );

    if (allowed.has("excel_list_tables"))
      server.registerTool(
        "excel_list_tables",
        {
          description: "List the tables in a workbook.",
          inputSchema: z.object({ site, workbook_id: workbook }),
          annotations: readOnly,
        },
        ({ site: s, workbook_id }) =>
          audited(ctx, "excel_list_tables", `${s} ${workbook_id}`, async (mark) => {
            const base = await workbookPath(ctx, mark, s, workbook_id);
            const data = await ok(ctx, mark, {
              method: "GET",
              path: `${base}/tables?$select=id,name,showHeaders`,
            });
            return json({
              tables: (data.value as { name: string; showHeaders?: boolean }[]).map((t) => ({
                name: t.name,
                headers: t.showHeaders ?? true,
              })),
            });
          }),
      );

    if (allowed.has("excel_read_table"))
      server.registerTool(
        "excel_read_table",
        {
          description: `Read a table's header and rows, at most ${MAX_TABLE_ROWS} rows at a time.`,
          inputSchema: z.object({
            site,
            workbook_id: workbook,
            table,
            max_rows: z.number().int().min(1).max(MAX_TABLE_ROWS).default(100),
            skip: z.number().int().min(0).max(MAX_ROW).default(0),
          }),
          annotations: readOnly,
        },
        ({ site: s, workbook_id, table: t, max_rows, skip }) =>
          audited(ctx, "excel_read_table", `${s} ${workbook_id} ${t}`, async (mark) => {
            const segment = tableSegment(t);
            const base = await workbookPath(ctx, mark, s, workbook_id);
            const header = await ok(ctx, mark, {
              method: "GET",
              path: `${base}/tables/${segment}/headerRowRange?$select=values`,
            });
            // One more than asked, to tell whether more rows follow.
            const data = await ok(ctx, mark, {
              method: "GET",
              path: `${base}/tables/${segment}/rows?$top=${max_rows + 1}&$skip=${skip}`,
            });
            const rows = (data.value as { values?: Cell[][] }[]).map((r) => r.values?.[0] ?? []);
            const cells = rows.slice(0, max_rows).reduce((n, r) => n + r.length, 0);
            if (cells > MAX_CELLS)
              throw new Denied(`That is more than ${MAX_CELLS} cells. Ask for fewer rows.`);
            return json({
              headers: header.values?.[0] ?? [],
              rows: rows.slice(0, max_rows),
              more: rows.length > max_rows,
            });
          }),
      );

    if (allowed.has("excel_update_range"))
      server.registerTool(
        "excel_update_range",
        {
          description: `Write values into a range of cells (at most ${MAX_CELLS}). values must have the range's size; a null leaves that cell as it is. Text is written as text, never as a formula (send numbers as numbers: text such as "-5" stays text).`,
          inputSchema: z.object({
            site,
            workbook_id: workbook,
            worksheet,
            range: z.string().max(30).describe("e.g. B2:C3"),
            values: grid.describe("Rows of cells, e.g. [[1, 2], [3, 4]]"),
          }),
          annotations: edits,
        },
        ({ site: s, workbook_id, worksheet: sheet, range, values }) =>
          audited(
            ctx,
            "excel_update_range",
            `${s} ${workbook_id} ${sheet}!${range}`,
            async (mark) => {
              const { address, rows, columns } = parseRange(range);
              checkShape(values, rows, columns);
              const segment = sheetSegment(sheet);
              const base = await workbookPath(ctx, mark, s, workbook_id);
              const data = await inSession(ctx, mark, base, (headers) =>
                ok(ctx, mark, {
                  method: "PATCH",
                  path: `${base}/worksheets/${segment}/range(address='${address}')`,
                  headers,
                  body: { values: literal(values) },
                }),
              );
              return json({ address: data?.address ?? address, updated: rows * columns });
            },
          ),
      );

    if (allowed.has("excel_add_table_rows"))
      server.registerTool(
        "excel_add_table_rows",
        {
          description: `Add rows (at most ${MAX_TABLE_ROWS}) to the end of a table. Each row has one cell per table column.`,
          inputSchema: z.object({
            site,
            workbook_id: workbook,
            table,
            rows: z.array(z.array(cell).max(MAX_COLUMN)).min(1).max(MAX_TABLE_ROWS),
          }),
          annotations: edits,
        },
        ({ site: s, workbook_id, table: t, rows }) =>
          audited(ctx, "excel_add_table_rows", `${s} ${workbook_id} ${t}`, async (mark) => {
            const segment = tableSegment(t);
            const width = rows[0].length;
            if (width === 0 || rows.some((r) => r.length !== width))
              throw new Denied("Every row needs the same number of cells, one per table column");
            if (rows.length * width > MAX_CELLS)
              throw new Denied(`Add at most ${MAX_CELLS} cells at a time`);
            const base = await workbookPath(ctx, mark, s, workbook_id);
            await inSession(ctx, mark, base, (headers) =>
              ok(ctx, mark, {
                method: "POST",
                path: `${base}/tables/${segment}/rows`,
                headers,
                body: { values: literal(rows) },
              }),
            );
            return json({ added: rows.length });
          }),
      );
  },
};
