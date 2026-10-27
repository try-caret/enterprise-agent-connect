import { z } from "zod";
import { soqlObject, SoqlRejected } from "./soql.js";
import {
  audited,
  Denied,
  json,
  listLimit,
  VendorFailed,
  type ToolContext,
  type Toolset,
} from "../../packages/connector-kit/src/common.js";

const API = "/services/data/v63.0";
const NAME = /^[A-Za-z][A-Za-z0-9_]{0,79}$/;
const RECORD_ID = /^[A-Za-z0-9]{15}(?:[A-Za-z0-9]{3})?$/;

type SfRecord = { attributes?: { type?: string }; [field: string]: unknown };

async function ok(
  ctx: ToolContext,
  mark: (s: number) => void,
  req: Parameters<ToolContext["call"]>[0],
) {
  const res = await ctx.call(req);
  mark(res.status);
  if (res.status < 200 || res.status >= 300) {
    const errors = Array.isArray(res.data)
      ? (res.data as { errorCode?: string; message?: string }[])
      : [];
    const detail = errors
      .map((e) => [e.errorCode, e.message].filter(Boolean).join(": "))
      .join("; ");
    throw new VendorFailed(
      res.status,
      `Salesforce answered ${res.status}${detail ? `: ${detail}` : ""}`,
    );
  }
  return res.data as any;
}

// Granted objects, lowercased: Salesforce object names are case-insensitive.
function granted(ctx: ToolContext): Set<string> | null {
  const objects = listLimit(ctx.limits, "objects");
  return objects && new Set(objects.map((o) => o.toLowerCase()));
}

function checkObject(ctx: ToolContext, object: string): string {
  if (!NAME.test(object)) throw new Denied("Give an object API name such as Account or Invoice__c");
  const allowed = granted(ctx);
  if (allowed && !allowed.has(object.toLowerCase()))
    throw new Denied(`This agent is not allowed to use ${object} records`);
  return object;
}

function recordPath(ctx: ToolContext, object: string, id: string): string {
  checkObject(ctx, object);
  if (!RECORD_ID.test(id)) throw new Denied("Give a 15 or 18 character record id");
  return `${API}/sobjects/${object}/${id}`;
}

// The id must belong to the granted object: Salesforce can serve a record under another
// object's path, so its type is read back before it is changed.
async function checkRecordType(
  ctx: ToolContext,
  mark: (s: number) => void,
  object: string,
  id: string,
): Promise<string> {
  const path = recordPath(ctx, object, id);
  const current: SfRecord = await ok(ctx, mark, { method: "GET", path: `${path}?fields=Id` });
  if (current.attributes?.type?.toLowerCase() !== object.toLowerCase())
    throw new Denied(`${id} is not a record of type ${object}`);
  return path;
}

// Salesforce deletes a record's cascade children with it (master-detail, and standard ones
// such as an Account's Contacts), and their cascade children in turn. With an objects limit,
// a delete that would take records of other objects along, at any depth, is refused. Not
// counted: Salesforce's own bookkeeping that goes with any record (feeds, field history,
// sharing rows, change events) and children no one can query. A child only automation
// creates (not createable) is still real data, so it counts.
const SYSTEM_CHILD = /(Feed|History|Share|ChangeEvent)$/;
type ChildRelationship = { childSObject: string; cascadeDelete: boolean };
type SObjectInfo = { name: string; queryable?: boolean };

export function cascadeChildren(children: ChildRelationship[], usable: Set<string>): string[] {
  const names = children
    .filter((r) => r.cascadeDelete && !SYSTEM_CHILD.test(r.childSObject))
    .map((r) => r.childSObject)
    .filter((n) => usable.has(n.toLowerCase()));
  return [...new Set(names)];
}

export function cascadeRefusal(object: string, children: string[]): string {
  const shown = children.slice(0, 10).join(", ");
  const more = children.length > 10 ? ` and ${children.length - 10} more` : "";
  return `Deleting this ${object} record also deletes its related ${shown}${more} records, which this agent is not allowed to use`;
}

// Breadth-first through the allowed cascade children: each allowed object is described once,
// so the walk ends after at most one describe per object in the limit.
async function checkCascade(ctx: ToolContext, mark: (s: number) => void, object: string) {
  const allowed = granted(ctx);
  if (!allowed) return;
  const global = await ok(ctx, mark, { method: "GET", path: `${API}/sobjects` });
  const usable = new Set(
    (global.sobjects as SObjectInfo[]).filter((o) => o.queryable).map((o) => o.name.toLowerCase()),
  );
  const outside = new Set<string>();
  const seen = new Set([object.toLowerCase()]);
  for (let level = [object]; level.length;) {
    const next: string[] = [];
    for (const name of level) {
      const described = await ok(ctx, mark, {
        method: "GET",
        path: `${API}/sobjects/${name}/describe`,
      });
      for (const child of cascadeChildren(described.childRelationships, usable)) {
        const k = child.toLowerCase();
        if (!allowed.has(k)) outside.add(child);
        else if (!seen.has(k)) {
          seen.add(k);
          next.push(child);
        }
      }
    }
    level = next;
  }
  if (outside.size) throw new Denied(cascadeRefusal(object, [...outside]));
}

const fieldValues = z
  .record(z.string(), z.union([z.string(), z.number(), z.boolean(), z.null()]))
  .describe("Field API names and their values.");

function checkFields(fields: Record<string, unknown>): string[] {
  const names = fieldNames(Object.keys(fields));
  if (names.length === 0 || names.length > 100) throw new Denied("Give between 1 and 100 fields");
  return names;
}

function fieldNames(fields: string[]): string[] {
  const bad = fields.find((f) => !NAME.test(f));
  if (bad) throw new Denied(`${bad} is not a field API name on this object`);
  return fields;
}

export const salesforceTools: Toolset = {
  names: [
    "salesforce_query",
    "salesforce_get_record",
    "salesforce_create_record",
    "salesforce_update_record",
    "salesforce_delete_record",
  ],

  register(server, ctx, allowed) {
    const readOnly = { readOnlyHint: true, openWorldHint: true } as const;
    const fieldList = z.array(z.string()).max(100);

    if (allowed.has("salesforce_query"))
      server.registerTool(
        "salesforce_query",
        {
          description:
            "Run a read-only SOQL query, e.g. `SELECT Id, Name FROM Account WHERE Industry = 'Energy' LIMIT 20`. " +
            "If this agent is limited to some objects, query one of them at a time: no subqueries or relationship fields.",
          inputSchema: z.object({
            soql: z.string().min(1).max(10000),
            max_records: z.number().int().min(1).max(200).default(50),
          }),
          annotations: readOnly,
        },
        ({ soql, max_records }) =>
          audited(ctx, "salesforce_query", soql, async (mark) => {
            const allowed = granted(ctx);
            if (allowed) {
              let object: string;
              try {
                object = soqlObject(soql);
              } catch (error) {
                if (error instanceof SoqlRejected) throw new Denied(error.message);
                throw error;
              }
              checkObject(ctx, object);
            }
            const data = await ok(ctx, mark, {
              method: "GET",
              path: `${API}/query?q=${encodeURIComponent(soql)}`,
              headers: { "sforce-query-options": "batchSize=200" },
            });
            const records = (data.records as SfRecord[]).filter((r) => {
              const type = r.attributes?.type?.toLowerCase();
              return !allowed || type === "aggregateresult" || (type && allowed.has(type));
            });
            return json({
              total: data.totalSize,
              records: records.slice(0, max_records),
              more: records.length > max_records || data.done === false,
            });
          }),
      );

    if (allowed.has("salesforce_get_record"))
      server.registerTool(
        "salesforce_get_record",
        {
          description: "Read one Salesforce record by object and id, optionally only some fields.",
          inputSchema: z.object({
            object: z.string().describe("Object API name, e.g. Account or Invoice__c."),
            id: z.string().describe("15 or 18 character record id."),
            fields: fieldList
              .optional()
              .describe("Field API names. All readable fields if omitted."),
          }),
          annotations: readOnly,
        },
        ({ object, id, fields }) =>
          audited(ctx, "salesforce_get_record", `${object}/${id}`, async (mark) => {
            const path = recordPath(ctx, object, id);
            const query = fields?.length ? `?fields=${fieldNames(fields).join(",")}` : "";
            const record: SfRecord = await ok(ctx, mark, { method: "GET", path: path + query });
            if (record.attributes?.type?.toLowerCase() !== object.toLowerCase())
              throw new Denied(`${id} is not a record of type ${object}`);
            return json(record);
          }),
      );

    if (allowed.has("salesforce_update_record"))
      server.registerTool(
        "salesforce_update_record",
        {
          description: "Update fields on an existing Salesforce record.",
          inputSchema: z.object({
            object: z.string().describe("Object API name, e.g. Account or Invoice__c."),
            id: z.string().describe("15 or 18 character record id."),
            fields: fieldValues,
          }),
          annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
        },
        ({ object, id, fields }) =>
          audited(ctx, "salesforce_update_record", `${object}/${id}`, async (mark) => {
            recordPath(ctx, object, id);
            const names = checkFields(fields);
            const path = await checkRecordType(ctx, mark, object, id);
            await ok(ctx, mark, { method: "PATCH", path, body: fields });
            return json({ updated: true, fields: names });
          }),
      );

    if (allowed.has("salesforce_create_record"))
      server.registerTool(
        "salesforce_create_record",
        {
          description: "Create a Salesforce record of an object this agent may use.",
          inputSchema: z.object({
            object: z.string().describe("Object API name, e.g. Account or Invoice__c."),
            fields: fieldValues,
          }),
          annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
        },
        ({ object, fields }) =>
          audited(ctx, "salesforce_create_record", object, async (mark, retarget) => {
            checkObject(ctx, object);
            checkFields(fields);
            const created = await ok(ctx, mark, {
              method: "POST",
              path: `${API}/sobjects/${object}`,
              body: fields,
            });
            retarget(`${object}/${created.id}`);
            return json({ id: created.id });
          }),
      );

    if (allowed.has("salesforce_delete_record"))
      server.registerTool(
        "salesforce_delete_record",
        {
          description:
            "Delete a Salesforce record. It goes to the Recycle Bin, where Salesforce keeps it for 15 days.",
          inputSchema: z.object({
            object: z.string().describe("Object API name, e.g. Account or Invoice__c."),
            id: z.string().describe("15 or 18 character record id."),
          }),
          annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
        },
        ({ object, id }) =>
          audited(ctx, "salesforce_delete_record", `${object}/${id}`, async (mark) => {
            const path = await checkRecordType(ctx, mark, object, id);
            await checkCascade(ctx, mark, object);
            await ok(ctx, mark, { method: "DELETE", path });
            return json({ deleted: true });
          }),
      );
  },
};
