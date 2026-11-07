import { z } from "zod";
import { isSchemaId, scopeAql } from "./aql.js";
import { JqlRejected } from "./jql.js";
import { ASSETS_PATH } from "./vendor.js";
import {
  audited,
  Denied,
  json,
  listLimit,
  VendorFailed,
  type ToolContext,
  type Toolset,
} from "../../packages/connector-kit/src/common.js";

// Jira Assets (Jira Service Management's CMDB): find objects with AQL, read one, list an object
// type's attributes, and change named attributes of one existing object. Generic on purpose:
// attribute names and status values differ per customer, so nothing here knows any object
// type. Nothing creates or deletes objects. Paths and shapes from Atlassian's Assets OpenAPI
// (developer.atlassian.com/cloud/assets/swagger.v3.json, read November 7, 2026): POST
// object/aql, GET/PUT object/{id}, GET objecttype/{id}/attributes, GET config/statustype.
// The grant's `asset_schemas` limit (object schema ids) holds every call: AQL is confined to
// them and every object answered or changed is checked on its objectType.objectSchemaId.

type Mark = (status: number) => void;
type Obj = Record<string, any>;

const OBJECT_ID = /^[1-9][0-9]{0,17}$/;
const path = (rest: string) => `${ASSETS_PATH}${rest}`;
// An object type attribute's `type` (ObjectTypeAttribute in the OpenAPI): 0 is a plain value
// whose kind is `defaultType.name` (Text, Date, Select...).
const ATTRIBUTE_TYPES: Record<number, string> = {
  1: "Object reference",
  2: "User",
  4: "Group",
  7: "Status",
};
const STATUS = 7;

function assetsError(res: { status: number; data: unknown }): string {
  const d = res.data as {
    errorMessages?: unknown[];
    errors?: Record<string, unknown>;
    message?: unknown;
  } | null;
  const detail = [
    ...(Array.isArray(d?.errorMessages) ? d.errorMessages : []),
    ...Object.values(d?.errors ?? {}),
    ...(typeof d?.message === "string" ? [d.message] : []),
  ]
    .filter((m) => typeof m === "string" && m)
    .join("; ");
  return `Jira Assets answered ${res.status}${detail ? `: ${detail}` : ""}`;
}

async function ok(ctx: ToolContext, mark: Mark, req: Parameters<ToolContext["call"]>[0]) {
  const res = await ctx.call(req);
  mark(res.status);
  if (res.status < 200 || res.status >= 300) throw new VendorFailed(res.status, assetsError(res));
  return res.data as any;
}

function allowedSchema(schemas: string[] | null, id: unknown): boolean {
  return schemas === null || (id !== undefined && id !== null && schemas.includes(String(id)));
}

// One value as an agent reads it: a referenced object, status, user or group by its name, any
// other value as stored (dates in ISO 8601, not the account's display format).
function shownValue(v: Obj): unknown {
  if (v?.referencedObject) return v.displayValue ?? v.referencedObject.label ?? null;
  if (v?.status) return v.status.name ?? v.displayValue ?? null;
  if (v?.user) return v.user.displayName ?? v.displayValue ?? null;
  if (v?.group) return v.group.name ?? v.displayValue ?? null;
  return v?.value ?? v?.displayValue ?? null;
}

// An object with its attributes by name. `names`: attribute id to name, for answers (AQL)
// whose attributes carry only their objectTypeAttributeId.
function shownObject(o: Obj, names: Map<string, string> = new Map()) {
  const attributes: Record<string, unknown> = {};
  for (const a of (o.attributes ?? []) as Obj[]) {
    const id = String(a.objectTypeAttributeId ?? a.objectTypeAttribute?.id ?? "");
    const name = a.objectTypeAttribute?.name ?? names.get(id) ?? `attribute ${id}`;
    const values = ((a.objectAttributeValues ?? []) as Obj[]).map(shownValue);
    attributes[name] = values.length === 0 ? null : values.length === 1 ? values[0] : values;
  }
  return {
    id: String(o.id),
    key: o.objectKey ?? null,
    label: o.label ?? null,
    object_type: o.objectType?.name ?? null,
    object_type_id: o.objectType?.id === undefined ? null : String(o.objectType.id),
    schema_id:
      o.objectType?.objectSchemaId === undefined ? null : String(o.objectType.objectSchemaId),
    updated: o.updated ?? null,
    attributes,
  };
}

async function fetchObject(ctx: ToolContext, mark: Mark, id: string): Promise<Obj> {
  if (!OBJECT_ID.test(id)) throw new Denied("Give an Assets object id, a number such as 88");
  const object = await ok(ctx, mark, { method: "GET", path: path(`object/${id}`) });
  if (!allowedSchema(listLimit(ctx.limits, "asset_schemas"), object?.objectType?.objectSchemaId))
    throw new Denied(`Assets object ${id} is in an object schema this agent is not allowed to use`);
  return object;
}

// An object type's attributes. Unless the caller already checked the type's object (`checked`),
// refused when any attribute's object type is in a schema outside the limit: each names its own
// object type and schema in the answer of objecttype/{id}/attributes.
async function typeAttributes(
  ctx: ToolContext,
  mark: Mark,
  typeId: string,
  checked = false,
): Promise<Obj[]> {
  if (!OBJECT_ID.test(typeId)) throw new Denied("Give an object type id, a number such as 23");
  const attrs = (await ok(ctx, mark, {
    method: "GET",
    path: path(`objecttype/${typeId}/attributes`),
  })) as Obj[];
  const schemas = listLimit(ctx.limits, "asset_schemas");
  if (
    !checked &&
    schemas !== null &&
    attrs.some((a) => !allowedSchema(schemas, a.objectType?.objectSchemaId))
  )
    throw new Denied(
      `Object type ${typeId} is in an object schema this agent is not allowed to use`,
    );
  return attrs;
}

const attributeKind = (a: Obj) =>
  a.type === 0 || a.type === undefined
    ? (a.defaultType?.name ?? "Text")
    : (ATTRIBUTE_TYPES[a.type] ?? `type ${a.type}`);

export const assetsTools: Toolset = {
  names: [
    "jira_assets_search",
    "jira_assets_get_object",
    "jira_assets_list_attributes",
    "jira_assets_update_object",
  ],

  register(server, ctx, allowed) {
    const readOnly = { readOnlyHint: true, openWorldHint: true } as const;
    const objectId = z.string().max(20).describe("Assets object id, e.g. 88.");

    if (allowed.has("jira_assets_search"))
      server.registerTool(
        "jira_assets_search",
        {
          description:
            "Find Jira Assets objects with AQL, e.g. " +
            '`objectType = Laptop AND Status = "In stock" ORDER BY "Date Received" ASC`. ' +
            "Attribute names with spaces go in double quotes; ORDER BY takes one attribute. " +
            "Each object comes with its attributes by name. Results are limited to the object " +
            "schemas this agent may use. Find attribute names with jira_assets_list_attributes.",
          inputSchema: z.object({
            aql: z.string().min(1).max(4000).describe("AQL query. ORDER BY is allowed at the end."),
            max_results: z.number().int().min(1).max(50).default(20),
            start_at: z.number().int().min(0).max(10000).default(0),
          }),
          annotations: readOnly,
        },
        ({ aql, max_results, start_at }) =>
          audited(ctx, "jira_assets_search", aql, async (mark, retarget) => {
            const schemas = listLimit(ctx.limits, "asset_schemas");
            let scoped: string;
            try {
              scoped = scopeAql(aql, schemas);
            } catch (error) {
              if (error instanceof JqlRejected) throw new Denied(error.message);
              throw error;
            }
            retarget(scoped);
            const data = await ok(ctx, mark, {
              method: "POST",
              path: path(
                `object/aql?startAt=${start_at}&maxResults=${max_results}&includeAttributes=true`,
              ),
              body: { qlQuery: scoped },
            });
            const names = new Map<string, string>(
              ((data?.objectTypeAttributes ?? []) as Obj[]).map((a) => [String(a.id), a.name]),
            );
            const values = (data?.values ?? []) as Obj[];
            const objects = values
              .filter((o) => allowedSchema(schemas, o.objectType?.objectSchemaId))
              .map((o) => shownObject(o, names));
            // isLast comes as a string in Atlassian's own example.
            const last = data?.isLast === true || data?.isLast === "true";
            return json({
              objects,
              total: typeof data?.total === "number" ? data.total : null,
              more: !last && values.length > 0,
              next_start_at: last ? null : start_at + values.length,
            });
          }),
      );

    if (allowed.has("jira_assets_get_object"))
      server.registerTool(
        "jira_assets_get_object",
        {
          description:
            "Read one Jira Assets object by id: its key, label, object type, schema and every " +
            "attribute by name (a referenced object, status, user or group by its name).",
          inputSchema: z.object({ object_id: objectId }),
          annotations: readOnly,
        },
        ({ object_id }) =>
          audited(ctx, "jira_assets_get_object", object_id, async (mark) =>
            json(shownObject(await fetchObject(ctx, mark, object_id))),
          ),
      );

    if (allowed.has("jira_assets_list_attributes"))
      server.registerTool(
        "jira_assets_list_attributes",
        {
          description:
            "List the attributes of a Jira Assets object type (its object_type_id from a search " +
            "or object): name, kind (Text, Date, Select, Status, User, Object reference...), " +
            "whether it can be edited, is required or takes several values, and the options of " +
            "a Select.",
          inputSchema: z.object({
            object_type_id: z.string().max(20).describe("Object type id, e.g. 23."),
          }),
          annotations: readOnly,
        },
        ({ object_type_id }) =>
          audited(ctx, "jira_assets_list_attributes", object_type_id, async (mark) =>
            json(
              (await typeAttributes(ctx, mark, object_type_id)).map((a) => ({
                id: String(a.id),
                name: a.name,
                kind: attributeKind(a),
                editable: a.editable === true && a.system !== true,
                required: (a.minimumCardinality ?? 0) > 0,
                multiple: a.maximumCardinality !== 1,
                options:
                  typeof a.options === "string" && a.options
                    ? a.options.split(",").map((o: string) => o.trim())
                    : null,
                description: a.description ?? null,
              })),
            ),
          ),
      );

    if (allowed.has("jira_assets_update_object"))
      server.registerTool(
        "jira_assets_update_object",
        {
          description:
            "Change named attributes of one existing Jira Assets object; every other attribute " +
            "is left as it is. Values are text: dates as YYYY-MM-DD, a Select its option, a " +
            "Status its name, an Object reference the object's key (e.g. ITSM-88), a User the " +
            "Atlassian account id. A list sets several values; null clears the attribute. " +
            "Answers with the object read back after the change.",
          inputSchema: z.object({
            object_id: objectId,
            attributes: z
              .record(
                z.string().min(1).max(200),
                z.union([z.string().max(10000), z.array(z.string().max(10000)).max(50), z.null()]),
              )
              .describe('Attribute name to new value, e.g. {"Status": "In use"}.'),
          }),
          annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
        },
        ({ object_id, attributes }) =>
          audited(
            ctx,
            "jira_assets_update_object",
            `${object_id} (${Object.keys(attributes).join(", ")})`,
            async (mark) => {
              const changes = Object.entries(attributes);
              if (changes.length === 0) throw new Denied("Name at least one attribute to change");
              if (changes.length > 20) throw new Denied("Change at most 20 attributes at a time");
              const object = await fetchObject(ctx, mark, object_id);
              const typeId = String(object.objectType?.id ?? "");
              const schemaId = String(object.objectType?.objectSchemaId ?? "");
              const defs = await typeAttributes(ctx, mark, typeId, true);
              const byName = new Map(defs.map((a) => [String(a.name).toLowerCase(), a]));
              const editable = defs.filter((a) => a.editable === true && a.system !== true);
              let statuses: Obj[] | null = null;
              const body = [];
              for (const [name, value] of changes) {
                const def = byName.get(name.trim().toLowerCase());
                if (!def)
                  throw new Denied(
                    `This object type has no attribute ${name}. Editable ones: ${editable
                      .map((a) => a.name)
                      .join(", ")}`,
                  );
                if (def.editable !== true || def.system === true)
                  throw new Denied(`${def.name} cannot be edited`);
                let values = value === null ? [] : Array.isArray(value) ? value : [value];
                if (values.length > 1 && def.maximumCardinality === 1)
                  throw new Denied(`${def.name} takes one value`);
                if (def.type === STATUS && values.length) {
                  if (!isSchemaId(schemaId)) throw new Error("Object has no schema id");
                  statuses ??= (await ok(ctx, mark, {
                    method: "GET",
                    path: path(`config/statustype?objectSchemaId=${schemaId}`),
                  })) as Obj[];
                  // A Status attribute may allow only some statuses (typeValueMulti, ids).
                  const allowedIds = ((def.typeValueMulti ?? []) as unknown[]).map(String);
                  const usable = statuses.filter(
                    (s) => !allowedIds.length || allowedIds.includes(String(s.id)),
                  );
                  values = values.map((v) => {
                    const want = v.trim().toLowerCase();
                    const found = usable.find(
                      (s) =>
                        String(s.id) === want ||
                        String(s.name ?? "").toLowerCase() === want ||
                        String(s.displayName ?? "").toLowerCase() === want,
                    );
                    if (!found)
                      throw new Denied(
                        `${def.name} has no status ${v}. Use one of: ${usable
                          .map((s) => s.name)
                          .join(", ")}`,
                      );
                    return String(found.id);
                  });
                }
                body.push({
                  objectTypeAttributeId: String(def.id),
                  objectAttributeValues: values.map((v) => ({ value: v })),
                });
              }
              await ok(ctx, mark, {
                method: "PUT",
                path: path(`object/${object_id}`),
                body: { objectTypeId: typeId, attributes: body },
              });
              const after = await fetchObject(ctx, mark, object_id);
              return json({
                updated: changes.map(([name]) => byName.get(name.trim().toLowerCase())!.name),
                object: shownObject(after),
              });
            },
          ),
      );
  },
};
