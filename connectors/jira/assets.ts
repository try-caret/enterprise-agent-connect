import { z } from "zod";
import { isAssetsId, scopeAql } from "./aql.js";
import { JqlRejected } from "./jql.js";
import { answerOf } from "./answers.js";
import { ASSETS_PATH, WORKSPACE_PROBLEM } from "./vendor.js";
import {
  audited,
  Denied,
  json,
  listLimit,
  VendorFailed,
  type ToolContext,
  type Toolset,
} from "../../packages/connector-kit/src/common.js";

// Jira Assets (Jira Service Management's CMDB): search with AQL, read an object, list a type's
// attributes, change named attributes of one existing object. No create or delete. Generic: no
// object type is known here. The `asset_schemas` limit holds every call (AQL is confined to the
// schemas and every object is checked on objectType.objectSchemaId). Shapes from
// developer.atlassian.com/cloud/assets/swagger.v3.json, read November 7, 2026.

type Mark = (status: number) => void;
type Obj = Record<string, any>;

const path = (rest: string) => `${ASSETS_PATH}${rest}`;
// An object type attribute's `type` (ObjectTypeAttribute in the OpenAPI): 0 is a plain value
// whose kind is `defaultType.name` (Text, Date, Select...).
const REFERENCE = 1;
const USER = 2;
const GROUP = 4;
const STATUS = 7;
const SELECT = 10; // defaultType.id of a plain Select value
const ATTRIBUTE_TYPES: Record<number, string> = {
  [REFERENCE]: "Object reference",
  [USER]: "User",
  [GROUP]: "Group",
  [STATUS]: "Status",
};

// `type` is optional in the OpenAPI (its own example leaves it out and gives a reference
// attribute only referenceObjectTypeId), so without it a reference is known by those fields; a
// `type` that is there is trusted, so empty reference fields on other attributes change nothing.
const isReference = (def: Obj | undefined) =>
  def?.type === REFERENCE ||
  (def?.type === undefined && Boolean(def?.referenceObjectTypeId || def?.referenceType));
// A plain value (Text, Date, Select...): `type` 0, or no `type` and not a reference.
const isPlain = (def: Obj) => (def.type ?? 0) === 0 && !isReference(def);

const attributeId = (a: Obj) =>
  String(a.objectTypeAttributeId ?? a.objectTypeAttribute?.id ?? "");

const answer = answerOf("Jira Assets");
// Assets documents 401, not 403, for a missing seat, scope or role (seen live for a scope), and
// the framework adds a capability's `forbidden` hint only to a 403, so it is added to a 401
// here: the read capability's for reads, the update capability's for the PUT.
async function ok(ctx: ToolContext, mark: Mark, req: Parameters<ToolContext["call"]>[0]) {
  try {
    return await answer(ctx, mark, req);
  } catch (error) {
    if (!(error instanceof VendorFailed)) throw error;
    // The workspace lookup's refusal names its own fix (seat, scope): hintTool "" keeps the
    // framework from adding the capability's hint to a 403.
    if (error.message.includes(WORKSPACE_PROBLEM))
      throw new VendorFailed(error.status, error.message, "");
    if (error.status !== 401) throw error;
    const tool = req.method === "PUT" ? "jira_assets_update_object" : "jira_assets_search";
    const hint = ctx.forbidden?.[tool];
    throw hint ? new VendorFailed(401, `${error.message.replace(/\.$/, "")}. ${hint}`) : error;
  }
}

const canEdit = (def: Obj) => def.editable === true && def.system !== true;

function allowedSchema(schemas: string[] | null, id: unknown): boolean {
  return schemas === null || (id !== undefined && id !== null && schemas.includes(String(id)));
}

// One value as an agent reads it: a reference, status, user or group by its name (a referenced
// object outside the agent's schemas is not named), any other value as stored (ISO dates).
const HIDDEN_REFERENCE = "(an object in a schema this agent may not use)";
// Text longer than this is shortened, as other connectors do, so 50 objects stay readable.
const MAX_TEXT = 2000;
const shortened = (v: unknown) =>
  typeof v === "string" && v.length > MAX_TEXT ? `${v.slice(0, MAX_TEXT)}... (shortened)` : v;
function shownValue(v: Obj, schemas: string[] | null, reference: boolean): unknown {
  // `referencedObject` is optional in the OpenAPI, so a reference without it is hidden too.
  if (reference || v?.referencedObject || v?.referencedType === true) {
    if (!allowedSchema(schemas, v?.referencedObject?.objectType?.objectSchemaId))
      return HIDDEN_REFERENCE;
    return v.displayValue ?? v.referencedObject?.label ?? null;
  }
  if (v?.status) return v.status.name ?? v.displayValue ?? null;
  // The account id too: update takes a User by it, and display names are not unique.
  if (v?.user)
    return {
      name: v.user.displayName ?? v.displayValue ?? null,
      account_id: v.user.key ?? v.value ?? null,
    };
  if (v?.group) return v.group.name ?? v.displayValue ?? null;
  return shortened(v?.value ?? v?.displayValue ?? null);
}

// An object with its attributes by name. `defs`: attribute id to definition, for answers (AQL)
// whose attributes carry only their objectTypeAttributeId.
function shownObject(o: Obj, schemas: string[] | null, defs: Map<string, Obj> = new Map()) {
  const attributes: Record<string, unknown> = {};
  for (const a of (o.attributes ?? []) as Obj[]) {
    const id = attributeId(a);
    const def = a.objectTypeAttribute ?? defs.get(id);
    const name = def?.name ?? `attribute ${id}`;
    const reference = isReference(def);
    const values = ((a.objectAttributeValues ?? []) as Obj[]).map((v) =>
      shownValue(v, schemas, reference),
    );
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

// Under a schema limit a missing id gets the same refusal as one in another schema, so ids
// (sequential in Assets) cannot be walked to count what other schemas hold.
async function limitedGet(ctx: ToolContext, mark: Mark, rest: string, refusal: string) {
  try {
    return await ok(ctx, mark, { method: "GET", path: path(rest) });
  } catch (error) {
    // A missing workspace is the site's, not the object's: it keeps its own message.
    if (
      error instanceof VendorFailed &&
      error.status === 404 &&
      !error.message.includes(WORKSPACE_PROBLEM) &&
      listLimit(ctx.limits, "asset_schemas") !== null
    )
      throw new Denied(refusal);
    throw error;
  }
}

async function fetchObject(ctx: ToolContext, mark: Mark, id: string): Promise<Obj> {
  if (!isAssetsId(id)) throw new Denied("Give an Assets object id, a number such as 88");
  const refusal = `Assets object ${id} is not in an object schema this agent may use`;
  const object = await limitedGet(ctx, mark, `object/${id}`, refusal);
  if (!allowedSchema(listLimit(ctx.limits, "asset_schemas"), object?.objectType?.objectSchemaId))
    throw new Denied(refusal);
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
  if (!isAssetsId(typeId)) throw new Denied("Give an object type id, a number such as 23");
  const refusal = `Object type ${typeId} is not in an object schema this agent may use`;
  const rest = `objecttype/${typeId}/attributes`;
  const attrs = ((await limitedGet(ctx, mark, rest, refusal)) ?? []) as Obj[];
  const schemas = listLimit(ctx.limits, "asset_schemas");
  // An empty list names no schema to check, so under a limit it is refused like another's.
  if (
    !checked &&
    schemas !== null &&
    (attrs.length === 0 || attrs.some((a) => !allowedSchema(schemas, a.objectType?.objectSchemaId)))
  )
    throw new Denied(refusal);
  return attrs;
}

// A Status value's id. An id the attribute allows (`typeValueMulti`) is taken as given; anything
// else is looked up by name on an object of the type that already holds it, since the status
// list needs read:cmdb-config:jira, which service accounts are not offered (401, checked
// November 8, 2026). A status no object holds yet is refused without listing ids to guess from.
async function statusId(
  ctx: ToolContext,
  mark: Mark,
  schemaId: string,
  typeId: string,
  def: Obj,
  value: string,
): Promise<string> {
  const allowed = ((def.typeValueMulti ?? []) as unknown[]).map(String);
  const want = value.trim();
  if (isAssetsId(want) && allowed.includes(want)) return want;
  if (/["\\]/.test(want) || /["\\]/.test(String(def.name)))
    throw new Denied(`${def.name}'s status ${value} cannot be looked up by name`);
  const found = await statusByName(ctx, mark, schemaId, typeId, def, want);
  if (found) {
    if (allowed.length && !allowed.includes(found))
      throw new Denied(`${def.name} does not allow status ${value}`);
    return found;
  }
  // No allowed list is declared, so digits no name matched are taken as an id.
  if (isAssetsId(want) && !allowed.length) return want;
  throw new Denied(
    `${def.name} has no status ${value} on any object of this type, so it cannot be looked up ` +
      "(Atlassian does not let a service account read the status list). Ask a person to set " +
      "it in Assets, or for its status id.",
  );
}

// The id of a status named `want` on one object of the type that holds it, or null.
async function statusByName(
  ctx: ToolContext,
  mark: Mark,
  schemaId: string,
  typeId: string,
  def: Obj,
  want: string,
): Promise<string | null> {
  const data = await ok(ctx, mark, {
    method: "POST",
    path: path("object/aql?startAt=0&maxResults=1&includeAttributes=true"),
    body: {
      qlQuery:
        `objectSchemaId = ${schemaId} AND objectTypeId = ${typeId} ` +
        `AND "${def.name}" = "${want}"`,
    },
  });
  for (const o of (data?.values ?? []) as Obj[])
    for (const a of (o.attributes ?? []) as Obj[])
      if (attributeId(a) === String(def.id))
        for (const v of (a.objectAttributeValues ?? []) as Obj[])
          if (
            v?.status?.id !== undefined &&
            String(v.status.name ?? "").toLowerCase() === want.toLowerCase()
          )
            return String(v.status.id);
  return null;
}

// A vendor's or a refusal's message is written for people; anything else stays in the log.
const shownError = (error: unknown) =>
  error instanceof VendorFailed || error instanceof Denied ? error.message : "";

// A Select's options. Assets keeps them as one comma-separated text and a comma always
// separates two (tried live November 8, 2026: with options "Boston, MA,London", "Boston, MA" is
// refused and "MA" is taken).
const selectOptions = (def: Obj): string[] =>
  typeof def.options === "string" && def.options
    ? def.options.split(",").map((o: string) => o.trim()).filter(Boolean)
    : [];

// A Select value as one of its options, in the option's own spelling.
function selectOption(def: Obj, value: string): string | null {
  const want = value.trim().toLowerCase();
  return selectOptions(def).find((o) => o.toLowerCase() === want) ?? null;
}

const attributeKind = (a: Obj) =>
  isReference(a)
    ? ATTRIBUTE_TYPES[REFERENCE]
    : isPlain(a)
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
            "Attribute names with spaces go in double quotes; ORDER BY goes at the end. " +
            "Each object comes with its attributes by name. Results are limited to the object " +
            "schemas this agent may use. An agent held to schemas may not use HAVING, " +
            "inboundReferences, outboundReferences or dot notation (Owner.Name); one held to " +
            "schemas or projects may not use connectedTickets, single quotes, or a backslash " +
            'except before ". ' +
            "Find attribute names with jira_assets_list_attributes.",
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
              scoped = scopeAql(aql, schemas, listLimit(ctx.limits, "projects") !== null);
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
            const defs = new Map<string, Obj>(
              ((data?.objectTypeAttributes ?? []) as Obj[]).map((a) => [String(a.id), a]),
            );
            const values = (data?.values ?? []) as Obj[];
            const objects = values
              .filter((o) => allowedSchema(schemas, o.objectType?.objectSchemaId))
              .map((o) => shownObject(o, schemas, defs));
            // isLast comes as a string in Atlassian's own example.
            const last = data?.isLast === true || data?.isLast === "true";
            return json({
              objects,
              total: typeof data?.total === "number" ? data.total : null,
              more: !last && values.length > 0,
              next_start_at: !last && values.length > 0 ? start_at + values.length : null,
            });
          }),
      );

    if (allowed.has("jira_assets_get_object"))
      server.registerTool(
        "jira_assets_get_object",
        {
          description:
            "Read one Jira Assets object by id: its key, label, object type, schema and every " +
            "attribute by name (a referenced object, status or group by its name, a user as " +
            "{name, account_id}).",
          inputSchema: z.object({ object_id: objectId }),
          annotations: readOnly,
        },
        ({ object_id }) =>
          audited(ctx, "jira_assets_get_object", object_id, async (mark) =>
            json(
              shownObject(
                await fetchObject(ctx, mark, object_id),
                listLimit(ctx.limits, "asset_schemas"),
              ),
            ),
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
                editable: canEdit(a),
                required: (a.minimumCardinality ?? 0) > 0,
                multiple: a.maximumCardinality !== 1,
                options: selectOptions(a).length ? selectOptions(a) : null,
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
            "Change named attributes (at most 20) of one existing Jira Assets object; every " +
            "other attribute is left as it is, and the old values are not kept. Values are " +
            "text: a Date as YYYY-MM-DD, a DateTime as ISO 8601 with a time, a Select its " +
            "option, a User the Atlassian account id, a Status its id or its name (a name only " +
            "when some object of the type already has that status). A list sets several " +
            "values; only null clears, and never a required attribute. Object reference and " +
            "read-only attributes are refused. " +
            "Answers with the object as saved.",
          inputSchema: z.object({
            object_id: objectId,
            attributes: z
              .record(
                z.string().min(1).max(200),
                z.union([
                  z.string().max(10000),
                  z.array(z.string().max(10000)).min(1).max(50),
                  z.null(),
                ]),
              )
              .refine((a) => Object.keys(a).length <= 20, "Change at most 20 attributes at a time")
              .describe('Attribute name to new value, e.g. {"Status": "In use"}.'),
          }),
          annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
        },
        ({ object_id, attributes }) =>
          audited(
            ctx,
            "jira_assets_update_object",
            `${object_id} (${Object.keys(attributes).join(", ")})`,
            async (mark) => {
              const changes = Object.entries(attributes);
              if (changes.length === 0) throw new Denied("Name at least one attribute to change");
              const object = await fetchObject(ctx, mark, object_id);
              const typeId = String(object.objectType?.id ?? "");
              const schemaId = String(object.objectType?.objectSchemaId ?? "");
              const defs = await typeAttributes(ctx, mark, typeId, true);
              const byName = new Map(defs.map((a) => [String(a.name).toLowerCase(), a]));
              const editable = defs.filter(canEdit);
              const body = [];
              const named = new Set<string>();
              for (const [name, value] of changes) {
                const def = byName.get(name.trim().toLowerCase());
                if (!def)
                  throw new Denied(
                    `This object type has no attribute ${name}. Editable ones: ${editable
                      .map((a) => a.name)
                      .join(", ")}`,
                  );
                if (named.has(String(def.id))) throw new Denied(`${def.name} is named twice`);
                named.add(String(def.id));
                if (!canEdit(def))
                  throw new Denied(`${def.name} cannot be edited`);
                if (isReference(def))
                  throw new Denied(
                    `${def.name} links to other objects, which this tool does not change`,
                  );
                let values = value === null ? [] : Array.isArray(value) ? value : [value];
                if (!values.length && (def.minimumCardinality ?? 0) > 0)
                  throw new Denied(`${def.name} is required and cannot be cleared`);
                if (values.length > 1 && def.maximumCardinality === 1)
                  throw new Denied(`${def.name} takes one value`);
                if (def.type === STATUS && values.length) {
                  if (!isAssetsId(schemaId)) throw new Error("Object has no schema id");
                  const resolved = [];
                  for (const v of values)
                    resolved.push(await statusId(ctx, mark, schemaId, typeId, def, v));
                  values = resolved;
                }
                if (isPlain(def) && def.defaultType?.id === SELECT && values.length) {
                  values = values.map((v) => {
                    const found = selectOption(def, v);
                    if (!found)
                      throw new Denied(
                        `${def.name} has no option ${v}. ` +
                          `Use one of: ${selectOptions(def).join(" | ")}`,
                      );
                    return found;
                  });
                }
                body.push({
                  objectTypeAttributeId: String(def.id),
                  objectAttributeValues: values.map((v) => ({ value: v })),
                });
              }
              const saved = await ok(ctx, mark, {
                method: "PUT",
                path: path(`object/${object_id}`),
                body: { objectTypeId: typeId, attributes: body },
              });
              const updated = changes.map(([name]) => byName.get(name.trim().toLowerCase())!.name);
              const schemas = listLimit(ctx.limits, "asset_schemas");
              // The PUT answers the updated object (OpenAPI); read again only when it does not.
              if (
                String(saved?.id ?? "") === object_id &&
                saved?.objectType?.objectSchemaId !== undefined &&
                allowedSchema(schemas, saved.objectType.objectSchemaId)
              )
                return json({ updated, object: shownObject(saved, schemas) });
              // The change is made: a failed read-back must not tell the agent it failed (it
              // would retry), so it is reported beside the result and the row stays the write's.
              try {
                const after = await fetchObject(ctx, mark, object_id);
                return json({
                  updated,
                  object: shownObject(after, listLimit(ctx.limits, "asset_schemas")),
                });
              } catch (error) {
                if (!shownError(error))
                  console.error(
                    "[agent-connect] jira_assets_update_object read-back failed",
                    error,
                  );
                mark(200);
                return json({
                  updated,
                  object: null,
                  warning:
                    "The change was saved, but reading the object back failed: do not repeat " +
                    `it. ${shownError(error)}`.trim(),
                });
              }
            },
          ),
      );
  },
};
