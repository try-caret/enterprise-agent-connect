import { isDeepStrictEqual } from "node:util";
import { z } from "zod";
import { PARAMS, routes, schemas, type Answer, type RouteSpec } from "./api.js";

// openapi.json (OpenAPI 3.1, whose schemas are JSON Schema 2020-12, what zod writes) from the
// route list in api.ts. Named shapes (`.meta({ id })`) become components; the rest stay inline.

type Json = Record<string, unknown>;

export const ADMIN_V1 = "/agent-connect/admin/v1";

function convert(schema: z.ZodType, io: "input" | "output", components: Json): Json {
  const {
    $schema: _,
    $defs,
    ...json
  } = z.toJSONSchema(schema, {
    io,
    target: "draft-2020-12",
  }) as Json & { $defs?: Record<string, Json> };
  const ref = (value: unknown): unknown =>
    Array.isArray(value)
      ? value.map(ref)
      : value && typeof value === "object"
        ? Object.fromEntries(
            Object.entries(value).map(([k, v]) => [
              k,
              k === "$ref" && typeof v === "string"
                ? v.replace("#/$defs/", "#/components/schemas/")
                : ref(v),
            ]),
          )
        : value;
  for (const [id, def] of Object.entries($defs ?? {})) {
    const next = ref(def);
    // One id, one shape: a body and an answer naming the same schema must agree.
    if (id in components && !isDeepStrictEqual(components[id], next))
      throw new Error(`Schema ${id} differs between its uses`);
    components[id] = next;
  }
  return ref(json) as Json;
}

function answer(status: number, a: Answer, components: Json): Json {
  const description = status < 300 ? "OK" : "Refused";
  if (a === null) return { description };
  if ("file" in a)
    return { description, content: { [a.file]: { schema: { type: "string", format: "binary" } } } };
  return {
    description,
    content: { "application/json": { schema: convert(a, "output", components) } },
  };
}

export function openapi(): Json {
  const components: Json = {};
  const error = () => answer(400, schemas.Error, components);
  const paths: Record<string, Json> = {};
  for (const [operationId, spec] of Object.entries(routes) as [string, RouteSpec][]) {
    const parameters = [
      ...[...spec.path.matchAll(/\{(\w+)\}/g)].map(([, p]) => ({
        name: p,
        in: "path",
        required: true,
        schema: { type: "string", pattern: `^${PARAMS[p as keyof typeof PARAMS]}$` },
      })),
      ...Object.entries(spec.query ?? {}).map(([q, description]) => ({
        name: q,
        in: "query",
        required: false,
        description,
        schema: { type: "string" },
      })),
    ];
    const responses: Json = {};
    for (const [status, a] of Object.entries(spec.responses))
      responses[status] = answer(Number(status), a, components);
    // Every route: bad input, a wrong token, an unknown organization or route, a refusal by
    // the app being connected.
    for (const status of ["400", "401", "404", "422"]) responses[status] ??= error();
    (paths[spec.path] ??= {})[spec.method.toLowerCase()] = {
      operationId,
      summary: spec.summary,
      ...(spec.teams ? { tags: ["teams"] } : spec.events ? { tags: ["events"] } : {}),
      ...(parameters.length ? { parameters } : {}),
      ...(spec.body
        ? {
            requestBody: {
              required: !spec.body.safeParse(undefined).success,
              content: { "application/json": { schema: convert(spec.body, "input", components) } },
            },
          }
        : {}),
      responses,
    };
  }
  return {
    openapi: "3.1.0",
    info: {
      title: "Agent Connect admin API",
      version: "1",
      description:
        "Server to server only: a bearer token (AGENT_CONNECT_ADMIN_TOKEN), and requests " +
        "carrying X-Forwarded-For are refused. Routes tagged teams are served only where " +
        "Teams is set up, routes tagged events only where app events are.",
    },
    servers: [{ url: ADMIN_V1 }],
    security: [{ bearer: [] }],
    paths,
    components: {
      schemas: Object.fromEntries(
        Object.entries(components).sort(([a], [b]) => a.localeCompare(b)),
      ),
      securitySchemes: { bearer: { type: "http", scheme: "bearer" } },
    },
  };
}
