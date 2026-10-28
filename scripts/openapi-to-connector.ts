import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname } from "node:path";
import { pathToFileURL } from "node:url";
import { z } from "zod";
import { connectorSchema } from "../packages/connector-kit/src/connector.ts";

// Drafts a connector.json for the http connector from an OpenAPI 3.x spec (JSON):
//   npx tsx scripts/openapi-to-connector.ts <spec file or URL> --id <app id>
//     (--operations <id,id> | --tags <tag,tag>) [--out <file>]
// The draft loads, but a person finishes it: copy, capabilities, limits, answer limits.

type Json = Record<string, unknown>;
type Options = { id: string; operations?: string[]; tags?: string[]; source?: string };

const METHODS = ["get", "post", "put", "patch", "delete"] as const;
const SCALARS = new Set(["string", "number", "integer", "boolean"]);

const snake = (s: string) =>
  s
    .replace(/([a-z0-9])([A-Z])/g, "$1_$2")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^[^a-z]+|_+$/g, "");

const firstLine = (s: unknown) =>
  typeof s === "string"
    ? s
        .trim()
        .split(/\n\s*\n|\r?\n/)[0]
        .slice(0, 300)
    : "";

export function draftConnector(spec: Json, options: Options) {
  const notes: string[] = [];
  if (typeof spec.openapi !== "string" || !spec.openapi.startsWith("3."))
    throw new Error("This takes an OpenAPI 3.x spec (its openapi field starts with 3.)");

  // Follows local $refs (#/components/...); other specs' refs are left out.
  const resolve = (node: unknown, seen = 0): Json => {
    if (!node || typeof node !== "object") return {};
    const ref = (node as Json).$ref;
    if (typeof ref !== "string") return node as Json;
    if (!ref.startsWith("#/") || seen > 20) {
      notes.push(`Left out ${ref}: only references inside the spec are followed`);
      return {};
    }
    const target = ref
      .slice(2)
      .split("/")
      .reduce<unknown>(
        (at, key) => (at as Json | undefined)?.[key.replace(/~1/g, "/").replace(/~0/g, "~")],
        spec,
      );
    return resolve(target, seen + 1);
  };

  const objectOf = (raw: unknown): { properties: Json; required: string[] } => {
    const s = resolve(raw);
    const parts = Array.isArray(s.allOf) ? s.allOf.map((p) => objectOf(p)) : [];
    return {
      properties: Object.assign({}, ...parts.map((p) => p.properties), s.properties ?? {}) as Json,
      required: [
        ...parts.flatMap((p) => p.required),
        ...(Array.isArray(s.required) ? s.required : []),
      ],
    };
  };

  const typeOf = (s: Json) =>
    (Array.isArray(s.type) ? s.type.find((t) => t !== "null") : s.type) as string | undefined;

  // The http connector's param fields for a schema, or null for one it cannot take (an object).
  const shape = (raw: unknown) => {
    const s = resolve(raw);
    const type = typeOf(s);
    const items = type === "array" ? resolve(s.items) : null;
    const scalar = items ? typeOf(items) : type;
    if (!scalar || !SCALARS.has(scalar)) return null;
    const fits = (v: unknown) =>
      scalar === "integer"
        ? Number.isInteger(v)
        : typeof v === (scalar === "integer" ? "number" : scalar);
    const values = (items ?? s).enum;
    const allowed =
      Array.isArray(values) && values.length && values.every(fits) ? values : undefined;
    const ok = (v: unknown) => fits(v) && (!allowed || allowed.includes(v));
    const noDefault =
      s.default === undefined ||
      (items ? !Array.isArray(s.default) || !s.default.every(ok) : !ok(s.default));
    return {
      type: items ? "array" : scalar,
      ...(items ? { items: scalar } : {}),
      ...(allowed ? { enum: allowed } : {}),
      ...(noDefault ? {} : { default: s.default }),
    } as Json;
  };

  const scalarFields = (raw: unknown) =>
    Object.entries(objectOf(raw).properties)
      .filter(([, p]) => SCALARS.has(typeOf(resolve(p)) ?? ""))
      .map(([k]) => k);

  const resultOf = (op: Json) => {
    const responses = resolve(op.responses);
    const ok = Object.keys(responses).find((code) => /^2/.test(code));
    const schema =
      ok && resolve(resolve(resolve(responses[ok]).content)["application/json"]).schema;
    if (!schema) return undefined;
    const s = resolve(schema);
    if (typeOf(s) === "array") {
      const fields = scalarFields(s.items);
      return fields.length ? { fields } : undefined;
    }
    const props = objectOf(s).properties;
    // A list of records inside an object is taken as the page, e.g. {issues: [...], next}.
    const lists = Object.entries(props)
      .filter(([, p]) => typeOf(resolve(p)) === "array")
      .map(([k, p]) => [k, scalarFields(resolve(p).items)] as const)
      .filter(([, fields]) => fields.length);
    if (lists.length === 1) return { items: lists[0][0], fields: lists[0][1] };
    const fields = scalarFields(s);
    return fields.length ? { fields } : undefined;
  };

  const app = snake(options.id);
  const info = resolve(spec.info);
  const title =
    typeof info.title === "string" && info.title.trim() ? info.title.trim() : options.id;

  const picked: { path: string; method: (typeof METHODS)[number]; op: Json; shared: unknown[] }[] =
    [];
  for (const [path, rawItem] of Object.entries(resolve(spec.paths))) {
    const item = resolve(rawItem);
    for (const method of METHODS) {
      const op = item[method] as Json | undefined;
      if (!op) continue;
      const tags = Array.isArray(op.tags) ? (op.tags as string[]) : [];
      if (
        options.operations?.includes(op.operationId as string) ||
        options.tags?.some((t) => tags.includes(t))
      )
        picked.push({ path, method, op, shared: (item.parameters as unknown[]) ?? [] });
    }
  }
  const missing = (options.operations ?? []).filter(
    (id) => !picked.some((p) => p.op.operationId === id),
  );
  if (missing.length) throw new Error(`No operation has the id ${missing.join(", ")}`);
  if (!picked.length) throw new Error("No operation matched: name operations or tags");

  const tools = picked.map(({ path, method, op, shared }) => {
    const opId = typeof op.operationId === "string" ? op.operationId : `${method} ${path}`;
    const base = snake(opId);
    const name = base.startsWith(`${app}_`) ? base : `${app}_${base}`;
    const params: Record<string, Json> = {};
    const add = (key: string, param: Json, why: string) => {
      if (params[key]) notes.push(`${name}: left out ${why}, its name is taken by another param`);
      else params[key] = param;
    };

    const byPlace = new Map<string, Json>();
    for (const raw of [...shared, ...((op.parameters as unknown[]) ?? [])]) {
      const p = resolve(raw);
      byPlace.set(`${p.in}:${p.name}`, p);
    }
    let template = path;
    for (const p of byPlace.values()) {
      const wire = String(p.name);
      if (p.in !== "path" && p.in !== "query") {
        notes.push(
          `${name}: left out the ${p.in} param ${wire} (a fixed header goes in http.headers)`,
        );
        continue;
      }
      const fields = shape(p.schema);
      if (!fields) {
        notes.push(`${name}: left out ${wire}, whose type the http connector cannot take`);
        continue;
      }
      const key = snake(wire) || "param";
      if (p.in === "path") template = template.replaceAll(`{${wire}}`, `{${key}}`);
      if (p.in === "path") delete fields.default;
      const style = p.style ?? "form";
      const join =
        fields.type === "array" && p.in === "query"
          ? style === "spaceDelimited"
            ? " "
            : style === "pipeDelimited"
              ? "|"
              : p.explode === false
                ? ","
                : undefined
          : undefined;
      add(
        key,
        {
          ...fields,
          description: firstLine(p.description) || firstLine(resolve(p.schema).description) || wire,
          in: p.in,
          ...(key === wire || p.in === "path" ? {} : { name: wire }),
          ...(p.in === "path" || p.required === true || fields.default !== undefined
            ? {}
            : { required: false }),
          ...(join ? { join } : {}),
        },
        `query param ${wire}`,
      );
    }

    const body = resolve(op.requestBody);
    const content = resolve(body.content);
    if (method === "get" && op.requestBody)
      notes.push(`${name}: left out its request body (a GET)`);
    else if (op.requestBody && !content["application/json"])
      notes.push(`${name}: left out its request body, which is not JSON`);
    else if (op.requestBody) {
      const { properties, required } = objectOf(resolve(content["application/json"]).schema);
      if (!Object.keys(properties).length)
        notes.push(`${name}: its request body has no listed fields, so it sends none`);
      for (const [wire, raw] of Object.entries(properties)) {
        const prop = resolve(raw);
        if (prop.readOnly) continue;
        const fields = shape(prop);
        if (!fields) {
          notes.push(`${name}: left out body field ${wire}, an object or a type not supported`);
          continue;
        }
        const key = snake(wire) || "field";
        add(
          key,
          {
            ...fields,
            description: firstLine(prop.description) || wire,
            in: "body",
            ...(key === wire ? {} : { name: wire }),
            ...(required.includes(wire) || fields.default !== undefined ? {} : { required: false }),
          },
          `body field ${wire}`,
        );
      }
    }

    const result = method === "get" ? resultOf(op) : undefined;
    if (result)
      notes.push(`${name}: its answer is trimmed to fields guessed from the spec; check them`);
    return {
      name,
      description:
        firstLine(op.summary) || firstLine(op.description) || `${method.toUpperCase()} ${path}`,
      method: method.toUpperCase(),
      path: template,
      params,
      ...(result ? { result } : {}),
    };
  });

  const baseUrl = (() => {
    const server = resolve((spec.servers as unknown[] | undefined)?.[0]);
    let url = typeof server.url === "string" ? server.url : "";
    for (const [name, v] of Object.entries(resolve(server.variables)))
      url = url.replaceAll(`{${name}}`, String(resolve(v).default ?? ""));
    if (!url || !URL.canParse(url, options.source)) {
      notes.push("The spec names no server address: set http.baseUrl");
      return "https://app.example.com";
    }
    return new URL(url, options.source).href.replace(/\/+$/, "");
  })();
  const auth = authOf();
  function authOf(): { auth: Json; fields: Json[]; credential: string } {
    const requirement = [
      ...((picked[0].op.security as Json[] | undefined) ?? (spec.security as Json[]) ?? []),
    ][0];
    const schemeName = requirement && Object.keys(requirement)[0];
    const scheme = schemeName
      ? resolve(resolve(resolve(spec.components).securitySchemes)[schemeName])
      : {};
    const secret = (id: string, label: string) => ({ id, label, secret: true });
    const plain = (id: string, label: string) => ({ id, label, secret: false });
    if (scheme.type === "apiKey" && scheme.in === "header") {
      if (String(scheme.name).toLowerCase() === "authorization")
        notes.push("The key goes in Authorization: add auth.prefix if a word comes before it");
      return {
        auth: { type: "header", header: scheme.name, value: "api_key" },
        fields: [secret("api_key", "API key")],
        credential: "an API key",
      };
    }
    if (scheme.type === "apiKey" && scheme.in === "query")
      return {
        auth: { type: "query", param: scheme.name, value: "api_key" },
        fields: [secret("api_key", "API key")],
        credential: "an API key",
      };
    if (scheme.type === "http" && String(scheme.scheme).toLowerCase() === "basic")
      return {
        auth: { type: "basic", username: "username", password: "password" },
        fields: [plain("username", "Username"), secret("password", "Password")],
        credential: "a username and password",
      };
    const flow = resolve(resolve(scheme.flows).clientCredentials);
    if (scheme.type === "oauth2" && typeof flow.tokenUrl === "string")
      return {
        auth: {
          type: "oauth2",
          tokenUrl: new URL(flow.tokenUrl, baseUrl).href,
          clientId: "client_id",
          clientSecret: "client_secret",
        },
        fields: [plain("client_id", "Client ID"), secret("client_secret", "Client secret")],
        credential: "a client ID and secret",
      };
    if (!(scheme.type === "http" && String(scheme.scheme).toLowerCase() === "bearer"))
      notes.push(
        schemeName
          ? `Sign-in ${schemeName} (${String(scheme.type)}) does not map to one the http connector has; the draft uses a token, which may work if the app issues tokens for services`
          : "The spec names no sign-in; the draft uses a token",
      );
    return {
      auth: { type: "bearer", token: "token" },
      fields: [secret("token", "Token")],
      credential: "a token",
    };
  }

  const insecure = !baseUrl.startsWith("https://");
  if (insecure)
    notes.push("The address is plain http: allowInsecureHttp is on, check that it must be");

  const group = (
    label: string,
    id: string,
    description: string,
    list: typeof tools,
    destructive = false,
  ) =>
    list.length
      ? [
          {
            id,
            label,
            description,
            ...(destructive ? { destructive: true } : {}),
            scopes: [],
            permissions: [],
            tools: list.map((t) => t.name),
          },
        ]
      : [];
  const reads = tools.filter((t) => t.method === "GET");
  const writes = tools.filter((t) => !["GET", "DELETE"].includes(t.method));
  if (writes.length)
    notes.push(
      "Writes: mark any tool that overwrites data destructive, and move it with the deletes",
    );

  const check = reads.find((t) =>
    Object.values(t.params).every((p) => p.required === false || p.default !== undefined),
  );
  if (!check)
    notes.push("No read without required params to check the connection with: set verify.path");
  const shown = (check?.result as { fields?: string[] } | undefined)?.fields?.[0] ?? "id";

  const connector = {
    id: options.id,
    name: title,
    description: firstLine(info.description) || `Let agents use ${title}.`,
    logo: `/brand/registry/${options.id}.svg`,
    identity: `${auth.credential.replace(/^\w/, (c) => c.toUpperCase())} for ${title}.`,
    accountKey: "host",
    connectionKind: "guided",
    setup: [
      {
        title: "Create a credential for agents",
        body: `In ${title}, create ${auth.credential} for agents, with only the access they need.`,
      },
    ],
    afterConnect: [],
    fields: auth.fields,
    capabilities: [
      ...group("Read", "read", `Look things up in ${title}.`, reads),
      ...group("Make changes", "write", `Create and change things in ${title}.`, writes),
      ...group(
        "Delete",
        "delete",
        `Delete things in ${title}.`,
        tools.filter((t) => t.method === "DELETE"),
        true,
      ),
    ],
    limits: [],
    verify: {
      method: "GET",
      path: check?.path ?? "/",
      label: `${title} answers`,
      show: shown,
    },
    revoke: `Delete the credential in ${title}.`,
    http: {
      baseUrl,
      ...(insecure ? { allowInsecureHttp: true } : {}),
      auth: auth.auth,
      tools,
    },
  };
  notes.push(
    "Add limits (and answerLimits on reads by id) for the values agents may be narrowed to",
    "Add paging to list tools that take a cursor or page number",
    "Rewrite the name, description, setup and revoke copy in plain words for the admin",
  );
  return { connector, notes };
}

async function main() {
  const args = process.argv.slice(2);
  const flag = (name: string) => {
    const i = args.indexOf(`--${name}`);
    return i < 0 ? undefined : args.splice(i, 2)[1];
  };
  const id = flag("id");
  const operations = flag("operations")?.split(",");
  const tags = flag("tags")?.split(",");
  const source = args[0];
  if (!source || !id || !/^[a-z0-9-]+$/.test(id) || (!operations && !tags)) {
    console.error(
      "Usage: npx tsx scripts/openapi-to-connector.ts <spec file or URL> --id <app-id> (--operations <id,id> | --tags <tag,tag>) [--out <file>]",
    );
    process.exit(2);
  }
  const out = flag("out") ?? `${id}/connector.json`;
  const remote = /^https?:\/\//.test(source);
  const raw = remote ? await (await fetch(source)).text() : readFileSync(source, "utf8");
  let spec: Json;
  try {
    spec = JSON.parse(raw);
  } catch {
    throw new Error("The spec is not JSON. A YAML spec needs converting to JSON first");
  }
  const { connector, notes } = draftConnector(spec, {
    id,
    operations,
    tags,
    source: remote ? source : undefined,
  });
  mkdirSync(dirname(out), { recursive: true });
  writeFileSync(out, `${JSON.stringify(connector, null, 2)}\n`);
  const check = connectorSchema.safeParse(connector);
  console.log(
    `Wrote ${out}: a DRAFT with ${connector.http.tools.length} tools, not a finished connector. Before using it:`,
  );
  for (const note of notes) console.log(`- ${note}`);
  if (!check.success) console.log(`It does not load yet:\n${z.prettifyError(check.error)}`);
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href)
  main().catch((error) => {
    console.error((error as Error).message);
    process.exit(1);
  });
