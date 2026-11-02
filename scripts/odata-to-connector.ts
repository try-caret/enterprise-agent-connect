import { readFileSync } from "node:fs";
import { pathToFileURL } from "node:url";
import { localName, parseXml, type XmlNode } from "../connectors/soap/xml.ts";
import { snake } from "./wsdl-to-connector.ts";

// Drafts a connector.json with an `odata` block from a service's $metadata (v2 or v4): a list
// and a get tool per entity set, and a tool per unbound function or action import. Writes
// (create, update) are left to add by hand. A draft: trim `select`, add filters and limits.
//   npx tsx scripts/odata-to-connector.ts <$metadata file or https URL> --root <service root>
//     [--id sap-orders] [--prefix sap] [--sets A_SalesOrder,A_Customer]

const SELECTED = 12;

const all = (node: XmlNode, name: string): XmlNode[] =>
  node.children.flatMap((c) => [...(localName(c.name) === name ? [c] : []), ...all(c, name)]);
const kids = (node: XmlNode | undefined, name: string) =>
  node?.children.filter((c) => localName(c.name) === name) ?? [];
const attr = (node: XmlNode, name: string) =>
  Object.entries(node.attrs).find(([k]) => localName(k) === name)?.[1];

function valueType(edm: string | undefined) {
  const t = (edm ?? "").replace(/^Edm\./, "");
  if (/^(Int16|Int32|Int64|Byte|SByte)$/.test(t)) return "integer";
  if (/^(Double|Single)$/.test(t)) return "number";
  if (t === "Decimal") return "decimal";
  if (t === "Boolean") return "boolean";
  if (t === "Guid") return "guid";
  if (t === "DateTimeOffset") return "datetimeoffset";
  if (/^(DateTime|Date)$/.test(t)) return "datetime";
  return t === "String" ? "string" : null;
}

export function draftFromMetadata(
  metadata: string,
  {
    root,
    id = "odata-app",
    prefix = "app",
    sets,
  }: { root: string; id?: string; prefix?: string; sets?: string[] },
): { connector: Record<string, unknown>; notes: string[] } {
  const doc = parseXml(metadata);
  const version = attr(doc, "Version")?.startsWith("4") ? "4" : "2";
  const notes: string[] = [];
  const types = all(doc, "EntityType");
  const byName = (nodes: XmlNode[], qualified: string | undefined) =>
    nodes.find((n) => n.attrs.Name === (qualified ?? "").split(".").pop());
  const value = (p: XmlNode) => {
    const type = valueType(p.attrs.Type);
    return type ? { type, description: p.attrs.Name } : null;
  };

  const tools: Record<string, unknown>[] = [];
  const writes: string[] = [];
  for (const set of all(doc, "EntitySet")) {
    const name = set.attrs.Name;
    if (sets && !sets.includes(name)) continue;
    const type = byName(types, set.attrs.EntityType);
    if (!type) continue;
    const keys = kids(kids(type, "Key")[0], "PropertyRef").map((k) => k.attrs.Name);
    const props = kids(type, "Property");
    const keyValues = Object.fromEntries(
      keys.flatMap((k) => {
        const v = props.find((p) => p.attrs.Name === k);
        const typed = v && value(v);
        return typed ? [[snake(k), { ...typed, property: k }]] : [];
      }),
    );
    if (Object.keys(keyValues).length !== keys.length) {
      notes.push(`${name}: skipped (a key of a type the connector does not take)`);
      continue;
    }
    const select = [...keys, ...props.map((p) => p.attrs.Name).filter((n) => !keys.includes(n))];
    if (select.length > SELECTED)
      notes.push(`${name}: ${select.length} properties, the first ${SELECTED} selected`);
    const noun = snake(name);
    tools.push(
      {
        name: `${prefix}_list_${noun}`.slice(0, 64),
        description: `List ${name}. Answers {items, next}; pass next as cursor for more.`,
        kind: "list",
        entitySet: name,
        filters: Object.fromEntries(
          Object.entries(keyValues).map(([k, v]) => [k, { ...v, required: false }]),
        ),
        select: select.slice(0, SELECTED),
      },
      {
        name: `${prefix}_get_${noun}`.slice(0, 64),
        description: `Read one of ${name} by its key.`,
        kind: "get",
        entitySet: name,
        key: keyValues,
        select: select.slice(0, SELECTED),
      },
    );
  }

  const fn = (node: XmlNode, kind: "function" | "action", parameters: XmlNode[]) => {
    const params: Record<string, unknown> = {};
    for (const p of parameters) {
      const typed = value(p);
      if (!typed) {
        notes.push(`${node.attrs.Name}: parameter ${p.attrs.Name} has a type to map by hand`);
        continue;
      }
      params[snake(p.attrs.Name)] = { ...typed, property: p.attrs.Name };
    }
    tools.push({
      name: `${prefix}_${snake(node.attrs.Name)}`.slice(0, 64),
      description: `Calls ${node.attrs.Name}. Describe what it does.`,
      kind,
      function: node.attrs.Name,
      params,
    });
    if (kind === "action") writes.push(tools.at(-1)!.name as string);
  };
  if (version === "2")
    for (const f of all(doc, "FunctionImport"))
      fn(f, attr(f, "HttpMethod") === "POST" ? "action" : "function", kids(f, "Parameter"));
  else {
    for (const f of all(doc, "FunctionImport"))
      fn(f, "function", kids(byName(all(doc, "Function"), f.attrs.Function), "Parameter"));
    for (const a of all(doc, "ActionImport"))
      fn(a, "action", kids(byName(all(doc, "Action"), a.attrs.Action), "Parameter"));
  }
  if (!tools.length) throw new Error("No entity set or import could be drafted");
  notes.push("Add create and update tools by hand, with limits on the fields agents may set");

  const reads = tools.map((t) => t.name as string).filter((t) => !writes.includes(t));
  const capability = (cid: string, label: string, list: string[]) => ({
    id: cid,
    label,
    description: `${label}. Describe it for the admin.`,
    scopes: [],
    permissions: [],
    tools: list,
  });
  return {
    notes,
    connector: {
      id,
      name: id,
      description: `Let agents use ${id}.`,
      logo: `/brand/registry/${id}.svg`,
      identity: "A technical user of the OData service.",
      accountKey: "host",
      connectionKind: "guided",
      setup: [
        {
          title: "Create a technical user",
          body: "Ask the team that runs the service for a technical user with read access to what agents need, and nothing more.",
        },
      ],
      afterConnect: [],
      fields: [
        { id: "username", label: "User name", secret: false },
        { id: "password", label: "Password", secret: true },
      ],
      capabilities: [
        ...(reads.length ? [capability("read", "Read", reads)] : []),
        ...(writes.length ? [capability("actions", "Run actions", writes)] : []),
      ],
      limits: [],
      verify: {
        method: "GET",
        path: "/",
        label: "The service answers",
        show: version === "2" ? "d.EntitySets[]" : "value[].name",
      },
      revoke: "Lock or delete the technical user.",
      odata: {
        serviceRoot: root,
        version,
        auth: { type: "basic", username: "username", password: "password" },
        csrf: version === "2",
        tools,
      },
    },
  };
}

function flag(name: string): string | undefined {
  const i = process.argv.indexOf(`--${name}`);
  return i > 0 ? process.argv[i + 1] : undefined;
}

if (import.meta.url === pathToFileURL(process.argv[1] ?? "").href) {
  const source = process.argv[2];
  const root = flag("root");
  if (!source || !root)
    throw new Error("Usage: odata-to-connector.ts <$metadata file or URL> --root <service root>");
  const metadata = /^https?:\/\//.test(source)
    ? await (await fetch(source)).text()
    : readFileSync(source, "utf8");
  const { connector, notes } = draftFromMetadata(metadata, {
    root,
    id: flag("id"),
    prefix: flag("prefix"),
    sets: flag("sets")?.split(","),
  });
  for (const note of notes) console.error(`note: ${note}`);
  console.log(JSON.stringify(connector, null, 2));
}
