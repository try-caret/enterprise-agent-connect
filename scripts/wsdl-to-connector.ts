import { readFileSync } from "node:fs";
import { pathToFileURL } from "node:url";
import { localName, parseXml, type XmlNode } from "../connectors/soap/xml.ts";

// Drafts a connector.json with a `soap` block from a WSDL 1.1 (document/literal, the usual
// style): one tool per operation of the first SOAP port. A draft: review every tool, put
// them in the right capabilities, mark lookups readOnly and deletes destructive.
//   npx tsx scripts/wsdl-to-connector.ts <file or https URL> [--id hr] [--prefix hr] [--ws-security]

const SOAP12 = "http://schemas.xmlsoap.org/wsdl/soap12/";

export const snake = (name: string) =>
  name
    .replace(/([a-z0-9])([A-Z])/g, "$1_$2")
    .replace(/[^A-Za-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .toLowerCase()
    .replace(/^(?=[^a-z])/, "p_");

const all = (node: XmlNode, name: string): XmlNode[] =>
  node.children.flatMap((c) => [...(localName(c.name) === name ? [c] : []), ...all(c, name)]);
const kids = (node: XmlNode | undefined, name: string) =>
  node?.children.filter((c) => localName(c.name) === name) ?? [];
const named = (nodes: XmlNode[], ref: string | undefined) =>
  nodes.find((n) => n.attrs.name === localName(ref ?? ""));

function paramType(xsd: string | undefined): "string" | "integer" | "number" | "boolean" {
  const t = localName(xsd ?? "");
  if (/^(int|integer|long|short|byte|unsigned\w*|(non)?(negative|positive)Integer)$/.test(t))
    return "integer";
  if (/^(decimal|double|float)$/.test(t)) return "number";
  return t === "boolean" ? "boolean" : "string";
}

export function draftFromWsdl(
  wsdl: string,
  { id = "soap-app", prefix = "app", wsSecurity = false } = {},
): { connector: Record<string, unknown>; notes: string[] } {
  const doc = parseXml(wsdl);
  const notes: string[] = [];
  const schemas = all(doc, "schema");
  const elements = schemas.flatMap((s) => kids(s, "element").map((e) => ({ e, s })));
  const complexTypes = schemas.flatMap((s) => kids(s, "complexType"));
  const port = all(doc, "port").find((p) => kids(p, "address").length);
  if (!port) throw new Error("No SOAP port with an address in this WSDL");
  const address = kids(port, "address")[0];
  const prefixOf = (name: string) => (name.includes(":") ? name.split(":")[0] : "");
  const version = doc.attrs[`xmlns:${prefixOf(address.name)}`] === SOAP12 ? "1.2" : "1.1";
  const binding = named(
    all(doc, "binding").filter((b) => b.attrs.type),
    port.attrs.binding,
  );
  const portType = named(all(doc, "portType"), binding?.attrs.type);
  const messages = all(doc, "message");
  // An element's own fields: the direct children of its type's sequence (or all).
  const fieldsOf = (element: XmlNode | undefined) => {
    if (!element) return [];
    const type = kids(element, "complexType")[0] ?? named(complexTypes, element.attrs.type);
    const group = type && [...kids(type, "sequence"), ...kids(type, "all")][0];
    return kids(group, "element").filter((f) => f.attrs.name);
  };
  const namespaces: Record<string, string> = {};

  const tools = kids(binding, "operation").flatMap((op) => {
    const name = op.attrs.name;
    const abstract = named(kids(portType, "operation"), name);
    const message = (dir: string) => named(messages, kids(abstract, dir)[0]?.attrs.message);
    const found = elements.find(
      ({ e }) => e.attrs.name === localName(kids(message("input"), "part")[0]?.attrs.element ?? ""),
    );
    if (!found) {
      notes.push(`${name}: skipped (not document/literal with an element part)`);
      return [];
    }
    const { e: element, s: schema } = found;
    const ns = schema.attrs.targetNamespace;
    let nsPrefix = Object.keys(namespaces).find((k) => namespaces[k] === ns);
    if (!nsPrefix) {
      nsPrefix = `ns${Object.keys(namespaces).length + 1}`;
      namespaces[nsPrefix] = ns;
    }
    const fields = fieldsOf(element);
    const qualified = schema.attrs.elementFormDefault === "qualified";
    const params: Record<string, Record<string, unknown>> = {};
    let inner = "";
    for (const f of fields) {
      if (!f.attrs.type || kids(f, "complexType").length) {
        notes.push(`${name}: ${f.attrs.name} has nested fields; add it to the body by hand`);
        continue;
      }
      const param = snake(f.attrs.name);
      params[param] = {
        type: paramType(f.attrs.type),
        description: f.attrs.name,
        ...(f.attrs.minOccurs === "0" ? { required: false } : {}),
      };
      const tag = qualified ? `${nsPrefix}:${f.attrs.name}` : f.attrs.name;
      inner += `<${tag}>{${param}}</${tag}>`;
    }
    const output = kids(message("output"), "part")[0]?.attrs.element;
    const outputFields = fieldsOf(
      elements.find(({ e }) => e.attrs.name === localName(output ?? ""))?.e,
    );
    return [
      {
        name: `${prefix}_${snake(name)}`.slice(0, 64),
        description: `Calls ${name}. Describe what it does and what it answers.`,
        operation: name,
        ...(kids(op, "operation")[0]?.attrs.soapAction
          ? { soapAction: kids(op, "operation")[0].attrs.soapAction }
          : {}),
        body: `<${nsPrefix}:${element.attrs.name}>${inner}</${nsPrefix}:${element.attrs.name}>`,
        params,
        ...(output ? { result: localName(output) } : {}),
        show: outputFields[0]?.attrs.name,
      },
    ];
  });
  if (!tools.length) throw new Error("No operation could be drafted");
  const probe = tools.find((t) => Object.values(t.params).every((p) => p.required === false));
  if (!probe) notes.push("verify: no operation without required params; set verify.args by hand");
  const { show, ...verifyTool } = probe ?? tools[0];
  if (!show) notes.push("verify.show: set the answer field to show");
  notes.push("Every tool is in one capability and none is readOnly: sort them before use");

  const auth = wsSecurity
    ? { wsSecurity: { username: "username", password: "password" } }
    : { auth: { type: "basic", username: "username", password: "password" } };
  const service = all(doc, "service")[0]?.attrs.name ?? id;
  return {
    notes,
    connector: {
      id,
      name: service,
      description: `Let agents use ${service}.`,
      logo: `/brand/registry/${id}.svg`,
      identity: `An integration user of the ${service} web service.`,
      accountKey: "host",
      connectionKind: "guided",
      setup: [
        {
          title: "Create an integration user",
          body: `Ask the team that runs ${service} for a user that may call the operations agents need, and nothing more.`,
        },
      ],
      afterConnect: [],
      fields: [
        { id: "username", label: "User name", secret: false },
        { id: "password", label: "Password", secret: true },
      ],
      capabilities: [
        {
          id: "use",
          label: `Use ${service}`,
          description: "Every operation. Split reads and changes into their own capabilities.",
          scopes: [],
          permissions: [],
          tools: tools.map((t) => t.name),
        },
      ],
      limits: [],
      verify: { label: `${service} answers`, check: "operation" },
      revoke: `Ask the team that runs ${service} to disable the integration user.`,
      soap: {
        endpoint: address.attrs.location,
        version,
        namespaces,
        ...auth,
        verify: {
          tool: verifyTool.name,
          args: Object.fromEntries(
            Object.entries(verifyTool.params)
              .filter(([, p]) => p.required !== false)
              .map(([n]) => [n, "CHANGE_ME"]),
          ),
          show: show ?? "CHANGE_ME",
        },
        tools: tools.map(({ show: _, ...t }) => t),
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
  if (!source) throw new Error("Usage: wsdl-to-connector.ts <file or https URL> [--id] [--prefix]");
  const wsdl = /^https?:\/\//.test(source)
    ? await (await fetch(source)).text()
    : readFileSync(source, "utf8");
  const { connector, notes } = draftFromWsdl(wsdl, {
    id: flag("id"),
    prefix: flag("prefix"),
    wsSecurity: process.argv.includes("--ws-security"),
  });
  for (const note of notes) console.error(`note: ${note}`);
  console.log(JSON.stringify(connector, null, 2));
}
