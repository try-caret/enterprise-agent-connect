// A small XML reader for SOAP answers, WSDL and OData $metadata: elements, attributes, text
// and CDATA. A document type declaration is refused, so there are no external or expanding
// entities: only XML's five named ones and character references.

export type XmlNode = {
  name: string;
  attrs: Record<string, string>;
  children: XmlNode[];
  text: string;
};

export class XmlError extends Error {}

const MAX_DEPTH = 100;
const START =
  /<([A-Za-z_][\w.:-]*)((?:\s+[A-Za-z_][\w.:-]*\s*=\s*(?:"[^"<]*"|'[^'<]*'))*)\s*(\/?)>/y;
const ATTR = /([A-Za-z_][\w.:-]*)\s*=\s*(?:"([^"<]*)"|'([^'<]*)')/g;
const NAMED: Record<string, string> = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" };

function decode(raw: string): string {
  return raw.replace(/&(?:#x([0-9a-fA-F]+)|#([0-9]+)|([a-z]+))?;?/g, (m, hex, dec, name) => {
    if (!m.endsWith(";") || m === "&;") throw new XmlError("A bare & in the text");
    if (name !== undefined) {
      if (!Object.hasOwn(NAMED, name)) throw new XmlError(`Unknown entity &${name};`);
      return NAMED[name];
    }
    const code = hex !== undefined ? parseInt(hex, 16) : Number(dec);
    if (!(code > 0 && code <= 0x10ffff)) throw new XmlError("A character reference out of range");
    return String.fromCodePoint(code);
  });
}

export function parseXml(xml: string): XmlNode {
  const root: XmlNode = { name: "", attrs: {}, children: [], text: "" };
  const stack = [root];
  const until = (end: string, from: number) => {
    const at = xml.indexOf(end, from);
    if (at === -1) throw new XmlError(`Unclosed ${xml.slice(from, from + 9)}`);
    return at;
  };
  let i = 0;
  while (i < xml.length) {
    const lt = xml.indexOf("<", i);
    const text = xml.slice(i, lt === -1 ? xml.length : lt);
    if (stack.length > 1) stack.at(-1)!.text += decode(text);
    else if (text.trim()) throw new XmlError("Text outside the root element");
    if (lt === -1) break;
    if (xml.startsWith("<!--", lt)) i = until("-->", lt + 4) + 3;
    else if (xml.startsWith("<![CDATA[", lt)) {
      const end = until("]]>", lt + 9);
      if (stack.length === 1) throw new XmlError("Text outside the root element");
      stack.at(-1)!.text += xml.slice(lt + 9, end);
      i = end + 3;
    } else if (xml.startsWith("<!", lt))
      throw new XmlError("A document type declaration is not accepted");
    else if (xml.startsWith("<?", lt)) i = until("?>", lt + 2) + 2;
    else if (xml.startsWith("</", lt)) {
      const end = until(">", lt + 2);
      const name = xml.slice(lt + 2, end).trim();
      const open = stack.pop();
      if (stack.length === 0 || open!.name !== name)
        throw new XmlError(`Unexpected closing tag ${name}`);
      i = end + 1;
    } else {
      START.lastIndex = lt;
      const m = START.exec(xml);
      if (!m) throw new XmlError("A malformed tag");
      const node: XmlNode = { name: m[1], attrs: {}, children: [], text: "" };
      for (const a of m[2].matchAll(ATTR)) {
        if (Object.hasOwn(node.attrs, a[1])) throw new XmlError(`Attribute ${a[1]} twice`);
        node.attrs[a[1]] = decode(a[2] ?? a[3]);
      }
      if (stack.length === 1 && root.children.length) throw new XmlError("Two root elements");
      stack.at(-1)!.children.push(node);
      if (!m[3]) stack.push(node);
      if (stack.length > MAX_DEPTH) throw new XmlError("The XML is nested too deeply");
      i = START.lastIndex;
    }
  }
  if (stack.length !== 1 || root.children.length !== 1) throw new XmlError("Unclosed elements");
  return root.children[0];
}

export const localName = (name: string) => name.slice(name.indexOf(":") + 1);

// A key set as data, never as the object's prototype (an element named __proto__).
function put(out: Record<string, unknown>, key: string, value: unknown) {
  Object.defineProperty(out, key, { value, enumerable: true, writable: true, configurable: true });
}

// An element as plain JSON: names without prefixes, repeated children as a list, text
// trimmed, an element with only text as its text, xsi:nil as null. Attributes with a prefix
// (namespaces, xsi:type) are left out; others become keys, which child elements win over.
export function toJson(node: XmlNode): unknown {
  const nil = Object.entries(node.attrs).some(([k, v]) => localName(k) === "nil" && v === "true");
  if (nil) return null;
  const out: Record<string, unknown> = {};
  const listed = new Set<string>();
  for (const child of node.children) {
    const key = localName(child.name);
    const value = toJson(child);
    if (!listed.has(key)) put(out, key, value);
    else if (Array.isArray(out[key])) (out[key] as unknown[]).push(value);
    else put(out, key, [out[key], value]);
    listed.add(key);
  }
  for (const [k, v] of Object.entries(node.attrs))
    if (!k.includes(":") && k !== "xmlns" && !listed.has(k)) put(out, k, v);
  const text = node.text.trim();
  if (!Object.keys(out).length) return text;
  if (text) put(out, "_text", text);
  return out;
}

export const escapeXml = (value: string) =>
  value.replace(
    /[&<>"']/g,
    (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&apos;" })[c]!,
  );
