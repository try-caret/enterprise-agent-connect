import { crc32 } from "node:zlib";
import { DOMParser, XMLSerializer, type Document, type Element, type Node } from "@xmldom/xmldom";
import { unzipSync, zipSync, type Zippable } from "fflate";
import { Denied } from "../../packages/connector-kit/src/common.js";

// Untrusted Office files (zip of XML parts), opened with the limits of the library vetting
// (LIBRARY-SECURITY-VETTING, 2026-11-01): sizes are read from the zip's own directory and
// checked before anything is inflated, every inflated part must match its recorded CRC (fflate
// checks none, and a header that lies about a size would otherwise cut a part short), and XML
// that declares a DOCTYPE or an entity is never parsed.

export const MAX_ENTRIES = 2_000;
export const MAX_TOTAL_BYTES = 50 * 1024 * 1024;
export const MAX_XML_BYTES = 20 * 1024 * 1024;
const MAX_RATIO = 200;
// Below this a part cannot be a bomb, whatever its ratio (a run of spaces compresses well).
const RATIO_FLOOR = 1024 * 1024;
const MAX_DEPTH = 256;

type Directory = { name: string; crc: number; size: number; originalSize: number }[];

// The zip's central directory, read only for the CRCs fflate does not check. ZIP64 is refused:
// no Word file small enough to read here needs it.
function directory(data: Uint8Array): Directory {
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  let end = data.length - 22;
  while (end >= 0 && view.getUint32(end, true) !== 0x06054b50) {
    if (data.length - end > 65_557) end = -1;
    else end--;
  }
  if (end < 0) throw new Denied("That file is not a valid Word file (no zip directory)");
  if (end >= 20 && view.getUint32(end - 20, true) === 0x07064b50)
    throw new Denied("That file uses ZIP64, which is not read here");
  const count = view.getUint16(end + 10, true);
  if (count > MAX_ENTRIES)
    throw new Denied(`That file has more than ${MAX_ENTRIES} parts, so it is not read`);
  let at = view.getUint32(end + 16, true);
  const entries: Directory = [];
  for (let i = 0; i < count; i++) {
    if (at + 46 > data.length || view.getUint32(at, true) !== 0x02014b50)
      throw new Denied("That file is not a valid Word file (broken zip directory)");
    const flags = view.getUint16(at + 8, true);
    const nameLength = view.getUint16(at + 28, true);
    const raw = data.subarray(at + 46, at + 46 + nameLength);
    entries.push({
      name: flags & 0x800 ? new TextDecoder().decode(raw) : String.fromCharCode(...raw),
      crc: view.getUint32(at + 16, true),
      size: view.getUint32(at + 20, true),
      originalSize: view.getUint32(at + 24, true),
    });
    at += 46 + nameLength + view.getUint16(at + 30, true) + view.getUint16(at + 32, true);
  }
  return entries;
}

/** The zip's parts that `wanted` names, inflated only after every size check passed. */
export function openZip(
  data: Uint8Array,
  wanted: (name: string) => boolean = () => true,
): Map<string, Uint8Array> {
  const listed = directory(data);
  const expected = new Map<string, { crc: number; size: number }>();
  const seen = new Set<string>();
  let index = 0;
  let total = 0;
  let files: Record<string, Uint8Array>;
  try {
    files = unzipSync(data, {
      filter(file) {
        const entry = listed[index++];
        if (!entry || entry.name !== file.name)
          throw new Denied("That file is not a valid Word file (its zip directory disagrees)");
        if (index > MAX_ENTRIES)
          throw new Denied(`That file has more than ${MAX_ENTRIES} parts, so it is not read`);
        if (seen.has(file.name))
          throw new Denied(`That file holds ${file.name} twice, so it is not read`);
        seen.add(file.name);
        if (/(^|[/\\])\.\.([/\\]|$)/.test(file.name) || /^[/\\]/.test(file.name))
          throw new Denied("That file has a part outside the document, so it is not read");
        if (!wanted(file.name)) return false;
        if (!Number.isSafeInteger(file.originalSize) || !Number.isSafeInteger(file.size))
          throw new Denied("That file's zip directory is not readable");
        if (file.compression !== 0 && file.compression !== 8)
          throw new Denied("That file uses a compression Word does not, so it is not read");
        if (file.compression === 0 && file.size !== file.originalSize)
          throw new Denied("That file's zip directory is not readable");
        total += file.originalSize;
        if (file.originalSize > MAX_TOTAL_BYTES || total > MAX_TOTAL_BYTES)
          throw new Denied(
            `That file unpacks to more than ${MAX_TOTAL_BYTES / 1024 / 1024} MB, so it is not read (a zip bomb, or simply too large)`,
          );
        if (
          file.originalSize > RATIO_FLOOR &&
          file.originalSize > MAX_RATIO * Math.max(file.size, 1)
        )
          throw new Denied("That file is compressed too far to be a real document (a zip bomb)");
        expected.set(file.name, { crc: entry.crc, size: file.originalSize });
        return true;
      },
    });
  } catch (error) {
    if (error instanceof Denied) throw error;
    throw new Denied("That file is not a valid Word file (its zip could not be read)");
  }
  const parts = new Map<string, Uint8Array>();
  for (const [name, bytes] of Object.entries(files)) {
    const want = expected.get(name)!;
    if (bytes.length !== want.size)
      throw new Denied(`That file's zip directory disagrees with ${name}, so it is not read`);
    if ((crc32(bytes) >>> 0) !== want.crc)
      throw new Denied(`That file is damaged: ${name} does not match its checksum`);
    parts.set(name, bytes);
  }
  return parts;
}

/** A zip of `parts`, in the order given, each deflated (empty ones and folders stored). */
export function makeZip(parts: Map<string, Uint8Array>): Uint8Array {
  const input: Zippable = {};
  for (const [name, bytes] of parts) input[name] = [bytes, { level: bytes.length ? 6 : 0 }];
  return zipSync(input);
}

const stop = (level: string, message: string): never => {
  throw new Denied(`That file has XML that could not be read (${level}: ${message.split("\n")[0]})`);
};

/** One XML part, parsed without any DTD, with one root and a bounded depth. */
export function parseXml(bytes: Uint8Array, name: string): Document {
  if (bytes.length > MAX_XML_BYTES)
    throw new Denied(`${name} is larger than ${MAX_XML_BYTES / 1024 / 1024} MB, so it is not read`);
  let text: string;
  try {
    text = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch {
    throw new Denied(`${name} is not UTF-8 XML, so it is not read`);
  }
  if (/<!DOCTYPE|<!ENTITY/i.test(text))
    throw new Denied(`${name} declares a DOCTYPE or an entity, so it is not read`);
  let doc: Document;
  try {
    doc = new DOMParser({ locator: false, onError: stop }).parseFromString(
      text,
      "application/xml",
    );
  } catch (error) {
    if (error instanceof Denied) throw error;
    const cause = (error as { cause?: unknown }).cause;
    if (cause instanceof Denied) throw cause;
    throw new Denied(`${name} is not well-formed XML, so it is not read`);
  }
  const roots = Array.from(doc.childNodes).filter((n) => n.nodeType === 1);
  if (roots.length !== 1) throw new Denied(`${name} must have exactly one root element`);
  // Iterative, so a deep document cannot exhaust the stack here.
  const stack: [Node, number][] = [[roots[0], 1]];
  while (stack.length) {
    const [node, depth] = stack.pop()!;
    if (depth > MAX_DEPTH) throw new Denied(`${name} is nested too deeply, so it is not read`);
    for (let c = node.firstChild; c; c = c.nextSibling) if (c.nodeType === 1) stack.push([c, depth + 1]);
  }
  return doc;
}

export const serializeXml = (doc: Document) =>
  new TextEncoder().encode(new XMLSerializer().serializeToString(doc));

export const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

export const isW = (n: Node | null, local?: string): n is Element =>
  Boolean(n && n.nodeType === 1 && n.namespaceURI === W && (!local || n.localName === local));

export const children = (n: Node, local?: string): Element[] => {
  const out: Element[] = [];
  for (let c = n.firstChild; c; c = c.nextSibling) if (isW(c, local)) out.push(c);
  return out;
};

export const child = (n: Node, local: string): Element | null => children(n, local)[0] ?? null;

/** w:val of an element, or null. */
export const val = (e: Element | null): string | null =>
  e ? (e.getAttributeNS(W, "val") ?? e.getAttribute("w:val")) : null;

/** Whether an on/off property is on: present and not false, 0 or off. */
export const on = (e: Element | null): boolean =>
  Boolean(e) && !["false", "0", "off"].includes((val(e) ?? "true").toLowerCase());
