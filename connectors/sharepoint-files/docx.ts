import type { Document, Element, Node } from "@xmldom/xmldom";
import { Denied } from "../../packages/connector-kit/src/common.js";
import {
  child,
  children,
  isW,
  makeZip,
  on,
  openZip,
  parseXml,
  serializeXml,
  shownPart,
  val,
  W,
} from "./ooxml.js";
import { clip } from "./sandbox.js";

// Word files (.docx, and .docm read as data: macros are never run, only XML is read) as
// structured text, and new Word files made from a template by filling its named fields
// (content controls, w:sdt, with a w:tag). Hidden text is reported apart from the text, since it
// is a way to plant instructions for an agent; finding it is best effort. Called inside the
// docx worker (docx-worker.ts), never on the server's own thread.

export const MAX_TEXT = 20_000;
export const MAX_ROWS = 200;
export const MAX_CELLS = 12;
export const MAX_TAGS = 200;
// Word itself allows 63 columns.
export const MAX_ROW_CELLS = 64;
export const MAX_TABLE_ROWS = 10_000;
// Tags written into [[field:<tag>]] markers.
export const PLAIN_TAG = /^[A-Za-z0-9_.-]{1,64}$/;

const DOCUMENT = "word/document.xml";
const STYLES = "word/styles.xml";
const HEADER_FOOTER = /^word\/(header|footer)(\d*)\.xml$/;
const READ_PARTS = (name: string) =>
  name === DOCUMENT ||
  name === STYLES ||
  name === "word/comments.xml" ||
  HEADER_FOOTER.test(name);
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const MARK_OPEN = (tag: string) => `[[field:${tag}]]`;
const MARK_CLOSE = "[[/field]]";
// Each "[" followed by another, so "[[[" cannot leave a "[[" behind.
const MARK_LIKE = /\[(?=\[)/g;

export type Field = {
  tag: string;
  title: string | null;
  kind: "text" | "table";
  value: string;
  rows?: string[][];
  placeholder: boolean;
  repeated: number;
};

export type DocxContent = {
  paragraphs: { i: number; style: string | null; text: string; in_table: false }[];
  tables: { i: number; after_paragraph: number | null; rows: string[][] }[];
  // Paragraphs and tables in reading order, as indexes into the two lists above.
  order: ["paragraphs" | "tables", number][];
  fields: Field[];
  hidden: { where: string; reasons: string[]; text: string }[];
  headers_footers: { part: string; text: string }[];
  comments: { author: string | null; text: string; anchor: string }[];
  tracked_changes: { insertions: number; deletions: number };
  notes: string[];
};

// ---------------------------------------------------------------------------------------------
// Character properties, through styles

type Props = {
  vanish?: boolean;
  color?: string;
  theme?: string;
  sz?: number;
  scale?: number;
  fill?: string;
};

// The properties that can hide a run, from one rPr. Both branches of markup compatibility
// (mc:Choice, mc:Fallback) count: a run hidden in either is reported.
function collect(rPr: Element | null, into: Props, branch = false) {
  if (!rPr) return;
  for (let c = rPr.firstChild; c; c = c.nextSibling) {
    if (c.nodeType !== 1) continue;
    const e = c as Element;
    if (e.namespaceURI === MC && e.localName === "AlternateContent") {
      for (let b = e.firstChild; b; b = b.nextSibling)
        if (b.nodeType === 1 && (b as Element).namespaceURI === MC) collect(b as Element, into, true);
      continue;
    }
    if (!isW(e)) continue;
    switch (e.localName) {
      case "vanish":
        if (on(e)) into.vanish = true;
        else if (!branch) into.vanish = false;
        break;
      case "color":
        into.color = (val(e) ?? "").toUpperCase();
        into.theme = (e.getAttributeNS(W, "themeColor") ?? "").toLowerCase() || undefined;
        break;
      case "sz": {
        const v = val(e);
        if (v && /^\d+$/.test(v)) into.sz = Number(v);
        break;
      }
      case "w": {
        const v = val(e);
        if (v && /^\d+%?$/.test(v)) into.scale = Number(v.replace("%", ""));
        break;
      }
      case "shd":
        into.fill = (e.getAttributeNS(W, "fill") ?? "").toUpperCase() || undefined;
        break;
    }
  }
}

type Style = { type: string; basedOn: string | null; rPr: Element | null };

/** word/styles.xml: document defaults, styles by id (with basedOn), the default paragraph style. */
export class Styles {
  private readonly styles = new Map<string, Style>();
  private readonly cache = new Map<string, Props>();
  private readonly defaults: Props = {};
  private defaultParagraph: string | null = null;

  constructor(doc: Document | null) {
    const root = doc?.documentElement;
    if (!root || !isW(root, "styles")) return;
    const defaults = child(root, "docDefaults");
    const rPrDefault = defaults ? child(defaults, "rPrDefault") : null;
    collect(rPrDefault ? child(rPrDefault, "rPr") : null, this.defaults);
    for (const s of children(root, "style")) {
      const id = s.getAttributeNS(W, "styleId");
      if (!id) continue;
      const type = s.getAttributeNS(W, "type") ?? "paragraph";
      this.styles.set(id, { type, basedOn: val(child(s, "basedOn")), rPr: child(s, "rPr") });
      const isDefault = (s.getAttributeNS(W, "default") ?? "").toLowerCase();
      if (type === "paragraph" && ["1", "true", "on"].includes(isDefault)) this.defaultParagraph = id;
    }
  }

  // A style's properties with what it is based on underneath, at most 20 levels.
  private chain(id: string | null): Props {
    if (!id) return {};
    const hit = this.cache.get(id);
    if (hit) return hit;
    const ids: string[] = [];
    for (let at: string | null = id; at && ids.length < 20 && !ids.includes(at); ) {
      ids.push(at);
      at = this.styles.get(at)?.basedOn ?? null;
    }
    const props: Props = {};
    for (const x of ids.reverse()) collect(this.styles.get(x)?.rPr ?? null, props);
    this.cache.set(id, props);
    return props;
  }

  /** A run's effective properties: defaults, paragraph style, character style, its own. */
  run(rPr: Element | null, paragraphStyle: string | null): Props {
    const props: Props = { ...this.defaults };
    Object.assign(props, this.chain(paragraphStyle ?? this.defaultParagraph));
    const rStyle = rPr ? val(child(rPr, "rStyle")) : null;
    Object.assign(props, this.chain(rStyle));
    collect(rPr, props);
    return props;
  }
}

const NO_STYLES = new Styles(null);

const nearWhite = (hex: string | undefined) => {
  if (!hex || !/^[0-9A-F]{6}$/.test(hex)) return false;
  return [0, 2, 4].every((i) => parseInt(hex.slice(i, i + 2), 16) >= 0xf0);
};
const WHITE_THEMES = new Set(["background1", "light1"]);

/** Why a run's text is hidden (best effort): hidden, white or near-white, or tiny. */
export function hiddenReasons(props: Props): string[] {
  const reasons: string[] = [];
  if (props.vanish) reasons.push("hidden text");
  if ((props.theme && WHITE_THEMES.has(props.theme)) || (!props.theme && nearWhite(props.color)))
    reasons.push("white text");
  else if (props.fill && props.color && props.color !== "AUTO" && props.fill === props.color)
    reasons.push("text the colour of its shading");
  if ((props.sz !== undefined && props.sz <= 5) || (props.scale !== undefined && props.scale <= 10))
    reasons.push("tiny text");
  return reasons;
}

// ---------------------------------------------------------------------------------------------
// Reading

const tagOf = (sdt: Element): string | null => {
  const pr = child(sdt, "sdtPr");
  const tag = pr ? val(child(pr, "tag")) : null;
  return tag ? tag : null;
};

/** Where a field sits: inside a paragraph, around paragraphs or tables, rows, or cells. */
export function fieldPlace(sdt: Element): "inline" | "block" | "row" | "cell" {
  const parent = sdt.parentNode as Element | null;
  if (isW(parent, "tbl")) return "row";
  if (isW(parent, "tr")) return "cell";
  for (let n: Node | null = parent; n; n = (n as Node).parentNode) {
    if (isW(n, "p")) return "inline";
    if (isW(n, "body") || isW(n, "tc") || isW(n, "hdr") || isW(n, "ftr")) break;
  }
  return "block";
}

const contentOf = (sdt: Element) => child(sdt, "sdtContent");
const fieldTable = (sdt: Element) => {
  const content = contentOf(sdt);
  return content && fieldPlace(sdt) === "block" ? child(content, "tbl") : null;
};

// A paragraph's text as it is built: document text, and field markers kept apart, so look-alike
// markers are neutralized on the assembled text, never one w:t at a time.
type Part = string | { mark: string };
type Inline = { plain: string; parts: Part[] };

// What one reading records. A reader without one only produces text (a field's value, a
// comment's text): nothing is counted twice.
class Recorder {
  hidden: DocxContent["hidden"] = [];
  fields = new Map<string, Field>();
  anchors = new Map<string, string>();
  open = new Set<string>();
  insertions = 0;
  deletions = 0;
  textBoxes = false;
  neutralized = false;
}

class Reader {
  // The block fields being read, outermost first: their markers wrap each paragraph.
  private blockTags: string[] = [];
  private hiddenHere: { text: string; reasons: Set<string> } | null = null;
  private paragraphStyle: string | null = null;

  constructor(
    private readonly marks: boolean,
    private readonly rec: Recorder | null,
    private readonly styles: Styles,
  ) {}

  private render(parts: Part[]): string {
    let out = "";
    let text = "";
    const flush = () => {
      if (MARK_LIKE.test(text)) {
        MARK_LIKE.lastIndex = 0;
        if (this.rec) this.rec.neutralized = true;
        text = text.replace(MARK_LIKE, "[ ");
      }
      out += text;
      text = "";
    };
    for (const p of parts)
      if (typeof p === "string") text += p;
      else {
        flush();
        out += p.mark;
      }
    flush();
    return out;
  }

  private checkTag(tag: string) {
    if (this.marks && !PLAIN_TAG.test(tag))
      throw new Denied(
        "A field has a tag that is not a plain name (letters, digits, _ . - up to 64), so fields cannot be marked: read it without mark_fields",
      );
  }

  private visible(acc: Inline, s: string) {
    acc.plain += s;
    const last = acc.parts.length - 1;
    if (last >= 0 && typeof acc.parts[last] === "string") acc.parts[last] += s;
    else acc.parts.push(s);
    if (this.rec)
      for (const id of this.rec.open) this.rec.anchors.set(id, (this.rec.anchors.get(id) ?? "") + s);
  }

  private run(r: Element, acc: Inline) {
    const reasons = hiddenReasons(this.styles.run(child(r, "rPr"), this.paragraphStyle));
    for (let c = r.firstChild; c; c = c.nextSibling) {
      if (!isW(c)) {
        if (c.nodeType === 1 && /txbxContent|AlternateContent/.test(c.localName ?? "") && this.rec)
          this.rec.textBoxes = true;
        continue;
      }
      let s = "";
      if (c.localName === "t") s = c.textContent ?? "";
      else if (c.localName === "tab" || c.localName === "ptab") s = "\t";
      else if (c.localName === "br" || c.localName === "cr") s = "\n";
      else if (c.localName === "noBreakHyphen") s = "-";
      else if ((c.localName === "drawing" || c.localName === "pict") && this.rec) {
        if (c.getElementsByTagNameNS(W, "txbxContent").length) this.rec.textBoxes = true;
      }
      if (!s) continue;
      if (reasons.length) {
        if (this.rec && this.hiddenHere) {
          this.hiddenHere.text += s;
          reasons.forEach((x) => this.hiddenHere!.reasons.add(x));
        }
      } else this.visible(acc, s);
    }
  }

  private inline(node: Node, acc: Inline) {
    for (let c = node.firstChild; c; c = c.nextSibling) {
      if (!isW(c)) continue;
      switch (c.localName) {
        case "r":
          this.run(c, acc);
          break;
        case "del":
        case "moveFrom":
          if (this.rec) this.rec.deletions++;
          break;
        case "ins":
        case "moveTo":
          if (this.rec) this.rec.insertions++;
          this.inline(c, acc);
          break;
        case "sdt": {
          const content = contentOf(c);
          if (!content) break;
          const tag = tagOf(c);
          if (tag) this.checkTag(tag);
          const sub: Inline = { plain: "", parts: [] };
          this.inline(content, sub);
          if (tag && this.rec) this.field(c, tag, sub.plain);
          acc.plain += sub.plain;
          if (tag && this.marks) acc.parts.push({ mark: MARK_OPEN(tag) }, ...sub.parts, { mark: MARK_CLOSE });
          else for (const p of sub.parts) typeof p === "string" ? this.append(acc, p) : acc.parts.push(p);
          break;
        }
        case "commentRangeStart": {
          const id = c.getAttributeNS(W, "id");
          if (id && this.rec) this.rec.open.add(id);
          break;
        }
        case "commentRangeEnd": {
          const id = c.getAttributeNS(W, "id");
          if (id && this.rec) this.rec.open.delete(id);
          break;
        }
        case "hyperlink":
        case "smartTag":
        case "customXml":
        case "fldSimple":
        case "bdo":
        case "dir":
          this.inline(c, acc);
          break;
      }
    }
  }

  private append(acc: Inline, s: string) {
    const last = acc.parts.length - 1;
    if (last >= 0 && typeof acc.parts[last] === "string") acc.parts[last] += s;
    else acc.parts.push(s);
  }

  private field(sdt: Element, tag: string, inlineValue?: string) {
    const rec = this.rec!;
    const seen = rec.fields.get(tag);
    if (seen) {
      seen.repeated++;
      return;
    }
    const pr = child(sdt, "sdtPr");
    const table = fieldTable(sdt);
    const content = contentOf(sdt);
    const quiet = new Reader(false, null, this.styles);
    const value =
      inlineValue ?? (content ? quiet.lines(content, fieldPlace(sdt)).join("\n") : "");
    rec.fields.set(tag, {
      tag,
      title: pr ? val(child(pr, "alias")) : null,
      kind: table ? "table" : "text",
      value,
      ...(table ? { rows: quiet.tableRows(table) } : {}),
      placeholder: pr ? on(child(pr, "showingPlcHdr")) : false,
      repeated: 1,
    });
  }

  /** One paragraph's visible text, wrapped in the markers of the block fields around it. */
  paragraph(p: Element, where: string): string {
    const acc: Inline = { plain: "", parts: [] };
    const outer = this.hiddenHere;
    const outerStyle = this.paragraphStyle;
    const pPr = child(p, "pPr");
    this.paragraphStyle = pPr ? val(child(pPr, "pStyle")) : null;
    this.hiddenHere = { text: "", reasons: new Set() };
    this.inline(p, acc);
    if (this.rec && this.hiddenHere.text)
      this.rec.hidden.push({
        where,
        reasons: [...this.hiddenHere.reasons],
        text: this.hiddenHere.text,
      });
    this.hiddenHere = outer;
    this.paragraphStyle = outerStyle;
    let text = this.marks ? this.render(acc.parts) : acc.plain;
    if (this.marks)
      for (const tag of [...this.blockTags].reverse()) text = `${MARK_OPEN(tag)}${text}${MARK_CLOSE}`;
    return text;
  }

  /** Block content (a cell, a header, a field's content) as lines, nested tables flattened. */
  lines(node: Node, place: string = "block", where = "body"): string[] {
    if (place === "row") return this.tableRowsOf(node, where).map((r) => r.join("\t"));
    if (place === "cell") return this.cellsOf(node, where);
    if (place === "inline") {
      const acc: Inline = { plain: "", parts: [] };
      this.inline(node, acc);
      return [this.marks ? this.render(acc.parts) : acc.plain];
    }
    const out: string[] = [];
    for (let c = node.firstChild; c; c = c.nextSibling) {
      if (!isW(c)) continue;
      if (c.localName === "p") out.push(this.paragraph(c, where));
      else if (c.localName === "tbl") out.push(...this.tableRows(c, where).map((r) => r.join("\t")));
      else if (c.localName === "sdt") this.blockField(c, () => out.push(...this.lines(contentOf(c)!, "block", where)));
      else if (c.localName === "customXml") out.push(...this.lines(c, "block", where));
      else if (c.localName === "ins" || c.localName === "moveTo") {
        if (this.rec) this.rec.insertions++;
        out.push(...this.lines(c, "block", where));
      } else if ((c.localName === "del" || c.localName === "moveFrom") && this.rec)
        this.rec.deletions++;
    }
    return out;
  }

  private blockField(sdt: Element, read: () => void) {
    const content = contentOf(sdt);
    if (!content) return;
    const tag = tagOf(sdt);
    if (tag) this.checkTag(tag);
    if (tag && this.rec) this.field(sdt, tag);
    if (tag) this.blockTags.push(tag);
    try {
      read();
    } finally {
      if (tag) this.blockTags.pop();
    }
  }

  private cellsOf(node: Node, where: string, cells: string[] = []): string[] {
    for (let c = node.firstChild; c; c = c.nextSibling) {
      if (cells.length > MAX_ROW_CELLS)
        throw new Denied(`That file has a table with more than ${MAX_ROW_CELLS} cells in a row, so it is not read`);
      if (isW(c, "tc")) cells.push(this.lines(c, "block", where).join("\n"));
      else if (isW(c, "sdt")) this.blockField(c, () => this.cellsOf(contentOf(c)!, where, cells));
      else if (isW(c, "customXml")) this.cellsOf(c, where, cells);
    }
    if (cells.length > MAX_ROW_CELLS)
      throw new Denied(`That file has a table with more than ${MAX_ROW_CELLS} cells in a row, so it is not read`);
    return cells;
  }

  private tableRowsOf(node: Node, where: string, rows: string[][] = []): string[][] {
    for (let c = node.firstChild; c; c = c.nextSibling) {
      if (rows.length > MAX_TABLE_ROWS)
        throw new Denied(`That file has a table with more than ${MAX_TABLE_ROWS} rows, so it is not read`);
      if (isW(c, "tr")) rows.push(this.cellsOf(c, where));
      else if (isW(c, "sdt")) this.blockField(c, () => this.tableRowsOf(contentOf(c)!, where, rows));
      else if (isW(c, "customXml")) this.tableRowsOf(c, where, rows);
    }
    return rows;
  }

  tableRows(tbl: Element, where = "body"): string[][] {
    return this.tableRowsOf(tbl, where);
  }

  /** The document body: paragraphs and tables in reading order. */
  body(node: Node, out: Pick<DocxContent, "paragraphs" | "tables" | "order">) {
    for (let c = node.firstChild; c; c = c.nextSibling) {
      if (!isW(c)) continue;
      if (c.localName === "p") {
        const i = out.paragraphs.length;
        const pPr = child(c, "pPr");
        const text = this.paragraph(c, `paragraph ${i}`);
        out.paragraphs.push({ i, style: pPr ? val(child(pPr, "pStyle")) : null, text, in_table: false });
        out.order.push(["paragraphs", i]);
      } else if (c.localName === "tbl") {
        const i = out.tables.length;
        const after = out.paragraphs.length ? out.paragraphs.length - 1 : null;
        out.tables.push({ i, after_paragraph: after, rows: this.tableRows(c, `table ${i}`) });
        out.order.push(["tables", i]);
      } else if (c.localName === "sdt") this.blockField(c, () => this.body(contentOf(c)!, out));
      else if (c.localName === "customXml") this.body(c, out);
      else if (c.localName === "ins" || c.localName === "moveTo") {
        if (this.rec) this.rec.insertions++;
        this.body(c, out);
      } else if ((c.localName === "del" || c.localName === "moveFrom") && this.rec)
        this.rec.deletions++;
    }
  }
}

const bodyOf = (doc: Document): Element => {
  const body = doc.documentElement && isW(doc.documentElement, "document") ? child(doc.documentElement, "body") : null;
  if (!body) throw new Denied("That file is not a Word document (it has no document body)");
  return body;
};

/** The parts of a Word file needed to read it, or (for filling) all of them. */
export function openDocx(bytes: Uint8Array, all = false): Map<string, Uint8Array> {
  const parts = openZip(bytes, all ? undefined : READ_PARTS);
  if (!parts.has(DOCUMENT)) throw new Denied("That file is not a Word document (no word/document.xml)");
  return parts;
}

export function readDocx(bytes: Uint8Array, opts: { markFields?: boolean } = {}): DocxContent {
  const parts = openDocx(bytes);
  const styles = parts.has(STYLES) ? new Styles(parseXml(parts.get(STYLES)!, STYLES)) : NO_STYLES;
  const rec = new Recorder();
  const reader = new Reader(Boolean(opts.markFields), rec, styles);
  const out: DocxContent = {
    paragraphs: [],
    tables: [],
    order: [],
    fields: [],
    hidden: [],
    headers_footers: [],
    comments: [],
    tracked_changes: { insertions: 0, deletions: 0 },
    notes: [],
  };
  reader.body(bodyOf(parseXml(parts.get(DOCUMENT)!, DOCUMENT)), out);

  const headerParts = [...parts.keys()]
    .map((name) => ({ name, m: HEADER_FOOTER.exec(name) }))
    .filter((p) => p.m)
    .sort((a, b) => a.m![1].localeCompare(b.m![1]) || Number(a.m![2] || 0) - Number(b.m![2] || 0));
  const headerRec = new Recorder();
  for (const { name, m } of headerParts) {
    const part = `${m![1]}${m![2]}`;
    const root = parseXml(parts.get(name)!, name).documentElement!;
    const text = new Reader(false, headerRec, styles).lines(root, "block", part).join("\n");
    out.headers_footers.push({ part, text });
  }

  const comments = parts.get("word/comments.xml");
  if (comments) {
    const root = parseXml(comments, "word/comments.xml").documentElement!;
    for (const c of children(root, "comment")) {
      const id = c.getAttributeNS(W, "id") ?? "";
      out.comments.push({
        author: c.getAttributeNS(W, "author") || null,
        text: new Reader(false, null, styles).lines(c).join("\n"),
        anchor: (rec.anchors.get(id) ?? "").slice(0, 500),
      });
    }
  }

  out.fields = [...rec.fields.values()];
  out.hidden = [...rec.hidden, ...headerRec.hidden];
  out.tracked_changes = { insertions: rec.insertions, deletions: rec.deletions };
  if (rec.insertions || rec.deletions)
    out.notes.push("Tracked insertions are shown in the text; tracked deletions are left out.");
  if (rec.textBoxes || headerRec.textBoxes) out.notes.push("Text in text boxes and drawings is not read.");
  if (rec.neutralized)
    out.notes.push("The document's own text contains [[: shown as [ [ so it cannot pass for a field marker.");
  return out;
}

// ---------------------------------------------------------------------------------------------
// Filling

export type FillValue = string | { text: string; highlight?: boolean } | { rows: string[][] };

// eslint-disable-next-line no-control-regex -- control characters are what it refuses
const CONTROL = /[\u0000-\u0008\u000b\u000c\u000e-\u001f￾￿]/;

const allFields = (doc: Document) =>
  Array.from(doc.getElementsByTagNameNS(W, "sdt")).filter(
    (s) => tagOf(s) !== null && isInBody(s),
  );

function isInBody(n: Node): boolean {
  for (let p = n.parentNode; p; p = p.parentNode) if (isW(p, "body")) return true;
  return false;
}

const contains = (outer: Element, inner: Element) => {
  for (let p = inner.parentNode; p; p = p.parentNode) if (p === outer) return true;
  return false;
};

/** Deep copy of `el` without revision marks and section breaks. */
function cleanCopy(el: Element | null): Element | null {
  if (!el) return null;
  const copy = el.cloneNode(true) as Element;
  for (const name of ["rPrChange", "pPrChange", "ins", "del", "sectPr"])
    for (const x of Array.from(copy.getElementsByTagNameNS(W, name))) x.parentNode?.removeChild(x);
  return copy;
}

const firstRun = (node: Element): Element | null =>
  Array.from(node.getElementsByTagNameNS(W, "r")).find(
    (r) => !r.parentNode || !(isW(r.parentNode, "del") || isW(r.parentNode, "moveFrom")),
  ) ?? null;

// rPr children that come after w:highlight in the schema's sequence (CT_RPr).
const AFTER_HIGHLIGHT = new Set([
  "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang",
  "eastAsianLayout", "specVanish", "oMath",
]);

function runProps(doc: Document, template: Element | null, highlight: boolean): Element | null {
  let rPr = cleanCopy(template ? child(template, "rPr") : null);
  // Word's "click here" text is styled PlaceholderText; typed text does not keep it.
  const style = rPr ? child(rPr, "rStyle") : null;
  if (style && /^placeholder\s*text$/i.test(val(style) ?? "")) rPr!.removeChild(style);
  if (highlight) {
    rPr ??= doc.createElementNS(W, "w:rPr");
    const old = child(rPr, "highlight");
    if (old) rPr.removeChild(old);
    const mark = doc.createElementNS(W, "w:highlight");
    mark.setAttributeNS(W, "w:val", "yellow");
    const before = children(rPr).find((c) => AFTER_HIGHLIGHT.has(c.localName ?? ""));
    rPr.insertBefore(mark, before ?? null);
  }
  return rPr && rPr.firstChild ? rPr : null;
}

/** Runs for one line of text: text, with tabs as w:tab. */
function runsFor(doc: Document, line: string, rPr: Element | null): Element[] {
  if (!line) return [];
  const r = doc.createElementNS(W, "w:r");
  if (rPr) r.appendChild(rPr.cloneNode(true));
  line.split("\t").forEach((piece, i) => {
    if (i) r.appendChild(doc.createElementNS(W, "w:tab"));
    if (!piece) return;
    const t = doc.createElementNS(W, "w:t");
    t.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
    t.appendChild(doc.createTextNode(piece));
    r.appendChild(t);
  });
  return [r];
}

function paragraphs(doc: Document, text: string, pPr: Element | null, rPr: Element | null): Element[] {
  return text.split("\n").map((line) => {
    const p = doc.createElementNS(W, "w:p");
    if (pPr) p.appendChild(pPr.cloneNode(true));
    for (const r of runsFor(doc, line, rPr)) p.appendChild(r);
    return p;
  });
}

const clear = (node: Element, keep: (c: Node) => boolean = () => false) => {
  for (const c of Array.from(node.childNodes)) if (!keep(c)) node.removeChild(c);
};

function fillText(doc: Document, sdt: Element, text: string, highlight: boolean) {
  const content = contentOf(sdt)!;
  const run = firstRun(content);
  const rPr = runProps(doc, run, highlight);
  if (fieldPlace(sdt) === "inline") {
    clear(content);
    text.split("\n").forEach((line, i) => {
      if (i) {
        const br = doc.createElementNS(W, "w:r");
        if (rPr) br.appendChild(rPr.cloneNode(true));
        br.appendChild(doc.createElementNS(W, "w:br"));
        content.appendChild(br);
      }
      for (const r of runsFor(doc, line, rPr)) content.appendChild(r);
    });
    return;
  }
  const p = content.getElementsByTagNameNS(W, "p")[0] ?? null;
  const pPr = cleanCopy(p ? child(p, "pPr") : null);
  clear(content);
  for (const para of paragraphs(doc, text, pPr && pPr.firstChild ? pPr : null, rPr))
    content.appendChild(para);
}

function setCell(doc: Document, tc: Element, text: string) {
  const p = child(tc, "p");
  const pPr = cleanCopy(p ? child(p, "pPr") : null);
  const rPr = runProps(doc, p ? firstRun(p) : null, false);
  clear(tc, (c) => isW(c, "tcPr"));
  for (const para of paragraphs(doc, text, pPr && pPr.firstChild ? pPr : null, rPr)) tc.appendChild(para);
}

const ROW_KEEP = new Set(["trPr", "tblPrEx", "tc"]);

function fillTable(doc: Document, sdt: Element, rows: string[][]) {
  const tbl = fieldTable(sdt)!;
  const trs = children(tbl, "tr");
  const template = trs.at(-1)!;
  for (const tr of trs.slice(1)) tbl.removeChild(tr);
  for (const values of rows) {
    const tr = template.cloneNode(true) as Element;
    clear(tr, (c) => isW(c) && ROW_KEEP.has(c.localName ?? ""));
    children(tr, "tc").forEach((tc, i) => setCell(doc, tc, values[i] ?? ""));
    tbl.appendChild(tr);
  }
}

/** Clears Word's placeholder flag and any data binding, which would overwrite the value. */
function settle(sdt: Element) {
  const pr = child(sdt, "sdtPr");
  if (!pr) return;
  for (const c of Array.from(pr.childNodes))
    if (c.nodeType === 1 && (isW(c, "showingPlcHdr") || c.localName === "dataBinding"))
      pr.removeChild(c);
}

type Plan = { sdt: Element; tag: string; value: FillValue };

function plan(doc: Document, fields: Record<string, FillValue>): { plans: Plan[]; unfilled: string[] } {
  const tags = Object.keys(fields);
  if (tags.length > MAX_TAGS) throw new Denied(`Fill at most ${MAX_TAGS} fields at a time`);
  const found = allFields(doc);
  const byTag = new Map<string, Element[]>();
  for (const s of found) byTag.set(tagOf(s)!, [...(byTag.get(tagOf(s)!) ?? []), s]);
  const plans: Plan[] = [];
  for (const tag of tags) {
    // The caller's own key, echoed back: a tag found in the template is that same text.
    const name = clip(tag, 100);
    const sdts = byTag.get(tag);
    if (!sdts) throw new Denied(`${name} is not a field in the template`);
    const value = fields[tag];
    const texts =
      typeof value === "string" ? [value] : "rows" in value ? value.rows.flat() : [value.text];
    for (const t of texts) {
      if (t.length > MAX_TEXT)
        throw new Denied(`The value for ${name} is longer than ${MAX_TEXT} characters`);
      if (CONTROL.test(t)) throw new Denied(`The value for ${name} has control characters`);
    }
    for (const sdt of sdts) {
      const place = fieldPlace(sdt);
      if (place === "row" || place === "cell")
        throw new Denied(`${name} wraps table rows or cells, which cannot be filled here`);
      if (!contentOf(sdt)) throw new Denied(`${name} has no content to replace`);
      const table = fieldTable(sdt);
      if (typeof value === "object" && "rows" in value) {
        if (!table) throw new Denied(`${name} is a text field: give it text, not rows`);
        const trs = children(table, "tr");
        if (trs.length < 2 || children(table).some((c) => ["sdt", "customXml"].includes(c.localName ?? "")))
          throw new Denied(`${name}'s table needs a header row and a row to repeat`);
        if (value.rows.length > MAX_ROWS)
          throw new Denied(`Give ${name} at most ${MAX_ROWS} rows`);
        const columns = children(trs.at(-1)!, "tc").length;
        for (const row of value.rows) {
          if (row.length > MAX_CELLS)
            throw new Denied(`A row for ${name} has ${row.length} cells; at most ${MAX_CELLS}`);
          if (row.length > columns)
            throw new Denied(`A row for ${name} has ${row.length} cells; its table has ${columns} columns`);
        }
      } else if (table && !(typeof value === "string" && value === ""))
        throw new Denied(`${name} is a table field: give it {"rows": [...]}`);
      else if (contentOf(sdt)!.getElementsByTagNameNS(W, "sectPr").length)
        throw new Denied(`${name} holds a section break, which cannot be replaced`);
      plans.push({ sdt, tag, value });
    }
  }
  for (const a of plans)
    for (const b of plans)
      if (a !== b && contains(a.sdt, b.sdt))
        throw new Denied(`${clip(b.tag, 100)} is inside ${clip(a.tag, 100)}: fill one of them, not both`);
  const unfilled = [...new Set(found.map((s) => tagOf(s)!))].filter((t) => !(t in fields));
  return { plans, unfilled };
}

/** The document with these fields' content taken out and their flags settled: what a fill must
 *  not change, text and markup alike. The content is put back afterwards. */
function skeleton(doc: Document, sdts: Element[]): string {
  const taken = sdts.map((sdt) => {
    settle(sdt);
    const content = contentOf(sdt);
    const kids = content ? Array.from(content.childNodes) : [];
    for (const k of kids) content!.removeChild(k);
    return { content, kids };
  });
  const text = new TextDecoder().decode(serializeXml(doc));
  for (const { content, kids } of taken) for (const k of kids) content!.appendChild(k);
  return text;
}

const same = (a: Uint8Array, b: Uint8Array) =>
  a.length === b.length && Buffer.from(a.buffer, a.byteOffset, a.length).equals(Buffer.from(b.buffer, b.byteOffset, b.length));

/** Checks a filled file against its template before it may be uploaded. Throws Denied. The
 *  output's document.xml is parsed once more here: the check is on the bytes to be uploaded. */
export function guard(
  template: Map<string, Uint8Array>,
  before: string,
  output: Uint8Array,
  tags: Set<string>,
) {
  const after = openDocx(output, true);
  const names = [...template.keys()];
  if (names.length !== after.size || names.some((n) => !after.has(n)))
    throw new Denied("The filled file's parts differ from the template's, so it was not saved (a bug)");
  for (const name of names)
    if (name !== DOCUMENT && !same(template.get(name)!, after.get(name)!))
      throw new Denied(`The fill changed ${shownPart(name, "another part")}, so it was not saved (a bug)`);
  const doc = parseXml(after.get(DOCUMENT)!, DOCUMENT);
  if (before !== skeleton(doc, allFields(doc).filter((s) => tags.has(tagOf(s)!))))
    throw new Denied("The fill changed text outside the fields, so it was not saved (a bug)");
}

export type Sabotage = (doc: Document, parts: Map<string, Uint8Array>) => void;

// The template parsed once: planned, its skeleton taken, filled and serialized. Separate from
// the guard so the template's DOM can be collected before the output is parsed.
function fillParts(
  parts: Map<string, Uint8Array>,
  fields: Record<string, FillValue>,
  sabotage?: Sabotage,
) {
  const doc = parseXml(parts.get(DOCUMENT)!, DOCUMENT);
  bodyOf(doc);
  const { plans, unfilled } = plan(doc, fields);
  const before = skeleton(doc, plans.map((p) => p.sdt));
  for (const { sdt, value } of plans) {
    if (typeof value === "object" && "rows" in value) fillTable(doc, sdt, value.rows);
    else if (fieldTable(sdt)) fillTable(doc, sdt, []);
    else if (typeof value === "string") fillText(doc, sdt, value.replace(/\r\n?/g, "\n"), false);
    else fillText(doc, sdt, value.text.replace(/\r\n?/g, "\n"), Boolean(value.highlight));
    settle(sdt);
  }
  const out = new Map(parts);
  sabotage?.(doc, out);
  out.set(DOCUMENT, serializeXml(doc));
  return { bytes: makeZip(out), before, unfilled };
}

/** A new Word file from `template` with `fields` filled. `sabotage` is for the guard's test. */
export function fillDocx(
  template: Uint8Array,
  fields: Record<string, FillValue>,
  sabotage?: Sabotage,
): { bytes: Uint8Array; filled: string[]; unfilled: string[] } {
  const parts = openDocx(template, true);
  const { bytes, before, unfilled } = fillParts(parts, fields, sabotage);
  guard(parts, before, bytes, new Set(Object.keys(fields)));
  return { bytes, filled: Object.keys(fields), unfilled };
}
