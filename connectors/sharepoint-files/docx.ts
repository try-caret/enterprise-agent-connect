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
  val,
  W,
} from "./ooxml.js";

// Word files (.docx, and .docm read as data: macros are never run, only XML is read) as
// structured text, and new Word files made from a template by filling its named fields
// (content controls, w:sdt, with a w:tag). Hidden text (w:vanish, white, tiny) is reported apart
// from the text, since it is a way to plant instructions for an agent.

export const MAX_TEXT = 20_000;
export const MAX_ROWS = 200;
export const MAX_CELLS = 12;
export const MAX_TAGS = 200;

const DOCUMENT = "word/document.xml";
const HEADER_FOOTER = /^word\/(header|footer)(\d*)\.xml$/;
const READ_PARTS = (name: string) =>
  name === DOCUMENT || name === "word/comments.xml" || HEADER_FOOTER.test(name);
const MARK_OPEN = (tag: string) => `[[field:${tag}]]`;
const MARK_CLOSE = "[[/field]]";
const MARK_LIKE = /\[\[(?=field:|\/field\]\])/g;

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
// Reading

/** Why a run's text is hidden, from its own properties: the same three rules as v1. */
export function hiddenReasons(rPr: Element | null): string[] {
  if (!rPr) return [];
  const reasons: string[] = [];
  if (on(child(rPr, "vanish"))) reasons.push("hidden text");
  const color = (val(child(rPr, "color")) ?? "").toUpperCase();
  if (color === "FFFFFF" || color === "FEFEFE") reasons.push("white text");
  const size = val(child(rPr, "sz"));
  if (size !== null && /^\d+$/.test(size) && Number(size) <= 4) reasons.push("tiny text");
  return reasons;
}

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

type Inline = { plain: string; shown: string };

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

  constructor(
    private readonly marks: boolean,
    private readonly rec: Recorder | null,
  ) {}

  private text(s: string): string {
    if (!this.marks || !MARK_LIKE.test(s)) return s;
    MARK_LIKE.lastIndex = 0;
    if (this.rec) this.rec.neutralized = true;
    return s.replace(MARK_LIKE, "[ [");
  }

  private visible(acc: Inline, s: string) {
    acc.plain += s;
    acc.shown += this.text(s);
    if (this.rec) for (const id of this.rec.open) this.rec.anchors.set(id, (this.rec.anchors.get(id) ?? "") + s);
  }

  private run(r: Element, acc: Inline) {
    const reasons = hiddenReasons(child(r, "rPr"));
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
          const sub: Inline = { plain: "", shown: "" };
          this.inline(content, sub);
          if (tag && this.rec) this.field(c, tag, sub.plain);
          acc.plain += sub.plain;
          acc.shown += tag && this.marks ? `${MARK_OPEN(tag)}${sub.shown}${MARK_CLOSE}` : sub.shown;
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
    const quiet = new Reader(false, null);
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
    const acc: Inline = { plain: "", shown: "" };
    const outer = this.hiddenHere;
    this.hiddenHere = { text: "", reasons: new Set() };
    this.inline(p, acc);
    if (this.rec && this.hiddenHere.text)
      this.rec.hidden.push({
        where,
        reasons: [...this.hiddenHere.reasons],
        text: this.hiddenHere.text,
      });
    this.hiddenHere = outer;
    let text = acc.shown;
    if (this.marks) for (const tag of [...this.blockTags].reverse()) text = `${MARK_OPEN(tag)}${text}${MARK_CLOSE}`;
    return text;
  }

  /** Block content (a cell, a header, a field's content) as lines, nested tables flattened. */
  lines(node: Node, place: string = "block", where = "body"): string[] {
    if (place === "row") return this.tableRowsOf(node, where).map((r) => r.join("\t"));
    if (place === "cell") return this.cellsOf(node, where);
    if (place === "inline") {
      const acc: Inline = { plain: "", shown: "" };
      this.inline(node, acc);
      return [acc.shown];
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
    if (tag && this.rec) this.field(sdt, tag);
    if (tag) this.blockTags.push(tag);
    try {
      read();
    } finally {
      if (tag) this.blockTags.pop();
    }
  }

  private cellsOf(node: Node, where: string): string[] {
    const cells: string[] = [];
    for (let c = node.firstChild; c; c = c.nextSibling) {
      if (isW(c, "tc")) cells.push(this.lines(c, "block", where).join("\n"));
      else if (isW(c, "sdt")) this.blockField(c, () => cells.push(...this.cellsOf(contentOf(c)!, where)));
      else if (isW(c, "customXml")) cells.push(...this.cellsOf(c, where));
    }
    return cells;
  }

  private tableRowsOf(node: Node, where: string): string[][] {
    const rows: string[][] = [];
    for (let c = node.firstChild; c; c = c.nextSibling) {
      if (isW(c, "tr")) rows.push(this.cellsOf(c, where));
      else if (isW(c, "sdt")) this.blockField(c, () => rows.push(...this.tableRowsOf(contentOf(c)!, where)));
      else if (isW(c, "customXml")) rows.push(...this.tableRowsOf(c, where));
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
  const rec = new Recorder();
  const reader = new Reader(Boolean(opts.markFields), rec);
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
    const text = new Reader(false, headerRec).lines(root, "block", part).join("\n");
    out.headers_footers.push({ part, text });
  }

  const comments = parts.get("word/comments.xml");
  if (comments) {
    const root = parseXml(comments, "word/comments.xml").documentElement!;
    for (const c of children(root, "comment")) {
      const id = c.getAttributeNS(W, "id") ?? "";
      out.comments.push({
        author: c.getAttributeNS(W, "author") || null,
        text: new Reader(false, null).lines(c).join("\n"),
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
    out.notes.push("The document's own text contains [[field: or [[/field]]: shown as [ [ so it cannot pass for a field marker.");
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
    const sdts = byTag.get(tag);
    if (!sdts) throw new Denied(`${tag} is not a field in the template`);
    const value = fields[tag];
    const texts =
      typeof value === "string" ? [value] : "rows" in value ? value.rows.flat() : [value.text];
    for (const t of texts) {
      if (t.length > MAX_TEXT)
        throw new Denied(`The value for ${tag} is longer than ${MAX_TEXT} characters`);
      if (CONTROL.test(t)) throw new Denied(`The value for ${tag} has control characters`);
    }
    for (const sdt of sdts) {
      const place = fieldPlace(sdt);
      if (place === "row" || place === "cell")
        throw new Denied(`${tag} wraps table rows or cells, which cannot be filled here`);
      if (!contentOf(sdt)) throw new Denied(`${tag} has no content to replace`);
      const table = fieldTable(sdt);
      if (typeof value === "object" && "rows" in value) {
        if (!table) throw new Denied(`${tag} is a text field: give it text, not rows`);
        const trs = children(table, "tr");
        if (trs.length < 2 || children(table).some((c) => ["sdt", "customXml"].includes(c.localName ?? "")))
          throw new Denied(`${tag}'s table needs a header row and a row to repeat`);
        if (value.rows.length > MAX_ROWS)
          throw new Denied(`Give ${tag} at most ${MAX_ROWS} rows`);
        const columns = children(trs.at(-1)!, "tc").length;
        for (const row of value.rows)
          if (row.length > Math.min(columns, MAX_CELLS))
            throw new Denied(`A row for ${tag} has ${row.length} cells; its table has ${columns} columns`);
      } else if (table && !(typeof value === "string" && value === ""))
        throw new Denied(`${tag} is a table field: give it {"rows": [...]}`);
      else if (contentOf(sdt)!.getElementsByTagNameNS(W, "sectPr").length)
        throw new Denied(`${tag} holds a section break, which cannot be replaced`);
      plans.push({ sdt, tag, value });
    }
  }
  for (const a of plans)
    for (const b of plans)
      if (a !== b && contains(a.sdt, b.sdt))
        throw new Denied(`${b.tag} is inside ${a.tag}: fill one of them, not both`);
  const unfilled = [...new Set(found.map((s) => tagOf(s)!))].filter((t) => !(t in fields));
  return { plans, unfilled };
}

/** The document with every filled field's content taken out and its flags settled: what must
 *  not change, text and markup alike. */
function skeleton(bytes: Uint8Array, tags: Set<string>): string {
  const doc = parseXml(bytes, DOCUMENT);
  for (const sdt of allFields(doc))
    if (tags.has(tagOf(sdt)!)) {
      const content = contentOf(sdt);
      if (content) clear(content);
      settle(sdt);
    }
  return new TextDecoder().decode(serializeXml(doc));
}

const same = (a: Uint8Array, b: Uint8Array) =>
  a.length === b.length && Buffer.from(a.buffer, a.byteOffset, a.length).equals(Buffer.from(b.buffer, b.byteOffset, b.length));

/** Checks a filled file against its template before it may be uploaded. Throws Denied. */
export function guard(template: Map<string, Uint8Array>, output: Uint8Array, tags: Set<string>) {
  const after = openDocx(output, true);
  const names = [...template.keys()];
  if (names.length !== after.size || names.some((n) => !after.has(n)))
    throw new Denied("The filled file's parts differ from the template's, so it was not saved (a bug)");
  for (const name of names)
    if (name !== DOCUMENT && !same(template.get(name)!, after.get(name)!))
      throw new Denied(`The fill changed ${name}, so it was not saved (a bug)`);
  if (skeleton(template.get(DOCUMENT)!, tags) !== skeleton(after.get(DOCUMENT)!, tags))
    throw new Denied("The fill changed text outside the fields, so it was not saved (a bug)");
}

export type Sabotage = (doc: Document, parts: Map<string, Uint8Array>) => void;

/** A new Word file from `template` with `fields` filled. `sabotage` is for the guard's test. */
export function fillDocx(
  template: Uint8Array,
  fields: Record<string, FillValue>,
  sabotage?: Sabotage,
): { bytes: Uint8Array; filled: string[]; unfilled: string[] } {
  const parts = openDocx(template, true);
  const doc = parseXml(parts.get(DOCUMENT)!, DOCUMENT);
  bodyOf(doc);
  const { plans, unfilled } = plan(doc, fields);
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
  const bytes = makeZip(out);
  guard(parts, bytes, new Set(Object.keys(fields)));
  return { bytes, filled: Object.keys(fields), unfilled };
}
