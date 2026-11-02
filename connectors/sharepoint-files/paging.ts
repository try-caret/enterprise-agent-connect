import { Denied } from "../../packages/connector-kit/src/common.js";

// Answers of at most 300,000 characters: a document is a list of items (paragraphs, tables,
// fields...), an item too large for one answer is split into marked parts, and an answer holds
// the items from a cursor that fit.

export const MAX_ANSWER_CHARS = 300_000;
// Room in each answer for the file's details, counts, `next` and the note.
export const ITEM_BUDGET = MAX_ANSWER_CHARS - 5_000;
const MAX_PIECES = 100_000;

export type Piece = Record<string, unknown>;
export type Item = { section: string; obj: Piece };

const length = (o: unknown) => JSON.stringify(o).length;

// One item in two, on its largest property that can still be split: rows, then a single row's
// cells, then a text. Null when nothing can be split further.
function split(obj: Piece): [Piece, Piece] | null {
  const keys = Object.keys(obj)
    .filter((k) => typeof obj[k] === "string" || (k === "rows" && Array.isArray(obj[k])))
    .sort((a, b) => length(obj[b]) - length(obj[a]));
  for (const key of keys) {
    if (key === "rows") {
      const rows = obj.rows as string[][];
      const first = (obj.first_row as number | undefined) ?? 0;
      if (rows.length > 1) {
        const mid = Math.ceil(rows.length / 2);
        return [
          { ...obj, rows: rows.slice(0, mid), first_row: first },
          { ...obj, rows: rows.slice(mid), first_row: first + mid },
        ];
      }
      const cells = rows[0] ?? [];
      if (cells.some((c) => c.length > 1)) {
        const cut = (c: string) => Math.ceil(c.length / 2);
        return [
          { ...obj, rows: [cells.map((c) => c.slice(0, cut(c)))], first_row: first },
          { ...obj, rows: [cells.map((c) => c.slice(cut(c)))], first_row: first },
        ];
      }
      continue;
    }
    const text = obj[key] as string;
    if (text.length < 2) continue;
    const mid = Math.ceil(text.length / 2);
    return [
      { ...obj, [key]: text.slice(0, mid) },
      { ...obj, [key]: text.slice(mid) },
    ];
  }
  return null;
}

/** An item as pieces that each fit `max`, in order; iterative, so no input exhausts the stack. */
export function halve(obj: Piece, max: number): Piece[] {
  const out: Piece[] = [];
  const stack: Piece[] = [obj];
  while (stack.length) {
    if (out.length + stack.length > MAX_PIECES)
      throw new Denied("That document is too large to read in parts");
    const next = stack.pop()!;
    const halves = length(next) > max ? split(next) : null;
    if (halves) stack.push(halves[1], halves[0]);
    else out.push(next);
  }
  return out;
}

export function pieces(section: string, objs: Piece[], max = ITEM_BUDGET): Item[] {
  return objs.flatMap((obj) => {
    const parts = halve(obj, max);
    return parts.map((p, k) => ({
      section,
      obj: parts.length > 1 ? { ...p, part: k + 1, parts: parts.length } : p,
    }));
  });
}

/** The items from `start` that fit in one answer, by section, and where the next one starts. */
export function page(items: Item[], start: number, sections: string[], max = ITEM_BUDGET) {
  const out: Record<string, Piece[]> = Object.fromEntries(sections.map((s) => [s, []]));
  let used = 0;
  let i = start;
  for (; i < items.length; i++) {
    const size = length(items[i].obj) + 1;
    if (i > start && used + size > max) break;
    used += size;
    out[items[i].section].push(items[i].obj);
  }
  return { sections: out, next: i < items.length ? i : null };
}
