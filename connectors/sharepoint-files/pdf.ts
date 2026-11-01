import { sandboxed, type SandboxLimits } from "./sandbox.js";

// A PDF's text, read by pdf.js in a worker_thread (pdf-worker.mjs) through sandbox.ts: the
// vetting's limits (512 MB old generation, 64 MB young, 4 MB stack, an empty environment, no
// network, a 20 s wall clock, then terminate()), one PDF at a time, and the process's memory
// watched, since pdf.js inflates streams into ArrayBuffers the heap limit does not count.
// Encrypted PDFs are refused; at most 300 pages are read.

export const MAX_PAGES = 300;
// Text past this is left out with a note (paging then serves it 300,000 characters at a time).
export const MAX_PDF_CHARS = 5_000_000;

export type PdfText = {
  pages: { page: number; text: string }[];
  pageCount: number;
  notes: string[];
  // Whether the worker refused to load a canvas package (module hooks, Node 22.15 and later).
  canvasBlocked: boolean;
};

// `noAssets` (tests): give pdf.js no font, CMap or wasm folder at all.
export type PdfLimits = SandboxLimits & { noAssets?: boolean };

type Raw = { pages: PdfText["pages"]; pageCount: number; cut: boolean; canvasBlocked: boolean };

export async function readPdf(bytes: Uint8Array, limits: PdfLimits = {}): Promise<PdfText> {
  const m = await sandboxed<Raw>(
    "pdf",
    new URL("./pdf-worker.mjs", import.meta.url),
    { data: bytes, maxPages: MAX_PAGES, maxChars: MAX_PDF_CHARS, noAssets: Boolean(limits.noAssets) },
    limits,
  );
  const notes: string[] = [];
  if (m.pageCount > MAX_PAGES) notes.push(`Only the first ${MAX_PAGES} of ${m.pageCount} pages are read.`);
  if (m.cut) notes.push(`Text past ${MAX_PDF_CHARS} characters is left out.`);
  return { pages: m.pages, pageCount: m.pageCount, notes, canvasBlocked: Boolean(m.canvasBlocked) };
}
