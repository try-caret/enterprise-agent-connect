import { Worker } from "node:worker_threads";
import { Denied } from "../../packages/connector-kit/src/common.js";

// A PDF's text, read by pdf.js in a worker_thread (pdf-worker.mjs) with the vetting's limits:
// 512 MB old generation, 64 MB young, 4 MB stack, an empty environment, no network, a 20 s wall
// clock, then terminate(). Encrypted PDFs are refused; at most 300 pages are read.

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
export type PdfLimits = { timeoutMs?: number; maxOldGenerationSizeMb?: number; noAssets?: boolean };

export function readPdf(bytes: Uint8Array, limits: PdfLimits = {}): Promise<PdfText> {
  const worker = new Worker(new URL("./pdf-worker.mjs", import.meta.url), {
    workerData: {
      data: bytes,
      maxPages: MAX_PAGES,
      maxChars: MAX_PDF_CHARS,
      noAssets: Boolean(limits.noAssets),
    },
    resourceLimits: {
      maxOldGenerationSizeMb: limits.maxOldGenerationSizeMb ?? 512,
      maxYoungGenerationSizeMb: 64,
      stackSizeMb: 4,
    },
    env: {},
    // Not the parent's flags (a loader such as tsx): the worker is plain JavaScript.
    execArgv: [],
    stdout: true,
    stderr: true,
  });
  worker.stdout.resume();
  worker.stderr.resume();
  return new Promise<PdfText>((resolve, reject) => {
    let settled = false;
    const done = (fn: () => void) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      void worker.terminate();
      fn();
    };
    const timer = setTimeout(
      () => done(() => reject(new Denied("That PDF took too long to read, so it was stopped"))),
      limits.timeoutMs ?? 20_000,
    );
    worker.on("message", (m: any) =>
      done(() => {
        if (m?.ok) {
          const notes: string[] = [];
          if (m.pageCount > MAX_PAGES)
            notes.push(`Only the first ${MAX_PAGES} of ${m.pageCount} pages are read.`);
          if (m.cut) notes.push(`Text past ${MAX_PDF_CHARS} characters is left out.`);
          resolve({
            pages: m.pages,
            pageCount: m.pageCount,
            notes,
            canvasBlocked: Boolean(m.canvasBlocked),
          });
        } else if (m?.reason === "encrypted")
          reject(new Denied("That PDF is encrypted or has a password, so it is not read"));
        else reject(new Denied("That PDF could not be read"));
      }),
    );
    worker.on("error", (error: Error & { code?: string }) =>
      done(() =>
        reject(
          new Denied(
            error.code === "ERR_WORKER_OUT_OF_MEMORY"
              ? "That PDF needs too much memory to read, so it was stopped"
              : "That PDF could not be read",
          ),
        ),
      ),
    );
    worker.on("exit", () => done(() => reject(new Denied("That PDF could not be read"))));
  });
}
