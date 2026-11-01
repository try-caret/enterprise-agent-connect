// Reads a PDF's text with pdf.js, inside a worker_thread started through sandbox.ts with memory
// limits, an empty environment, a wall clock and its memory watched. Answers {ok, value} or
// {ok: false, denied}. Plain JavaScript, so it runs the same from the sources
// (tests) and from dist/ (the build copies it). Settings from the library vetting (2026-11-01).
import module from "node:module";
import { fileURLToPath } from "node:url";
import { parentPort, workerData } from "node:worker_threads";

// Text extraction never draws, so the optional native canvas is never loaded, even if a
// machine has it installed.
const canvasBlocked = typeof module.registerHooks === "function";
if (canvasBlocked)
  module.registerHooks({
    resolve(specifier, context, next) {
      if (specifier === "canvas" || specifier.startsWith("@napi-rs/canvas"))
        throw new Error("canvas is not available to the PDF reader");
      return next(specifier, context);
    },
  });
// No network: pdf.js is given the bytes and local font and CMap folders.
globalThis.fetch = async () => {
  throw new Error("no network in the PDF reader");
};
// pdf.js warns on import that Node should use its legacy build; nothing reads that.
console.log = () => {};
console.warn = () => {};

const { data, maxPages, maxChars, noAssets } = workerData;

let needsPassword = false;
const ENCRYPTED = "That PDF is encrypted or has a password, so it is not read";
const dir = (rel) => fileURLToPath(new URL(rel, import.meta.resolve("pdfjs-dist/build/pdf.mjs")));

try {
  const pdfjs = await import("pdfjs-dist/build/pdf.mjs");
  // The parser runs in this thread (pdf.js's "fake worker"), under this worker's limits.
  globalThis.pdfjsWorker = await import("pdfjs-dist/build/pdf.worker.mjs");
  const task = pdfjs.getDocument({
    data: new Uint8Array(data),
    disableFontFace: true,
    useSystemFonts: false,
    useWorkerFetch: false,
    stopAtErrors: true,
    enableXfa: false,
    isOffscreenCanvasSupported: false,
    disableAutoFetch: true,
    disableStream: true,
    maxImageSize: 1_000_000,
    verbosity: 0,
    ...(noAssets
      ? {}
      : {
          standardFontDataUrl: dir("../standard_fonts/"),
          cMapUrl: dir("../cmaps/"),
          cMapPacked: true,
          wasmUrl: dir("../wasm/"),
        }),
  });
  // A password is never tried.
  task.onPassword = () => {
    needsPassword = true;
    void task.destroy();
  };
  const pdf = await task.promise;
  const { info } = await pdf.getMetadata();
  if (info?.EncryptFilterName) {
    parentPort.postMessage({ ok: false, denied: ENCRYPTED });
  } else {
    const pages = [];
    let chars = 0;
    let cut = false;
    const last = Math.min(pdf.numPages, maxPages);
    for (let n = 1; n <= last; n++) {
      const page = await pdf.getPage(n);
      const content = await page.getTextContent();
      let text = "";
      for (const item of content.items) if ("str" in item) text += item.str + (item.hasEOL ? "\n" : "");
      page.cleanup();
      if (chars + text.length > maxChars) {
        text = text.slice(0, Math.max(0, maxChars - chars));
        cut = true;
      }
      chars += text.length;
      pages.push({ page: n, text });
      if (cut) break;
    }
    parentPort.postMessage({
      ok: true,
      value: { pages, pageCount: pdf.numPages, cut, canvasBlocked },
    });
  }
  await pdf.destroy();
} catch (error) {
  const name = error?.name ?? "";
  parentPort.postMessage({
    ok: false,
    denied: needsPassword || name === "PasswordException" ? ENCRYPTED : "That PDF could not be read",
  });
}
