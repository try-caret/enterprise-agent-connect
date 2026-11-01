import { parentPort, workerData } from "node:worker_threads";
import { Denied } from "../../packages/connector-kit/src/common.js";
import { fillDocx, readDocx, type DocxContent, type FillValue } from "./docx.js";
import { pieces, type Item } from "./paging.js";

// Reads or fills one Word file inside a worker_thread started by sandbox.ts, so XML parsing
// (xmldom has no limits of its own) runs under the worker's memory limits and clock, never on
// the server's own thread. Any failure that is not a refusal becomes a plain one.

type Job =
  | { op: "read"; bytes: Uint8Array; markFields: boolean }
  | { op: "fill"; bytes: Uint8Array; fields: Record<string, FillValue> };

export type ReadAnswer = {
  items: Item[];
  tracked_changes: DocxContent["tracked_changes"];
  notes: string[];
};

function items(doc: DocxContent): Item[] {
  return [
    ...doc.order.flatMap(([section, i]) =>
      pieces(section, [section === "paragraphs" ? doc.paragraphs[i] : doc.tables[i]]),
    ),
    ...pieces("fields", doc.fields),
    ...pieces("hidden", doc.hidden),
    ...pieces("headers_footers", doc.headers_footers),
    ...pieces("comments", doc.comments),
  ];
}

const job = workerData as Job;
try {
  if (job.op === "read") {
    const doc = readDocx(job.bytes, { markFields: job.markFields });
    const answer: ReadAnswer = {
      items: items(doc),
      tracked_changes: doc.tracked_changes,
      notes: doc.notes,
    };
    parentPort!.postMessage({ ok: true, value: answer });
  } else {
    parentPort!.postMessage({ ok: true, value: fillDocx(job.bytes, job.fields) });
  }
} catch (error) {
  const failure = error as Error;
  parentPort!.postMessage({
    ok: false,
    denied: error instanceof Denied ? error.message : "That file could not be read",
    // For the server's log only, never the caller: the error's class and message, cut short.
    ...(error instanceof Denied
      ? {}
      : { reason: `${failure?.name ?? "Error"}: ${String(failure?.message ?? error).slice(0, 200)}` }),
  });
}
