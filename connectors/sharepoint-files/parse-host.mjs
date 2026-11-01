// One parse in a process of its own, started by sandbox.ts with child_process.fork. The parser
// (docx-worker or pdf-worker) runs in a worker_thread here under resourceLimits, while this
// thread, free the whole time, polls this process's own resident memory: ArrayBuffers (a PDF
// stream pdf.js inflates) count there though the worker's heap limit misses them, and nothing
// another request allocates in the server does. Past the cap the process exits at once with
// code 3, which takes the parser and all its memory with it. Plain JavaScript, so it runs the
// same from the sources and from dist/. It never outlives the server: it exits when the IPC
// channel closes (the server went away) and on a clock of its own, a little past the server's.
import { serialize } from "node:v8";
import { Worker } from "node:worker_threads";

const MEMORY = 3;
const FAILED = 4;
const GONE = 5;
const CLOCK = 6;
const WATCH_MS = 5;
// The largest answer relayed to the server.
const MAX_ANSWER_BYTES = 80 * 1024 * 1024;

process.on("disconnect", () => process.exit(GONE));
// No job within this long: nothing is coming.
const idle = setTimeout(() => process.exit(CLOCK), 10_000);

process.once("message", ({ file, typescript, data, maxOldGenerationSizeMb, memoryCapMb, timeoutMs }) => {
  clearTimeout(idle);
  setTimeout(() => process.exit(CLOCK), timeoutMs + 2_000).unref?.();
  const base = process.memoryUsage.rss();
  const watch = setInterval(() => {
    if (process.memoryUsage.rss() - base > memoryCapMb * 1024 * 1024) process.exit(MEMORY);
  }, WATCH_MS);
  const worker = new Worker(new URL(file), {
    workerData: data,
    resourceLimits: { maxOldGenerationSizeMb, maxYoungGenerationSizeMb: 64, stackSizeMb: 4 },
    env: {},
    execArgv: typescript ? ["--import", "tsx"] : [],
    stdout: true,
    stderr: true,
  });
  worker.stdout.resume();
  worker.stderr.resume();
  const finish = (code) => {
    clearInterval(watch);
    process.exit(code);
  };
  let answered = false;
  worker.once("message", (m) => {
    answered = true;
    let size;
    try {
      size = serialize(m).byteLength;
    } catch {
      return finish(FAILED);
    }
    if (size > MAX_ANSWER_BYTES) return finish(FAILED);
    try {
      process.send(m, (error) => finish(error ? FAILED : 0));
    } catch {
      finish(FAILED);
    }
  });
  worker.once("error", (error) => {
    if (answered) return;
    answered = true;
    if (error?.code === "ERR_WORKER_OUT_OF_MEMORY") return finish(MEMORY);
    // An uncaught error in the parser: its class name alone goes to the server's log (a message
    // can carry document text).
    const name = String(error?.name ?? "");
    const reason = /^[A-Za-z][A-Za-z0-9_]{0,63}$/.test(name) ? name : "Error";
    try {
      process.send({ ok: false, denied: "", reason }, () => finish(FAILED));
    } catch {
      finish(FAILED);
    }
  });
  worker.once("exit", () => {
    if (!answered) finish(FAILED);
  });
});
