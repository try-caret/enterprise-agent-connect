// One parse in a process of its own, started by sandbox.ts with child_process.fork. The parser
// (docx-worker or pdf-worker) runs in a worker_thread here under resourceLimits, while this
// thread, free the whole time, polls this process's own resident memory: ArrayBuffers (a PDF
// stream pdf.js inflates) count there though the worker's heap limit misses them, and nothing
// another request allocates in the server does. Past the cap the process exits at once with
// code 3, which takes the parser and all its memory with it. Plain JavaScript, so it runs the
// same from the sources and from dist/.
import { Worker } from "node:worker_threads";

const MEMORY = 3;
const FAILED = 4;
const WATCH_MS = 5;

process.once("message", ({ file, typescript, data, maxOldGenerationSizeMb, memoryCapMb }) => {
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
    process.send(m, () => finish(0));
  });
  worker.once("error", (error) => {
    if (!answered) finish(error?.code === "ERR_WORKER_OUT_OF_MEMORY" ? MEMORY : FAILED);
  });
  worker.once("exit", () => {
    if (!answered) finish(FAILED);
  });
});
