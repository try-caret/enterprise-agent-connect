import { Worker } from "node:worker_threads";
import { Denied } from "../../packages/connector-kit/src/common.js";

// Untrusted documents are parsed in worker_threads, one of each kind at a time per server.
// A worker's resourceLimits cap its JavaScript heap but not ArrayBuffers (a PDF stream inflated
// by pdf.js, say), so the process's resident memory is also watched while the worker runs and
// the worker is terminated past a cap. Every path (answer, refusal, error, timeout, memory)
// terminates it.

export type SandboxLimits = {
  timeoutMs?: number;
  maxOldGenerationSizeMb?: number;
  // Growth of the process's resident memory, from the worker's start, that stops it.
  memoryCapMb?: number;
  // How long a call waits for the one worker of its kind before "busy".
  waitMs?: number;
};

export type Kind = "pdf" | "docx";

const DEFAULTS: Record<Kind, Required<SandboxLimits>> = {
  pdf: { timeoutMs: 20_000, maxOldGenerationSizeMb: 512, memoryCapMb: 768, waitMs: 5_000 },
  docx: { timeoutMs: 20_000, maxOldGenerationSizeMb: 768, memoryCapMb: 1_024, waitMs: 5_000 },
};
const WATCH_MS = 20;

/** File-controlled text (part names, tags, parser messages) cut to fit an error or audit row. */
export const clip = (s: string, n = 200) => (s.length > n ? `${s.slice(0, n)}...` : s);

class Slot {
  private busy = false;
  private waiting: (() => void)[] = [];

  acquire(waitMs: number): Promise<() => void> {
    const release = () => {
      const next = this.waiting.shift();
      if (next) next();
      else this.busy = false;
    };
    if (!this.busy) {
      this.busy = true;
      return Promise.resolve(release);
    }
    return new Promise((resolve, reject) => {
      const grant = () => {
        clearTimeout(timer);
        resolve(release);
      };
      const timer = setTimeout(() => {
        this.waiting = this.waiting.filter((g) => g !== grant);
        reject(new Denied("Another document is being read here: busy, try again in a moment"));
      }, waitMs);
      this.waiting.push(grant);
    });
  }
}

const slots: Record<Kind, Slot> = { pdf: new Slot(), docx: new Slot() };

/** A worker module beside this one: .ts from the sources (tests, through tsx), .js from dist. */
export function sibling(name: string): URL {
  const fromSource = import.meta.url.endsWith(".ts");
  return new URL(`./${name}${fromSource ? ".ts" : ".js"}`, import.meta.url);
}

export type Words = { what: string };

/** Runs `file` with `data` as workerData; it answers {ok: true, value} or {ok: false, denied}. */
export async function sandboxed<T>(
  kind: Kind,
  file: URL,
  data: unknown,
  given: SandboxLimits = {},
): Promise<T> {
  const limits = { ...DEFAULTS[kind], ...given };
  const what = kind === "pdf" ? "That PDF" : "That file";
  const release = await slots[kind].acquire(limits.waitMs);
  try {
    return await run<T>(file, data, limits, what);
  } finally {
    release();
  }
}

function run<T>(file: URL, data: unknown, limits: Required<SandboxLimits>, what: string): Promise<T> {
  const typescript = file.pathname.endsWith(".ts");
  const base = process.memoryUsage.rss();
  const worker = new Worker(file, {
    workerData: data,
    resourceLimits: {
      maxOldGenerationSizeMb: limits.maxOldGenerationSizeMb,
      maxYoungGenerationSizeMb: 64,
      stackSizeMb: 4,
    },
    env: {},
    // Not the parent's flags; from the sources only the TypeScript loader.
    execArgv: typescript ? ["--import", "tsx"] : [],
    stdout: true,
    stderr: true,
  });
  worker.stdout.resume();
  worker.stderr.resume();
  return new Promise<T>((resolve, reject) => {
    let settled = false;
    const done = (fn: () => void) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      clearInterval(watch);
      void worker.terminate();
      fn();
    };
    const tooMuch = () => reject(new Denied(`${what} needs too much memory to read, so it was stopped`));
    const timer = setTimeout(
      () => done(() => reject(new Denied(`${what} took too long to read, so it was stopped`))),
      limits.timeoutMs,
    );
    const watch = setInterval(() => {
      if (process.memoryUsage.rss() - base > limits.memoryCapMb * 1024 * 1024) done(tooMuch);
    }, WATCH_MS);
    worker.on("message", (m: { ok?: boolean; value?: T; denied?: string }) =>
      done(() =>
        m?.ok
          ? resolve(m.value as T)
          : reject(new Denied(clip(m?.denied ?? `${what} could not be read`, 300))),
      ),
    );
    worker.on("error", (error: Error & { code?: string }) =>
      done(() =>
        error.code === "ERR_WORKER_OUT_OF_MEMORY"
          ? tooMuch()
          : reject(new Denied(`${what} could not be read`)),
      ),
    );
    worker.on("exit", () => done(() => reject(new Denied(`${what} could not be read`))));
  });
}
