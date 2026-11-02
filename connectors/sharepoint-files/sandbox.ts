import { fork, type ChildProcess } from "node:child_process";
import { serialize } from "node:v8";
import { Denied } from "../../packages/connector-kit/src/common.js";

// Untrusted documents are parsed in a process of their own (parse-host.mjs, through
// child_process.fork), the parser inside it in a worker_thread with resourceLimits. The process
// watches its own resident memory and exits past a cap, so a kill only ever hits the parse that
// grew, never one that shares the server with a request holding memory, and all of the parse's
// memory goes with it (resourceLimits do not count ArrayBuffers). Two parses of each kind run at
// once, at most one per agent; others wait briefly, then are told "busy". A place is given back
// only once the parse's process has exited, on every path.

export type SandboxLimits = {
  timeoutMs?: number;
  maxOldGenerationSizeMb?: number;
  // Growth of the parse process's own resident memory that stops it.
  memoryCapMb?: number;
  // How long a call waits for a place before "busy".
  waitMs?: number;
};

export type Kind = "pdf" | "docx";

type Limits = Required<SandboxLimits>;

const DEFAULTS: Record<Kind, Limits> = {
  pdf: { timeoutMs: 20_000, maxOldGenerationSizeMb: 512, memoryCapMb: 768, waitMs: 5_000 },
  docx: { timeoutMs: 20_000, maxOldGenerationSizeMb: 768, memoryCapMb: 1_024, waitMs: 5_000 },
};
export const PLACES = 2;
// Exit codes from parse-host.mjs.
const MEMORY = 3;
// The largest answer taken from a parse process.
const MAX_ANSWER_BYTES = 80 * 1024 * 1024;

// `reason`: why a parse failed (the error's class name only), for the server's log only.
type Answer<T> = { ok: true; value: T } | { ok: false; denied: string; reason?: string };

// A parse process's message, checked before it is believed: an object with ok true and a value,
// or ok false and a refusal, within the size cap. Anything else is null.
function checked<T>(m: unknown): Answer<T> | null {
  if (typeof m !== "object" || m === null) return null;
  const a = m as { ok?: unknown; value?: unknown; denied?: unknown; reason?: unknown };
  try {
    if (serialize(m).byteLength > MAX_ANSWER_BYTES) return null;
  } catch {
    return null;
  }
  if (a.ok === true && "value" in a && a.value !== undefined) return { ok: true, value: a.value as T };
  if (a.ok === false && typeof a.denied === "string")
    return {
      ok: false,
      denied: a.denied,
      ...(typeof a.reason === "string" ? { reason: a.reason } : {}),
    };
  return null;
}

/** Text a caller gave, or a parse process answered, cut to fit an error or audit row. Text
 *  from a file itself never gets this far: refusals about a file are fixed wording (ooxml.ts). */
export const clip = (s: string, n = 200) => (s.length > n ? `${s.slice(0, n)}...` : s);

// Two places per kind, one per agent at a time.
class Places {
  private readonly owners = new Map<string, number>();
  private active = 0;
  private waiting: { owner: string; grant: () => void }[] = [];

  private free(owner: string) {
    return this.active < PLACES && !this.owners.has(owner);
  }

  private take(owner: string): () => void {
    this.active++;
    this.owners.set(owner, 1);
    let given = false;
    return () => {
      if (given) return;
      given = true;
      this.active--;
      this.owners.delete(owner);
      const next = this.waiting.find((w) => this.free(w.owner));
      if (next) {
        this.waiting = this.waiting.filter((w) => w !== next);
        next.grant();
      }
    };
  }

  acquire(owner: string, waitMs: number): Promise<() => void> {
    if (this.free(owner)) return Promise.resolve(this.take(owner));
    return new Promise((resolve, reject) => {
      const entry = {
        owner,
        grant: () => {
          clearTimeout(timer);
          resolve(this.take(owner));
        },
      };
      const timer = setTimeout(() => {
        this.waiting = this.waiting.filter((w) => w !== entry);
        reject(new Denied("Other documents are being read here: busy, try again in a moment"));
      }, waitMs);
      this.waiting.push(entry);
    });
  }
}

const places: Record<Kind, Places> = { pdf: new Places(), docx: new Places() };
const live = new Set<ChildProcess>();

/** Parse processes still running (tests: none is left once a call has answered). */
export const liveParses = () => live.size;
/** Their process ids (tests). */
export const parsePids = () => [...live].map((c) => c.pid).filter((p): p is number => p !== undefined);

/** A worker module beside this one: .ts from the sources (tests, through tsx), .js from dist. */
export function sibling(name: string): URL {
  const fromSource = import.meta.url.endsWith(".ts");
  return new URL(`./${name}${fromSource ? ".ts" : ".js"}`, import.meta.url);
}

/** Runs `file` with `data` in a parse process; it answers {ok: true, value} or {ok: false, denied}.
 *  `owner`: the agent asking, for the one-place-per-agent rule. `host`: tests only, a stand-in for
 *  parse-host.mjs that answers as a broken parse process would. It is a separate argument, not a
 *  limit, so limits read from config can never choose the program the server starts. */
export async function sandboxed<T>(
  kind: Kind,
  file: URL,
  data: unknown,
  given: SandboxLimits = {},
  owner = "unknown",
  host = new URL("./parse-host.mjs", import.meta.url),
): Promise<T> {
  const limits = { ...DEFAULTS[kind], ...given };
  const what = kind === "pdf" ? "That PDF" : "That file";
  const release = await places[kind].acquire(owner, limits.waitMs);
  try {
    return await run<T>(host, file, data, limits, what);
  } finally {
    release();
  }
}

function run<T>(host: URL, file: URL, data: unknown, limits: Limits, what: string): Promise<T> {
  const child = fork(host, [], {
    env: {},
    // No heap flag here: V8 applies it to the worker too, under its resourceLimits (seen
    // 2026-11-01: --max-old-space-size=128 made a 768 MB worker run out at 128).
    execArgv: [],
    serialization: "advanced",
    stdio: ["ignore", "ignore", "ignore", "ipc"],
  });
  live.add(child);
  let ended: (code: number | null) => void = () => {};
  const exited = new Promise<number | null>((done) => (ended = done));
  child.once("exit", (code, signal) => {
    live.delete(child);
    ended(signal ? null : code);
  });
  let answer: Answer<T> | null = null;
  let messages = 0;
  let timedOut = false;
  child.on("message", (m) => {
    // One answer only; a second, or one of the wrong shape, voids it.
    messages++;
    answer = messages === 1 ? checked<T>(m) : null;
    if (messages > 1) child.kill("SIGKILL");
  });
  child.on("error", () => {
    // A process that never started has no exit to wait for.
    if (child.pid === undefined) {
      live.delete(child);
      ended(null);
    } else child.kill("SIGKILL");
  });
  const timer = setTimeout(() => {
    timedOut = true;
    child.kill("SIGKILL");
  }, limits.timeoutMs);
  child.send({
    file: file.href,
    typescript: file.pathname.endsWith(".ts"),
    data,
    maxOldGenerationSizeMb: limits.maxOldGenerationSizeMb,
    memoryCapMb: limits.memoryCapMb,
    timeoutMs: limits.timeoutMs,
  });
  // Every path waits for the process to be gone before answering, so its place and its memory
  // are given back together.
  return exited.then((code) => {
    clearTimeout(timer);
    const got = answer as Answer<T> | null;
    if (got?.ok) return got.value;
    if (got && !got.ok) {
      // Only a plain class name is logged: anything else from the parse process (document text,
      // a newline that forges a log line) becomes "Error".
      if (got.reason) {
        const name = /^[A-Za-z][A-Za-z0-9_]{0,63}$/.test(got.reason) ? got.reason : "Error";
        console.error(`[agent-connect] ${what.toLowerCase()} failed to parse: ${name}`);
      }
      // A refusal is one plain line: control, format (bidi, zero-width) and line-separator
      // characters become a space, so a parse process cannot break or reorder the audit row.
      const denied = got.denied.replace(/[\p{C}\u2028\u2029]+/gu, " ").trim();
      throw new Denied(clip(denied || `${what} could not be read`, 300));
    }
    if (timedOut) throw new Denied(`${what} took too long to read, so it was stopped`);
    if (code === MEMORY) throw new Denied(`${what} needs too much memory to read, so it was stopped`);
    throw new Denied(`${what} could not be read`);
  });
}
