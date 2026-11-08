import { randomUUID } from "node:crypto";
import type { AuditEntry } from "./store.js";

// Error reporting to Sentry (or anything taking its envelopes), on only with a DSN. Each kind is
// reported where the failure is decided, grouped by its own fingerprint, carrying ids and the
// error's class, code and status only: never its message, which can hold a vendor's or a
// person's words.

export type ErrorReporting = { dsn: string; environment?: string; release?: string };
export type ReportTags = Record<string, string | number | null | undefined>;
export type Report = (kind: string, error: unknown, tags?: ReportTags) => void;

export const noReport: Report = () => {};

const PER_KIND_MS = 60_000;

export function errorReporter(
  options: ErrorReporting | null | undefined,
  deps: { fetch?: typeof fetch; now?: () => number; log?: Pick<Console, "error"> } = {},
): Report {
  if (!options?.dsn) return noReport;
  const dsn = new URL(options.dsn);
  const project = dsn.pathname.split("/").pop();
  if (!dsn.username || !project) throw new Error("SENTRY_DSN is not a Sentry DSN");
  const prefix = dsn.pathname.slice(0, dsn.pathname.lastIndexOf("/"));
  const endpoint = `${dsn.protocol}//${dsn.host}${prefix}/api/${project}/envelope/`;
  const auth = `Sentry sentry_version=7, sentry_key=${dsn.username}, sentry_client=agent-connect/1`;
  const fetchImpl = deps.fetch ?? fetch;
  const now = deps.now ?? Date.now;
  const log = deps.log ?? console;
  const last = new Map<string, number>();

  return (kind, error, tags = {}) => {
    // Per kind and app, so one broken app neither floods the quota nor hides another.
    const limitKey = `${kind}\0${tags.app ?? ""}`;
    const at = now();
    if (at - (last.get(limitKey) ?? -Infinity) < PER_KIND_MS) return;
    last.set(limitKey, at);
    const e = error instanceof Error ? error : null;
    const code = (e as { code?: unknown } | null)?.code;
    const status = (e as { status?: unknown } | null)?.status;
    const eventId = randomUUID().replaceAll("-", "");
    const event = {
      event_id: eventId,
      timestamp: at / 1000,
      platform: "node",
      level: "error",
      logger: "agent-connect",
      environment: options.environment,
      release: options.release,
      fingerprint: ["agent-connect", kind],
      tags: Object.fromEntries(
        Object.entries({ kind, ...tags })
          .filter(([, v]) => v != null)
          .map(([k, v]) => [k, String(v)]),
      ),
      exception: {
        values: [
          {
            type: kind,
            value: [e?.constructor.name, code, status ?? tags.status]
              .filter((v) => v != null)
              .join(" "),
            stacktrace: e?.stack ? { frames: frames(e.stack) } : undefined,
          },
        ],
      },
    };
    const body = [
      JSON.stringify({ event_id: eventId, sent_at: new Date(at).toISOString() }),
      JSON.stringify({ type: "event" }),
      JSON.stringify(event),
    ].join("\n");
    void fetchImpl(endpoint, {
      method: "POST",
      headers: { "content-type": "application/x-sentry-envelope", "x-sentry-auth": auth },
      body,
      signal: AbortSignal.timeout(10_000),
    })
      .then((res) => res.body?.cancel())
      .catch((err) => log.error(`[agent-connect] error report ${kind} not sent`, err));
  };
}

/** The store, reporting every audited call an app refused with 401 or 403 (its credential). */
export function reportingVendorRefusals<S extends { audit(entry: AuditEntry): Promise<void> }>(
  store: S,
  report: Report,
): S {
  return {
    ...store,
    async audit(entry: AuditEntry) {
      await store.audit(entry);
      if (entry.vendorStatus === 401 || entry.vendorStatus === 403)
        report("vendor_auth_failed", null, {
          app: entry.connectorId,
          org: entry.organizationId,
          agent: entry.agentId,
          tool: entry.tool,
          status: entry.vendorStatus,
        });
    },
  };
}

// Where it happened, from the stack's "at" lines only (the first line is the message). Sentry
// wants the innermost frame last.
function frames(stack: string) {
  return stack
    .split("\n")
    .flatMap((line) => {
      const m = /^\s+at (?:(.+?) \()?(.+?):(\d+):(\d+)\)?$/.exec(line);
      return m ? [{ function: m[1], filename: m[2], lineno: +m[3], colno: +m[4] }] : [];
    })
    .reverse();
}
