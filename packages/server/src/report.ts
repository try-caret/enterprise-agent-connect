import type { AuditEntry } from "./store.js";

// Failures handed to the host's onError as they happen. Each kind is reported where the failure
// is decided, carrying ids and the error's class, code and status only: never its message, which
// can hold a vendor's or a person's words.

/** One failure, for any error service. Ids and codes only, never an error message. */
export type AgentConnectError = {
  // teams_send_failed, teams_event_failed, teams_run_start_failed, event_failed,
  // event_run_start_failed, event_poll_failed or vendor_auth_failed (README "Error reporting").
  kind: string;
  // Those that apply: org, agent, app, bot, subscription, source, tool, runner, what, status.
  tags: Record<string, string>;
  errorClass?: string;
  code?: string;
  status?: number;
  // Where it happened, innermost first (as in Error.stack), without the message line.
  frames?: { function?: string; file: string; line: number; column: number }[];
  // ISO 8601.
  timestamp: string;
};
export type OnError = (event: AgentConnectError) => void;

export type ReportTags = Record<string, string | number | null | undefined>;
export type Report = (kind: string, error: unknown, tags?: ReportTags) => void;

export const noReport: Report = () => {};

const PER_KIND_MS = 60_000;

export function errorReporter(
  onError: OnError | null | undefined,
  deps: { now?: () => number; log?: Pick<Console, "error"> } = {},
): Report {
  if (!onError) return noReport;
  const now = deps.now ?? Date.now;
  const log = deps.log ?? console;
  const last = new Map<string, number>();

  return (kind, error, tags = {}) => {
    // Per kind, app, organization, subscription and bot: one broken one neither floods the sink
    // nor hides another (a poll failure is reported only when it starts).
    const limitKey = [kind, tags.app, tags.org, tags.subscription, tags.bot].join("\0");
    const at = now();
    if (at - (last.get(limitKey) ?? -Infinity) < PER_KIND_MS) return;
    last.set(limitKey, at);
    const e = error instanceof Error ? (error as Error & { code?: unknown; status?: unknown }) : null;
    const status = typeof e?.status === "number" ? e.status : Number(tags.status) || undefined;
    const failed = (err: unknown) => log.error(`[agent-connect] onError failed for ${kind}`, err);
    try {
      const sent = onError({
        kind,
        tags: Object.fromEntries(
          Object.entries(tags)
            .filter(([, v]) => v != null)
            .map(([k, v]) => [k, String(v)]),
        ),
        errorClass: e?.constructor.name,
        code: e?.code == null ? undefined : String(e.code),
        status,
        frames: e?.stack ? frames(e.stack) : undefined,
        timestamp: new Date(at).toISOString(),
      });
      // An async hook that rejects: logged, never an unhandled rejection.
      void Promise.resolve(sent).catch(failed);
    } catch (err) {
      failed(err);
    }
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

// The stack's "at" lines only (the first line is the message).
function frames(stack: string) {
  return stack.split("\n").flatMap((line) => {
    const m = /^\s+at (?:(.+?) \()?(.+?):(\d+):(\d+)\)?$/.exec(line);
    return m ? [{ function: m[1], file: m[2], line: +m[3], column: +m[4] }] : [];
  });
}
