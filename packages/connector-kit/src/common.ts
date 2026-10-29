import type { CallToolResult, McpServer } from "@modelcontextprotocol/server";
import { ConnectError, type Limits, type VendorRequest, type VendorResponse } from "./types.js";

// What a tool sees of the calling agent: its value limits, the app connection to call
// through, and the audit log. Its capabilities are applied before registration: a tool
// they do not include is never registered.
export type ToolContext = {
  limits: Limits;
  /** The connection's settings, e.g. the sites it was connected with. */
  config: Record<string, string>;
  call(request: VendorRequest): Promise<VendorResponse>;
  /** Per tool: what to add to a 403 from the app (its capability's `forbidden`). */
  forbidden?: Record<string, string>;
  audit(entry: {
    tool: string;
    target: string | null;
    decision: "allowed" | "denied" | "failed";
    reason: string | null;
    vendorStatus: number | null;
  }): Promise<void>;
};

export type Toolset = {
  names: readonly string[];
  register(server: McpServer, ctx: ToolContext, allowed: ReadonlySet<string>): void;
};

/** One app's tools from several sets, e.g. SharePoint's files and spreadsheets. */
export const toolsets = (...sets: Toolset[]): Toolset => ({
  names: sets.flatMap((s) => s.names),
  register: (server, ctx, allowed) => sets.forEach((s) => s.register(server, ctx, allowed)),
});

// Refusal of a call the grant does not cover. Its message is returned to the agent.
export class Denied extends Error {}

// Non-2xx answer from the app. Its message is returned to the agent. `hintTool`: a call a tool
// makes under another capability (always-on reads, say) names a tool of that capability, whose
// `forbidden` is then the hint added to a 403.
export class VendorFailed extends Error {
  constructor(
    readonly status: number,
    message: string,
    readonly hintTool?: string,
  ) {
    super(message);
  }
}

export function listLimit(limits: Limits, id: string): string[] | null {
  const value = limits[id];
  if (value === undefined) return null;
  if (!Array.isArray(value)) throw new Error(`Limit ${id} must be a list`);
  return value.length ? value : null;
}

export function mapLimit(limits: Limits, id: string): Record<string, string> | null {
  const value = limits[id];
  if (value === undefined) return null;
  if (typeof value !== "object" || Array.isArray(value))
    throw new Error(`Limit ${id} must be a map`);
  return Object.keys(value).length ? value : null;
}

export function textLimit(limits: Limits, id: string): string | null {
  const value = limits[id];
  if (value === undefined || value === "") return null;
  if (typeof value !== "string") throw new Error(`Limit ${id} must be text`);
  return value;
}

// The audit target of a call refused before any tool ran: what it was sent, shortened.
export const argumentsTarget = (args: unknown): string | null =>
  args === undefined ? null : JSON.stringify(args).slice(0, 300);

export const json = (data: unknown): CallToolResult => ({
  content: [{ type: "text", text: JSON.stringify(data, null, 2) }],
});

// Runs one tool call and writes its audit row whatever the outcome. `retarget` replaces the
// audited target with what was actually sent, e.g. the project-scoped JQL.
export async function audited(
  ctx: ToolContext,
  tool: string,
  target: string | null,
  run: (
    vendorStatus: (status: number) => void,
    retarget: (sent: string) => void,
  ) => Promise<CallToolResult>,
): Promise<CallToolResult> {
  let status: number | null = null;
  const fail = (text: string): CallToolResult => ({
    isError: true,
    content: [{ type: "text", text }],
  });
  let result: CallToolResult;
  try {
    result = await run(
      (s) => (status = s),
      (sent) => (target = sent),
    );
  } catch (caught) {
    // A ConnectError (an app's redirect, an unreachable address) is written for people to read.
    const error = caught instanceof ConnectError ? new VendorFailed(502, caught.message) : caught;
    if (error instanceof Denied) {
      await ctx.audit({
        tool,
        target,
        decision: "denied",
        reason: error.message,
        // Set when the app answered and the answer was refused, e.g. by an answer limit.
        vendorStatus: status,
      });
      return fail(error.message);
    }
    if (error instanceof VendorFailed) {
      const hint = error.status === 403 ? ctx.forbidden?.[error.hintTool ?? tool] : undefined;
      const message = hint ? `${error.message.replace(/\.$/, "")}. ${hint}` : error.message;
      await ctx.audit({
        tool,
        target,
        decision: "failed",
        reason: message.slice(0, 500),
        vendorStatus: error.status,
      });
      return fail(message);
    }
    // Internal details (database, keys, config) stay in the log, not with the agent.
    console.error(`[agent-connect] ${tool} failed`, error);
    await ctx.audit({
      tool,
      target,
      decision: "failed",
      reason: "Internal error",
      vendorStatus: status,
    });
    return fail("The call failed. Try again later.");
  }
  // The call already happened (mail sent, record deleted): a failed audit write must not
  // tell the agent to repeat it.
  try {
    await ctx.audit({ tool, target, decision: "allowed", reason: null, vendorStatus: status });
  } catch (error) {
    console.error(`[agent-connect] ${tool} audit write failed after a completed call`, error);
  }
  return result;
}
