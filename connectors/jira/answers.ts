import { VendorFailed, type ToolContext } from "../../packages/connector-kit/src/common.js";

// One way to read an Atlassian answer for the issue tools and the Assets tools: a 2xx gives
// its data, anything else becomes VendorFailed with Atlassian's own messages.

export function atlassianError(product: string, res: { status: number; data: unknown }): string {
  const d = res.data as {
    errorMessages?: unknown[];
    errors?: Record<string, unknown>;
    message?: unknown;
  } | null;
  const detail = [
    ...(Array.isArray(d?.errorMessages) ? d.errorMessages : []),
    ...Object.values(d?.errors ?? {}),
    ...(typeof d?.message === "string" ? [d.message] : []),
  ]
    .filter((m) => typeof m === "string" && m)
    .join("; ");
  return `${product} answered ${res.status}${detail ? `: ${detail}` : ""}`;
}

export function answerOf(product: string) {
  return async (
    ctx: ToolContext,
    mark: (status: number) => void,
    req: Parameters<ToolContext["call"]>[0],
  ) => {
    const res = await ctx.call(req);
    mark(res.status);
    if (res.status < 200 || res.status >= 300)
      throw new VendorFailed(res.status, atlassianError(product, res));
    return res.data as any;
  };
}
