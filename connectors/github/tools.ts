import { z } from "zod";
import {
  audited,
  binary,
  Denied,
  json,
  listLimit,
  VendorFailed,
  type ToolContext,
  type Toolset,
} from "../../packages/connector-kit/src/common.js";
import { vendorError } from "../../packages/connector-kit/src/oauth.js";
import type { VendorRequest } from "../../packages/connector-kit/src/types.js";

// Reading a repository's code at a branch, tag or commit. A key's `repositories` limit keeps it
// to those owner/name pairs (GitHub's names ignore case, so the check does too). Refs and file
// paths are sent one encoded segment at a time, never with a . or .. segment. The archive takes
// a full commit SHA only, so the same call always gives the same code.

const REPOSITORY = /^[A-Za-z0-9_.-]{1,100}\/[A-Za-z0-9_.-]{1,100}$/;
const SHA = /^[0-9a-f]{40}$/;
const MAX_FILE_BYTES = 1024 * 1024;
export const MAX_TREE_ENTRIES = 10_000;

type Mark = (status: number) => void;

function repository(ctx: ToolContext, raw: string) {
  const parts = raw.split("/");
  if (!REPOSITORY.test(raw) || parts.some((p) => p === "." || p === ".."))
    throw new Denied("repository must be owner/name, such as acme/payments-api");
  const allowed = listLimit(ctx.limits, "repositories");
  if (allowed && !allowed.some((r) => r.toLowerCase() === raw.toLowerCase()))
    throw new Denied(`repository ${raw} is outside this agent's limits`);
  return `/repos/${parts.map(encodeURIComponent).join("/")}`;
}

// A ref or path as slash-separated encoded segments.
function segments(name: string, raw: string) {
  const parts = raw.split("/");
  if (!raw || raw.length > 1000 || /[\x00-\x1f]/.test(raw) || parts.some((p) => !p || p === "." || p === ".."))
    throw new Denied(`${name} is not a valid value`);
  return parts.map(encodeURIComponent).join("/");
}

async function ok(ctx: ToolContext, mark: Mark, request: VendorRequest) {
  const res = await ctx.call(request);
  mark(res.status);
  if (res.status < 200 || res.status >= 300) {
    const detail = vendorError(res.data);
    throw new VendorFailed(res.status, `GitHub answered ${res.status}${detail ? `: ${detail}` : ""}`);
  }
  return res;
}

const repoArg = z.string().max(201).describe("The repository as owner/name, such as acme/payments-api");
const refArg = z.string().max(1000).describe("A branch or tag name, or a commit SHA");

export const githubTools: Toolset = {
  names: ["github_resolve_ref", "github_list_tree", "github_read_file", "github_download_archive"],

  register(server, ctx, allowed) {
    const annotations = { readOnlyHint: true, destructiveHint: false, openWorldHint: true };

    if (allowed.has("github_resolve_ref"))
      server.registerTool(
        "github_resolve_ref",
        {
          description:
            "Find the commit a branch, tag or commit SHA points to. Its sha is the full commit SHA that github_download_archive takes.",
          inputSchema: z.object({ repository: repoArg, ref: refArg }).strict(),
          annotations,
        },
        (args) =>
          audited(ctx, "github_resolve_ref", `${args.repository}@${args.ref}`, async (mark) => {
            const res = await ok(ctx, mark, {
              method: "GET",
              path: `${repository(ctx, args.repository)}/commits/${segments("ref", args.ref)}`,
              headers: { accept: "application/vnd.github.sha" },
              responseType: "text",
            });
            const sha = String(res.data).trim();
            if (!SHA.test(sha)) throw new VendorFailed(res.status, "GitHub did not answer a commit SHA");
            return json({ repository: args.repository, ref: args.ref, sha });
          }),
      );

    if (allowed.has("github_list_tree"))
      server.registerTool(
        "github_list_tree",
        {
          description: `List every file and folder in a repository at a branch, tag or commit, or only those inside one folder. At most ${MAX_TREE_ENTRIES} entries; truncated is true when there were more.`,
          inputSchema: z
            .object({
              repository: repoArg,
              ref: refArg,
              path: z.string().max(1000).optional().describe("Only what is inside this folder, such as src/lib"),
            })
            .strict(),
          annotations,
        },
        (args) =>
          audited(ctx, "github_list_tree", `${args.repository}@${args.ref} ${args.path ?? ""}`.trim(), async (mark) => {
            if (args.path) segments("path", args.path);
            const prefix = args.path ? `${args.path}/` : "";
            const res = await ok(ctx, mark, {
              method: "GET",
              path: `${repository(ctx, args.repository)}/git/trees/${segments("ref", args.ref)}?recursive=1`,
            });
            const tree = res.data as {
              sha: string;
              truncated: boolean;
              tree: { path: string; type: string; size?: number }[];
            };
            const inside = tree.tree.filter((e) => e.path.startsWith(prefix));
            return json({
              sha: tree.sha,
              truncated: tree.truncated || inside.length > MAX_TREE_ENTRIES,
              entries: inside
                .slice(0, MAX_TREE_ENTRIES)
                .map((e) => ({ path: e.path, type: e.type, ...(e.size === undefined ? {} : { size: e.size }) })),
            });
          }),
      );

    if (allowed.has("github_read_file"))
      server.registerTool(
        "github_read_file",
        {
          description:
            "Read one file's content as text, at a branch, tag or commit. Files larger than 1 MB, and files that are not text, cannot be read here.",
          inputSchema: z
            .object({
              repository: repoArg,
              ref: refArg,
              path: z.string().max(1000).describe("The file's path in the repository, such as src/index.ts"),
            })
            .strict(),
          annotations,
        },
        (args) =>
          audited(ctx, "github_read_file", `${args.repository}@${args.ref} ${args.path}`, async (mark) => {
            const ref = encodeURIComponent(args.ref);
            const res = await ok(ctx, mark, {
              method: "GET",
              path: `${repository(ctx, args.repository)}/contents/${segments("path", args.path)}?ref=${ref}`,
            });
            const file = res.data as { type?: string; encoding?: string; content?: string; sha?: string; size?: number };
            if (Array.isArray(res.data)) throw new Denied(`${args.path} is a folder; list it with github_list_tree`);
            if (file.type !== "file") throw new Denied(`${args.path} is not a file`);
            if (file.encoding !== "base64" || (file.size ?? 0) > MAX_FILE_BYTES)
              throw new Denied(`${args.path} is larger than 1 MB; download the archive instead`);
            let text: string;
            try {
              text = new TextDecoder("utf-8", { fatal: true }).decode(Buffer.from(file.content ?? "", "base64"));
            } catch {
              throw new Denied(`${args.path} is not a text file`);
            }
            return json({ repository: args.repository, ref: args.ref, path: args.path, sha: file.sha, size: file.size, text });
          }),
      );

    if (allowed.has("github_download_archive"))
      server.registerTool(
        "github_download_archive",
        {
          description:
            "Download the code of one commit as a .tar.gz archive, for a program to unpack, not for reading. Takes the full commit SHA only (github_resolve_ref finds it for a branch or tag), so the same call always gives the same code. Archives larger than 8 MB are refused.",
          inputSchema: z
            .object({
              repository: repoArg,
              commit: z.string().max(100).describe("The full commit SHA, 40 lower-case hex characters"),
            })
            .strict(),
          annotations,
        },
        (args) =>
          audited(ctx, "github_download_archive", `${args.repository}@${args.commit}`, async (mark) => {
            if (!SHA.test(args.commit))
              throw new Denied("commit must be a full 40-character commit SHA; find it with github_resolve_ref");
            const path = `${repository(ctx, args.repository)}/tarball/${args.commit}`;
            const res = await ok(ctx, mark, { method: "GET", path, responseType: "bytes" });
            return binary(
              res.data as Uint8Array,
              res.headers?.["content-type"] || "application/gzip",
              `https://api.github.com${path}`,
            );
          }),
      );
  },
};
