import { z } from "zod";
import { normalizeSiteUrl, sitePath } from "./vendor.js";
import { ConnectError } from "../../packages/connector-kit/src/types.js";
import {
  audited,
  Denied,
  json,
  listLimit,
  mapLimit,
  VendorFailed,
  type ToolContext,
  type Toolset,
} from "../../packages/connector-kit/src/common.js";

// Files in the SharePoint sites of the agent's grant (its `sites` limit, and within a site
// the folder its `folders` limit names, if any), and the directory. Microsoft also enforces
// the site grants the customer gave the app (Sites.Selected).

export const MAX_READ_BYTES = 1_000_000;
// The MCP endpoint takes request bodies up to 4 MiB, and base64 grows a file by a third.
export const MAX_UPLOAD_BYTES = 3_000_000;
const ITEM_ID = /^[A-Za-z0-9!_-]{1,256}$/;
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const ITEM_FIELDS =
  "id,name,size,file,folder,webUrl,lastModifiedDateTime,parentReference,remoteItem";
const TEXT_TYPES =
  /^(text\/|application\/(json|xml|csv|x-yaml|yaml|javascript|sql)\b|[^;]*\+(json|xml)\b)/;

export type DriveItem = {
  id: string;
  name?: string;
  size?: number;
  file?: { mimeType?: string };
  folder?: { childCount?: number };
  webUrl?: string;
  lastModifiedDateTime?: string;
  parentReference?: { siteId?: string; driveId?: string; path?: string };
  // Set on a shortcut to an item elsewhere, possibly in another site.
  remoteItem?: unknown;
};

export async function ok(
  ctx: ToolContext,
  mark: (s: number) => void,
  req: Parameters<ToolContext["call"]>[0],
) {
  const res = await ctx.call(req);
  mark(res.status);
  if (res.status < 200 || res.status >= 300) {
    const message = (res.data as { error?: { message?: string } } | null)?.error?.message;
    throw new VendorFailed(
      res.status,
      `Microsoft Graph answered ${res.status}${message ? `: ${message}` : ""}`,
    );
  }
  return res.data as any;
}

// The key's sites, kept to those the connection was made with: a stale or typed-in address
// outside them is never reached, even if the app holds a wider Sites permission.
export function grantedSites(ctx: ToolContext): string[] {
  const sites = listLimit(ctx.limits, "sites");
  if (!sites)
    throw new Denied("No sites are set for this agent. Ask an administrator to add some.");
  const connected = (ctx.config.sites ?? "").split("\n").filter(Boolean);
  return sites.map((s) => normalizeSiteUrl(s)).filter((s) => connected.includes(s));
}

function checkSite(ctx: ToolContext, site: string): string {
  let normalized: string;
  try {
    normalized = normalizeSiteUrl(site);
  } catch (error) {
    if (error instanceof ConnectError)
      throw new Denied("Give a site address from graph_list_sites");
    throw error;
  }
  if (!grantedSites(ctx).includes(normalized))
    throw new Denied(`This agent is not allowed to use ${site}`);
  return normalized;
}

// A path inside a site's document library, as its segments. Refused: . and .., characters
// SharePoint does not allow in names, % (no encoded path ever reaches Graph), control
// characters, and names SharePoint would trim (outer spaces, a trailing dot).
// eslint-disable-next-line no-control-regex -- control characters are what it refuses
const BAD_NAME = /[\\:?#*"<>|%\u0000-\u001f\u007f]/;
const badName = (p: string) =>
  p === "." ||
  p === ".." ||
  BAD_NAME.test(p) ||
  p !== p.trim() ||
  p.endsWith(".") ||
  p.length > 255;

export function pathSegments(raw: string): string[] {
  const parts = raw.split("/").filter(Boolean);
  if (parts.some(badName)) throw new Denied(`${raw} is not a folder path such as Reports/2026`);
  return parts;
}

/** A folder path inside a site's document library, as Graph path segments. */
export function folderPath(raw: string): string {
  return pathSegments(raw).map(encodeURIComponent).join("/");
}

function fileName(raw: string): string {
  if (!raw || raw.includes("/") || badName(raw))
    throw new Denied("Give a file name without slashes or special characters");
  return encodeURIComponent(raw);
}

// SharePoint compares names without case. Both foldings must agree, so a character JavaScript
// folds into another (the Kelvin sign into k) does not pass for the folder's own name.
const sameName = (a: string, b: string) => {
  const [x, y] = [a.normalize("NFC"), b.normalize("NFC")];
  return x.toLowerCase() === y.toLowerCase() && x.toUpperCase() === y.toUpperCase();
};

/** Whether `path` is `folder` or inside it. */
export function within(path: string[], folder: string[]): boolean {
  return folder.length <= path.length && folder.every((f, i) => sameName(f, path[i]));
}

// The folder the key is kept to in a site (its `folders` limit), null for the whole site.
export function keptFolder(ctx: ToolContext, site: string): string[] | null {
  const folders = mapLimit(ctx.limits, "folders");
  if (!folders) return null;
  const granted = normalizeSiteUrl(site);
  const entry = Object.entries(folders).find(([k]) => normalizeSiteUrl(k) === granted);
  if (!entry) return null;
  const folder = pathSegments(entry[1]);
  return folder.length ? folder : null;
}

/** The folder a call asks for, as segments: by default the key's folder, never outside it. */
export function targetFolder(ctx: ToolContext, site: string, raw: string): string[] {
  const path = pathSegments(raw);
  const kept = keptFolder(ctx, checkSite(ctx, site));
  if (!kept) return path;
  if (!path.length) return kept;
  if (!within(path, kept))
    throw new Denied(`This agent may only use the ${kept.join("/")} folder in ${site}`);
  return path;
}

export type OpenSite = { id: string; folder: string[] | null };

// The site's Graph id, read back by address (which must still be the one granted), and the
// folder the key is kept to there.
export async function openSite(
  ctx: ToolContext,
  mark: (s: number) => void,
  site: string,
): Promise<OpenSite> {
  const granted = checkSite(ctx, site);
  const folder = keptFolder(ctx, granted);
  const found = await ok(ctx, mark, {
    method: "GET",
    path: `${sitePath(granted)}?$select=id,webUrl`,
  });
  if (!found?.id || normalizeSiteUrl(found.webUrl ?? "") !== granted)
    throw new Denied(`${site} did not resolve to the site this agent may use`);
  return { id: found.id as string, folder };
}

// Whether an item is in the site Graph names `id` ("host,site collection id,web id"). An item's
// parentReference.siteId is the site collection id alone (seen live), or the whole id.
export function inSite(item: DriveItem, id: string): boolean {
  const own = item.parentReference?.siteId?.toLowerCase();
  return Boolean(own) && (own === id.toLowerCase() || own === id.split(",")[1]?.toLowerCase());
}

// Graph ids are opaque: an item id from another site can be asked for under this site's
// path, so the item's own site is read back and must match.
export function checkItemSite(item: DriveItem, id: string) {
  if (!inSite(item, id)) throw new Denied("That file is not in a site this agent may use");
}

/** The folder an item sits in, from its parent's percent-encoded path ("/drive/root:/A/B"). */
export function itemFolder(item: DriveItem): string[] {
  const match = /^\/drives?(?:\/[^/]+)?\/root:(.*)$/.exec(item.parentReference?.path ?? "");
  if (!match) throw new Denied("Could not tell which folder that file is in");
  // Split before decoding, so an encoded slash stays inside its name.
  try {
    return match[1].split("/").filter(Boolean).map(decodeURIComponent);
  } catch {
    throw new Denied("Could not tell which folder that file is in");
  }
}

// An item read by id must be in the site, in the site's own document library (its default
// drive, the only one paths reach), and inside the key's folder there.
export async function checkItem(
  ctx: ToolContext,
  mark: (s: number) => void,
  site: OpenSite,
  item: DriveItem,
) {
  checkItemSite(item, site.id);
  if (item.remoteItem) throw new Denied("That is a shortcut to something outside this site");
  const drive = await ok(ctx, mark, {
    method: "GET",
    path: `/v1.0/sites/${encodeURIComponent(site.id)}/drive?$select=id`,
  });
  if (!drive?.id || item.parentReference?.driveId !== drive.id)
    throw new Denied("That file is not in this site's document library");
  if (site.folder && !within(itemFolder(item), site.folder))
    throw new Denied(`This agent may only use the ${site.folder.join("/")} folder in this site`);
}

/** Where a folder listing continues: Graph's $skiptoken from its next page link. */
export const cursorArg = z
  .string()
  .max(4000)
  .optional()
  .describe("The `next` value from the previous call with the same arguments: the following page.");
export const nextCursor = (data: { "@odata.nextLink"?: string }) => {
  const link = data["@odata.nextLink"];
  return link ? (URL.parse(link)?.searchParams.get("$skiptoken") ?? null) : null;
};
export const pageQuery = (top: number, cursor: string | undefined) =>
  `$top=${top}${cursor ? `&$skiptoken=${encodeURIComponent(cursor)}` : ""}`;

function summary(item: DriveItem) {
  return {
    id: item.id,
    name: item.name ?? "",
    kind: item.folder ? "folder" : "file",
    size: item.size ?? null,
    type: item.file?.mimeType ?? null,
    modified: item.lastModifiedDateTime ?? null,
    url: item.webUrl ?? null,
  };
}

// Buffer.from(x, "base64") skips characters it does not know, which would upload a
// corrupted file without an error.
export function decodeBase64(raw: string): Buffer {
  const compact = raw.replace(/\s+/g, "");
  if (compact.length % 4 !== 0 || !/^[A-Za-z0-9+/]*={0,2}$/.test(compact))
    throw new Denied("content is not valid base64");
  return Buffer.from(compact, "base64");
}

export function isText(mimeType: string | undefined): boolean {
  return Boolean(mimeType && TEXT_TYPES.test(mimeType.toLowerCase()));
}

export const graphTools: Toolset = {
  names: [
    "graph_list_sites",
    "graph_list_files",
    "graph_read_file",
    "graph_upload_file",
    "graph_list_users",
  ],

  register(server, ctx, allowed) {
    const readOnly = { readOnlyHint: true, openWorldHint: true } as const;
    const site = z.string().max(500).describe("Site address from graph_list_sites.");
    const folderArg = z
      .string()
      .max(1000)
      .default("")
      .describe("Folder path, e.g. Reports/2026. Empty: the top, or this agent's folder.");

    if (allowed.has("graph_list_sites"))
      server.registerTool(
        "graph_list_sites",
        {
          description:
            "List the SharePoint sites this agent may use. only_folder: the one folder it may use there.",
          inputSchema: z.object({}),
          annotations: readOnly,
        },
        () =>
          audited(ctx, "graph_list_sites", null, async (mark) => {
            const sites = [];
            for (const url of grantedSites(ctx)) {
              const res = await ctx.call({
                method: "GET",
                path: `${sitePath(url)}?$select=id,displayName,description,webUrl`,
              });
              mark(res.status);
              const data = res.data as { displayName?: string; description?: string } | null;
              sites.push(
                res.status === 200
                  ? {
                      site: url,
                      name: data?.displayName ?? "",
                      description: data?.description ?? "",
                      ...(keptFolder(ctx, url)
                        ? { only_folder: keptFolder(ctx, url)!.join("/") }
                        : {}),
                    }
                  : { site: url, unavailable: `Microsoft Graph answered ${res.status}` },
              );
            }
            return json({ sites });
          }),
      );

    if (allowed.has("graph_list_files"))
      server.registerTool(
        "graph_list_files",
        {
          description:
            "List files and folders in a site's document library, at its top or in a folder. When `more` is true, call again with `cursor` set to `next` for the following page.",
          inputSchema: z.object({
            site,
            folder: folderArg,
            max_results: z.number().int().min(1).max(200).default(50),
            cursor: cursorArg,
          }),
          annotations: readOnly,
        },
        ({ site: s, folder, max_results, cursor }) =>
          audited(ctx, "graph_list_files", `${s} ${folder}`.trim(), async (mark) => {
            const path = targetFolder(ctx, s, folder).map(encodeURIComponent).join("/");
            const { id } = await openSite(ctx, mark, s);
            const base = `/v1.0/sites/${encodeURIComponent(id)}/drive/root`;
            const data = await ok(ctx, mark, {
              method: "GET",
              path: `${path ? `${base}:/${path}:` : base}/children?${pageQuery(max_results, cursor)}&$select=${ITEM_FIELDS}`,
            });
            const items = (data.value as DriveItem[]).filter((i) => !i.remoteItem && inSite(i, id));
            return json({
              items: items.map(summary),
              more: Boolean(data["@odata.nextLink"]),
              next: nextCursor(data),
            });
          }),
      );

    if (allowed.has("graph_read_file"))
      server.registerTool(
        "graph_read_file",
        {
          description: `Read a text file (up to ${MAX_READ_BYTES / 1_000_000} MB) from a site's document library. Other files return their details only.`,
          inputSchema: z.object({
            site,
            item_id: z.string().max(300).describe("File id from graph_list_files."),
          }),
          annotations: readOnly,
        },
        ({ site: s, item_id }) =>
          audited(ctx, "graph_read_file", `${s} ${item_id}`, async (mark) => {
            if (!ITEM_ID.test(item_id)) throw new Denied("Give a file id from graph_list_files");
            const opened = await openSite(ctx, mark, s);
            const path = `/v1.0/sites/${encodeURIComponent(opened.id)}/drive/items/${encodeURIComponent(item_id)}`;
            const item: DriveItem = await ok(ctx, mark, {
              method: "GET",
              path: `${path}?$select=${ITEM_FIELDS}`,
            });
            await checkItem(ctx, mark, opened, item);
            if (item.folder) throw new Denied("That is a folder. Use graph_list_files.");
            const details = summary(item);
            if (!isText(item.file?.mimeType))
              return json({
                ...details,
                content: null,
                note: "Not a text file, so it is not downloaded.",
              });
            if ((item.size ?? Infinity) > MAX_READ_BYTES)
              return json({ ...details, content: null, note: "Too large to read here." });
            const content = await ok(ctx, mark, {
              method: "GET",
              path: `${path}/content`,
              responseType: "text",
            });
            return json({ ...details, content: String(content ?? "") });
          }),
      );

    if (allowed.has("graph_upload_file"))
      server.registerTool(
        "graph_upload_file",
        {
          description: `Add a file (up to ${MAX_UPLOAD_BYTES / 1_000_000} MB) to a site's document library, or replace one when replace is true.`,
          inputSchema: z.object({
            site,
            folder: folderArg,
            name: z.string().max(255).describe("File name, e.g. summary.md."),
            content: z.string().max(4_100_000),
            encoding: z.enum(["text", "base64"]).default("text"),
            replace: z.boolean().default(false).describe("Replace a file of the same name."),
          }),
          annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
        },
        ({ site: s, folder, name, content, encoding, replace }) =>
          audited(ctx, "graph_upload_file", `${s} ${folder}/${name}`, async (mark, retarget) => {
            const path = [
              ...targetFolder(ctx, s, folder).map(encodeURIComponent),
              fileName(name),
            ].join("/");
            const bytes = new Uint8Array(
              encoding === "base64" ? decodeBase64(content) : Buffer.from(content, "utf8"),
            );
            if (bytes.byteLength > MAX_UPLOAD_BYTES)
              throw new Denied(`Files up to ${MAX_UPLOAD_BYTES / 1_000_000} MB can be added`);
            const { id } = await openSite(ctx, mark, s);
            const conflict = replace ? "replace" : "fail";
            const item: DriveItem = await ok(ctx, mark, {
              method: "PUT",
              path: `/v1.0/sites/${encodeURIComponent(id)}/drive/root:/${path}:/content?@microsoft.graph.conflictBehavior=${conflict}`,
              body: bytes,
            });
            retarget(`${s} ${item.id}`);
            return json(summary(item));
          }),
      );

    if (allowed.has("graph_list_users"))
      server.registerTool(
        "graph_list_users",
        {
          description:
            "Find people in the organization by name, or list a group's members. Returns names, email addresses and job titles. When `more` is true, call again with `cursor` set to `next`.",
          inputSchema: z.object({
            search: z.string().max(100).optional().describe("Part of a name or email address."),
            group_id: z.string().max(36).optional().describe("A group's id, to list its members."),
            max_results: z.number().int().min(1).max(100).default(25),
            cursor: cursorArg.describe(
              "The `next` value from the previous call with the same search and group: the following page.",
            ),
          }),
          annotations: readOnly,
        },
        ({ search, group_id, max_results, cursor }) =>
          audited(ctx, "graph_list_users", group_id ?? search ?? null, async (mark) => {
            if (group_id !== undefined && !GUID.test(group_id))
              throw new Denied("Give a group id, a GUID");
            const fields = "id,displayName,mail,jobTitle,department";
            const query = [`$select=${fields}`, pageQuery(max_results, cursor)];
            const words = search?.replaceAll('"', "").trim();
            if (words)
              query.push(
                `$search=${encodeURIComponent(`"displayName:${words}" OR "mail:${words}"`)}`,
              );
            const base = group_id
              ? `/v1.0/groups/${group_id}/members/microsoft.graph.user`
              : "/v1.0/users";
            const data = await ok(ctx, mark, {
              method: "GET",
              path: `${base}?${query.join("&")}`,
              // $search needs this header.
              headers: words ? { consistencylevel: "eventual" } : {},
            });
            const users = (data.value as Record<string, string | null>[]).map((u) => ({
              id: u.id,
              name: u.displayName ?? null,
              email: u.mail ?? null,
              title: u.jobTitle ?? null,
              department: u.department ?? null,
            }));
            return json({ users, more: Boolean(data["@odata.nextLink"]), next: nextCursor(data) });
          }),
      );
  },
};
