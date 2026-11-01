import assert from "node:assert/strict";
import { test } from "node:test";
import { graphTools, MAX_SESSION_PUT_BYTES, sharePointLink } from "./tools.ts";
import { loadConnectors } from "../../packages/connector-kit/src/connector.ts";
import type { ToolContext } from "../../packages/connector-kit/src/common.ts";
import type { Limits, VendorRequest, VendorResponse } from "../../packages/connector-kit/src/types.ts";

// The download and upload link tools against a fake Graph. Site "editorial" (site collection
// SC1, library DRIVE1) holds ITEM1, a Word file in "Copyedit In", and FOLDER1; ITEM2 is a file
// in another site collection; ITEM3 sits in "Archive".

const SITE = "https://contoso.sharepoint.com/sites/editorial";
const SITE_ID = "contoso.sharepoint.com,SC1,WEB1";
const BASE = `/v1.0/sites/${encodeURIComponent(SITE_ID)}`;
const DOWNLOAD = "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=1&tempauth=t";
const UPLOAD = "https://contoso.sharepoint.com/_api/v2.0/drives/DRIVE1/items/root/uploadSession?guid=g";

type Item = Record<string, unknown>;
const file = (id: string, folder: string, extra: Item = {}): Item => ({
  id,
  name: `${id}.docx`,
  size: 48_213,
  file: {
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  },
  webUrl: `${SITE}/Shared%20Documents/${id}.docx`,
  lastModifiedDateTime: "2026-10-30T09:00:00Z",
  parentReference: { siteId: "SC1", driveId: "DRIVE1", path: `/drive/root:/${folder}` },
  "@microsoft.graph.downloadUrl": DOWNLOAD,
  ...extra,
});

function setup(
  opts: { limits?: Limits; items?: Record<string, Item>; uploadUrl?: unknown; sites?: string } = {},
) {
  const items: Record<string, Item> = opts.items ?? {
    ITEM1: file("ITEM1", "Copyedit%20In"),
    ITEM2: file("ITEM2", "Copyedit%20In", {
      parentReference: { siteId: "SC9", driveId: "DRIVE9", path: "/drive/root:/Copyedit%20In" },
    }),
    ITEM3: file("ITEM3", "Archive"),
    ITEM4: file("ITEM4", "Copyedit%20In", {
      parentReference: { siteId: "SC1", driveId: "DRIVE2", path: "/drive/root:/Copyedit%20In" },
    }),
    ITEM5: file("ITEM5", "Copyedit%20In", { remoteItem: { id: "X", parentReference: { siteId: "SC9" } } }),
    FOLDER1: { id: "FOLDER1", name: "Copyedit In", folder: { childCount: 2 }, parentReference: { siteId: "SC1", driveId: "DRIVE1", path: "/drive/root:" } },
  };
  const requests: VendorRequest[] = [];
  const audit: { tool: string; target: string | null }[] = [];
  const call = async (req: VendorRequest): Promise<VendorResponse> => {
    requests.push(req);
    const [path] = req.path.split("?");
    if (path === "/v1.0/sites/contoso.sharepoint.com:/sites/editorial")
      return { status: 200, data: { id: SITE_ID, webUrl: SITE } };
    if (path === `${BASE}/drive`) return { status: 200, data: { id: "DRIVE1" } };
    const item = /\/drive\/items\/([^/?]+)$/.exec(path);
    if (item)
      return items[item[1]]
        ? { status: 200, data: items[item[1]] }
        : { status: 404, data: { error: { message: "Item not found" } } };
    if (req.method === "POST" && path.endsWith(":/createUploadSession"))
      return {
        status: 200,
        data: {
          uploadUrl: "uploadUrl" in opts ? opts.uploadUrl : UPLOAD,
          expirationDateTime: "2026-11-01T12:00:00Z",
        },
      };
    return { status: 500, data: {} };
  };
  const ctx = {
    limits: opts.limits ?? { sites: [SITE] },
    config: { sites: opts.sites ?? SITE },
    call,
    async audit(e: { tool: string; target: string | null }) {
      audit.push({ tool: e.tool, target: e.target });
    },
  } as unknown as ToolContext;
  const tools = new Map<string, (args: unknown) => Promise<any>>();
  const server = {
    registerTool: (name: string, def: any, fn: (args: unknown) => Promise<any>) => {
      tools.set(name, (args: unknown) => fn(def.inputSchema.parse(args)));
    },
  };
  graphTools.register(server as never, ctx, new Set(graphTools.names));
  const run = async (name: string, args: object) => {
    const r = await tools.get(name)!(args);
    const text = r.content[0].text as string;
    return { isError: Boolean(r.isError), text, body: r.isError ? null : JSON.parse(text) };
  };
  return { run, requests, audit };
}

test("every SharePoint tool is in exactly one capability, the link tools in read and write files", () => {
  const sp = loadConnectors().find((c) => c.id === "sharepoint-files")!;
  const listed = sp.capabilities.flatMap((c) => c.tools);
  const fileTools = graphTools.names;
  for (const t of fileTools) assert.equal(listed.filter((x) => x === t).length, 1, t);
  const of = (id: string) => sp.capabilities.find((c) => c.id === id)!.tools;
  assert.ok(of("read_files").includes("graph_get_download_link"));
  assert.ok(of("write_files").includes("graph_create_upload_link"));
});

test("graph_get_download_link answers the file's details and its SharePoint download link", async () => {
  const { run, requests, audit } = setup();
  const r = await run("graph_get_download_link", { site: SITE, item_id: "ITEM1" });
  assert.equal(r.isError, false, r.text);
  assert.equal(r.body.download_url, DOWNLOAD);
  assert.equal(r.body.name, "ITEM1.docx");
  assert.equal(r.body.kind, "file");
  assert.equal(r.body.expires_in_minutes, 60);
  const get = requests.find((q) => q.path.includes("/drive/items/ITEM1"))!;
  assert.ok(!get.path.includes("$select"), "the link is only in the default answer");
  // The audit row names the file, never the link.
  assert.equal(audit.at(-1)!.target, `${SITE} ITEM1`);
  assert.ok(!JSON.stringify(audit).includes("tempauth"));
});

test("graph_get_download_link refuses a folder, another site's or library's file, a shortcut and a bad id", async () => {
  const { run, requests } = setup();
  const folder = await run("graph_get_download_link", { site: SITE, item_id: "FOLDER1" });
  assert.equal(folder.isError, true);
  assert.match(folder.text, /folder/);
  const other = await run("graph_get_download_link", { site: SITE, item_id: "ITEM2" });
  assert.equal(other.isError, true);
  assert.match(other.text, /not in a site this agent may use/);
  const before = requests.length;
  const bad = await run("graph_get_download_link", { site: SITE, item_id: "../x" });
  assert.equal(bad.isError, true);
  assert.match(bad.text, /Give a file id from graph_list_files/);
  assert.equal(requests.length, before, "refused before any call");
  const library = await run("graph_get_download_link", { site: SITE, item_id: "ITEM4" });
  assert.equal(library.isError, true);
  assert.match(library.text, /not in this site's document library/);
  const shortcut = await run("graph_get_download_link", { site: SITE, item_id: "ITEM5" });
  assert.equal(shortcut.isError, true);
  assert.match(shortcut.text, /shortcut/);
});

test("graph_get_download_link keeps to the key's folder and sites", async () => {
  const kept = setup({ limits: { sites: [SITE], folders: { [SITE]: "Copyedit In" } } });
  assert.equal((await kept.run("graph_get_download_link", { site: SITE, item_id: "ITEM1" })).isError, false);
  const outside = await kept.run("graph_get_download_link", { site: SITE, item_id: "ITEM3" });
  assert.equal(outside.isError, true);
  assert.match(outside.text, /only use the Copyedit In folder/);
  const other = await kept.run("graph_get_download_link", {
    site: "https://contoso.sharepoint.com/sites/finance",
    item_id: "ITEM1",
  });
  assert.equal(other.isError, true);
});

test("graph_get_download_link hands out only an https SharePoint link", async () => {
  for (const bad of [
    "http://contoso.sharepoint.com/download?x=1",
    "https://evil.example.com/contoso.sharepoint.com",
    "https://contoso.sharepoint.com.evil.example/x",
    undefined,
  ]) {
    const { run } = setup({
      items: { ITEM1: file("ITEM1", "Copyedit%20In", { "@microsoft.graph.downloadUrl": bad }) },
    });
    const r = await run("graph_get_download_link", { site: SITE, item_id: "ITEM1" });
    assert.equal(r.isError, true, String(bad));
    assert.match(r.text, /did not give a download link/);
  }
});

test("graph_create_upload_link opens a session that never replaces unless asked", async () => {
  const { run, requests } = setup();
  const r = await run("graph_create_upload_link", {
    site: SITE,
    folder: "Copyedit Out",
    name: "US12345red.docx",
  });
  assert.equal(r.isError, false, r.text);
  assert.deepEqual(r.body, {
    upload_url: UPLOAD,
    expires: "2026-11-01T12:00:00Z",
    max_bytes: MAX_SESSION_PUT_BYTES,
  });
  const post = requests.find((q) => q.method === "POST")!;
  assert.equal(post.path, `${BASE}/drive/root:/Copyedit%20Out/US12345red.docx:/createUploadSession`);
  assert.deepEqual(post.body, { item: { "@microsoft.graph.conflictBehavior": "fail" } });

  const again = setup();
  await again.run("graph_create_upload_link", {
    site: SITE,
    folder: "Copyedit Out",
    name: "US12345red.docx",
    replace: true,
  });
  const replace = again.requests.find((q) => q.method === "POST")!;
  assert.deepEqual(replace.body, { item: { "@microsoft.graph.conflictBehavior": "replace" } });
});

test("graph_create_upload_link refuses names and folders outside the key before any call", async () => {
  const { run, requests } = setup({
    limits: { sites: [SITE], folders: { [SITE]: "Copyedit Out" } },
  });
  const outside = await run("graph_create_upload_link", { site: SITE, folder: "Finance", name: "a.docx" });
  assert.equal(outside.isError, true);
  assert.match(outside.text, /only use the Copyedit Out folder/);
  const slash = await run("graph_create_upload_link", { site: SITE, folder: "", name: "a/b.docx" });
  assert.equal(slash.isError, true);
  const dots = await run("graph_create_upload_link", { site: SITE, folder: "Copyedit Out/..", name: "a.docx" });
  assert.equal(dots.isError, true);
  assert.equal(requests.length, 0);
  // An empty folder means the key's own folder.
  const ok = await run("graph_create_upload_link", { site: SITE, folder: "", name: "a.docx" });
  assert.equal(ok.isError, false, ok.text);
  assert.match(requests.find((q) => q.method === "POST")!.path, /root:\/Copyedit%20Out\/a\.docx:/);
});

test("graph_create_upload_link refuses an upload link that is not https SharePoint", async () => {
  const { run } = setup({ uploadUrl: "https://evil.example.com/upload" });
  const r = await run("graph_create_upload_link", { site: SITE, folder: "Out", name: "a.docx" });
  assert.equal(r.isError, true);
  assert.match(r.text, /did not give an upload link/);
});

test("sharePointLink accepts only https links on a SharePoint host", () => {
  assert.equal(sharePointLink(DOWNLOAD), DOWNLOAD);
  assert.equal(sharePointLink("https://contoso-my.sharepoint.com/personal/a/x"), "https://contoso-my.sharepoint.com/personal/a/x");
  assert.equal(sharePointLink("https://agency.sharepoint.us/x"), null);
  assert.equal(sharePointLink("https://contoso.sharepoint.com.@evil.example/x"), null);
  assert.equal(sharePointLink("https://sharepoint.com.evil.example/x"), null);
  assert.equal(sharePointLink("ftp://contoso.sharepoint.com/x"), null);
  assert.equal(sharePointLink(42), null);
});

test("a site in the key but no longer connected is never reached", async () => {
  const { run, requests } = setup({ sites: "https://contoso.sharepoint.com/sites/other" });
  const down = await run("graph_get_download_link", { site: SITE, item_id: "ITEM1" });
  assert.equal(down.isError, true);
  const up = await run("graph_create_upload_link", { site: SITE, folder: "Out", name: "a.docx" });
  assert.equal(up.isError, true);
  assert.match(up.text, /not allowed to use/);
  assert.equal(requests.length, 0);
});

test("graph_create_upload_link audits the folder opened and a replace, never the link", async () => {
  const kept = setup({ limits: { sites: [SITE], folders: { [SITE]: "Copyedit Out" } } });
  await kept.run("graph_create_upload_link", { site: SITE, folder: "", name: "US1red.docx" });
  assert.equal(kept.audit.at(-1)!.target, `${SITE} Copyedit Out/US1red.docx`);
  await kept.run("graph_create_upload_link", { site: SITE, folder: "", name: "US1red.docx", replace: true });
  assert.equal(kept.audit.at(-1)!.target, `${SITE} Copyedit Out/US1red.docx (replace)`);
  assert.ok(!JSON.stringify(kept.audit).includes("uploadSession"));
});

test("the upload limit is under Microsoft's 60 MiB per request", () => {
  assert.equal(MAX_SESSION_PUT_BYTES, 60 * 1024 * 1024 - 1);
});
