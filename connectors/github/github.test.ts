import assert from "node:assert/strict";
import { createHash, createPublicKey, generateKeyPairSync, verify } from "node:crypto";
import { test } from "node:test";
import type { McpServer } from "@modelcontextprotocol/server";
import type { ZodType } from "zod";
import { MAX_BINARY_BYTES, type ToolContext } from "../../packages/connector-kit/src/common.ts";
import { loadConnectors } from "../../packages/connector-kit/src/connector.ts";
import type { Limits, VendorConnection } from "../../packages/connector-kit/src/types.ts";
import { assertToolsImplemented } from "../../packages/server/src/service.ts";
import { githubTools } from "./tools.ts";
import { appJwt, githubVendor } from "./vendor.ts";

// The GitHub App sign-in and the code tools against a fake api.github.com and codeload.

const key = generateKeyPairSync("rsa", { modulusLength: 2048 }).privateKey;
// As GitHub downloads it: PKCS#1.
const PEM = key.export({ type: "pkcs1", format: "pem" }) as string;
const SHA = "0123456789abcdef0123456789abcdef01234567";
const ARCHIVE = Buffer.from([0x1f, 0x8b, 0x08, 0, 1, 2, 3]);
const conn: VendorConnection = {
  config: { client_id: "Iv23liAbc", installation_id: "77", owner: "acme" },
  secret: { private_key: PEM },
};

type Seen = { method: string; url: string; headers: Record<string, string> };

function fakeGitHub(
  options: { contents?: string; location?: string; archive?: Buffer; expired?: Set<string> } = {},
) {
  const seen: Seen[] = [];
  let minted = 0;
  const fetchImpl = (async (input: string | URL, init?: RequestInit) => {
    const url = new URL(String(input));
    const headers = Object.fromEntries(new Headers(init?.headers));
    seen.push({ method: init?.method ?? "GET", url: url.href, headers });
    const path = url.pathname;
    if (url.origin === "https://codeload.github.com")
      return new Response(options.archive ?? ARCHIVE, { headers: { "content-type": "application/x-gzip" } });
    if (path === "/app/installations/77")
      return Response.json({ account: { login: "acme" }, permissions: { metadata: "read", contents: options.contents ?? "read" } });
    if (path === "/app/installations/77/access_tokens")
      return Response.json(
        { token: `ghs_${++minted}`, expires_at: new Date(Date.now() + 3600_000).toISOString() },
        { status: 201 },
      );
    if (options.expired?.has(headers.authorization)) return Response.json({ message: "Bad credentials" }, { status: 401 });
    if (path === "/repos/acme/api/commits/release/1.2") return new Response(SHA, { headers: { "content-type": "application/vnd.github.sha" } });
    if (path === `/repos/acme/api/git/trees/${SHA}`)
      return Response.json({
        sha: "tree1",
        truncated: false,
        tree: [
          { path: "src", type: "tree" },
          { path: "src/index.ts", type: "blob", size: 20 },
          { path: "srcs.txt", type: "blob", size: 2 },
          { path: "README.md", type: "blob", size: 5 },
        ],
      });
    if (path === "/repos/acme/api/contents/src/index.ts")
      return Response.json({ type: "file", encoding: "base64", content: Buffer.from("export const a = 1;\n").toString("base64"), sha: "b1", size: 20 });
    if (path === "/repos/acme/api/contents/src") return Response.json([{ type: "file", path: "src/index.ts" }]);
    if (path === "/repos/acme/api/contents/logo.png")
      return Response.json({ type: "file", encoding: "base64", content: Buffer.from([0xff, 0xfe, 0x00, 0x89]).toString("base64"), size: 4 });
    if (path === `/repos/acme/api/tarball/${SHA}`)
      return new Response(null, {
        status: 302,
        headers: { location: options.location ?? `https://codeload.github.com/acme/api/legacy.tar.gz/${SHA}?token=short` },
      });
    return Response.json({ message: "Not Found" }, { status: 404 });
  }) as typeof fetch;
  return { seen, fetchImpl, minted: () => minted };
}

async function run(
  fake: ReturnType<typeof fakeGitHub>,
  tool: string,
  args: Record<string, unknown>,
  limits: Limits = {},
) {
  const vendor = githubVendor(fake.fetchImpl);
  const handlers = new Map<string, (a: Record<string, unknown>) => Promise<unknown>>();
  const recorder = {
    registerTool: (name: string, c: { inputSchema: ZodType }, h: (a: unknown) => unknown) =>
      handlers.set(name, async (a) => h(c.inputSchema.parse(a))),
  } as unknown as McpServer;
  const audit: { decision: string; target: string | null }[] = [];
  const ctx = {
    limits,
    config: conn.config,
    call: (r) => vendor.call(conn, r),
    audit: async (e) => void audit.push(e),
  } as ToolContext;
  githubTools.register(recorder, ctx, new Set(githubTools.names));
  const result = (await handlers.get(tool)!(args)) as {
    isError?: boolean;
    content: { text: string; resource?: { blob: string; mimeType: string; uri: string } }[];
  };
  return { result, text: result.content[0].text, audit };
}

test("the GitHub app loads with its tools in one capability", () => {
  const connector = loadConnectors().find((c) => c.id === "github")!;
  assertToolsImplemented([{ connector, toolset: githubTools, vendor: githubVendor() }]);
});

test("the app's JWT is RS256, issued by the client ID, and lives under ten minutes", () => {
  const now = Date.parse("2026-11-05T12:00:00Z");
  const [head, body, signature] = appJwt("Iv23liAbc", PEM, now).split(".");
  assert.deepEqual(JSON.parse(Buffer.from(head, "base64url").toString()), { alg: "RS256", typ: "JWT" });
  const claims = JSON.parse(Buffer.from(body, "base64url").toString());
  assert.equal(claims.iss, "Iv23liAbc");
  assert.equal(claims.iat, now / 1000 - 60);
  assert.ok(claims.exp - claims.iat <= 600 && claims.exp > now / 1000);
  assert.ok(verify("sha256", Buffer.from(`${head}.${body}`), createPublicKey(key), Buffer.from(signature, "base64url")));
});

test("connecting finds the installation's account, needs Contents read, and mints a token", async () => {
  let fake = fakeGitHub();
  const connected = await githubVendor(fake.fetchImpl).connect({
    client_id: " Iv23liAbc ",
    installation_id: "77",
    private_key: PEM.replace(/\n/g, " "),
  });
  assert.equal(connected.accountLabel, "acme");
  assert.deepEqual(connected.config, { client_id: "Iv23liAbc", installation_id: "77", owner: "acme" });
  assert.match(fake.seen[0].headers.authorization, /^Bearer ey/);
  assert.equal(fake.seen[0].headers["x-github-api-version"], "2022-11-28");
  assert.equal(fake.minted(), 1);

  fake = fakeGitHub({ contents: "none" });
  await assert.rejects(
    githubVendor(fake.fetchImpl).connect({ client_id: "Iv23liAbc", installation_id: "77", private_key: PEM }),
    /Contents repository permission/,
  );
  await assert.rejects(
    githubVendor(fake.fetchImpl).connect({ client_id: "Iv23liAbc", installation_id: "acme", private_key: PEM }),
    /installation ID/,
  );
  await assert.rejects(
    githubVendor(fake.fetchImpl).connect({ client_id: "Iv23liAbc", installation_id: "77", private_key: "nope" }),
    /whole private key/,
  );
});

test("calls carry the installation token, kept between calls and minted again after a 401", async () => {
  const fake = fakeGitHub({ expired: new Set(["Bearer ghs_1"]) });
  const out = await run(fake, "github_resolve_ref", { repository: "acme/api", ref: "release/1.2" });
  assert.deepEqual(JSON.parse(out.text), { repository: "acme/api", ref: "release/1.2", sha: SHA });
  assert.equal(fake.minted(), 2);
  const call = fake.seen.at(-1)!;
  assert.equal(call.url, "https://api.github.com/repos/acme/api/commits/release/1.2");
  assert.equal(call.headers.authorization, "Bearer ghs_2");
  assert.equal(call.headers.accept, "application/vnd.github.sha");
});

test("a tree lists what is inside a folder, and a file reads as text", async () => {
  const fake = fakeGitHub();
  let out = await run(fake, "github_list_tree", { repository: "acme/api", ref: SHA, path: "src" });
  assert.equal(fake.seen.at(-1)!.url, `https://api.github.com/repos/acme/api/git/trees/${SHA}?recursive=1`);
  assert.deepEqual(JSON.parse(out.text), {
    sha: "tree1",
    truncated: false,
    entries: [{ path: "src/index.ts", type: "blob", size: 20 }],
  });

  out = await run(fake, "github_read_file", { repository: "acme/api", ref: "main", path: "src/index.ts" });
  assert.equal(fake.seen.at(-1)!.url, "https://api.github.com/repos/acme/api/contents/src/index.ts?ref=main");
  assert.equal(JSON.parse(out.text).text, "export const a = 1;\n");
  out = await run(fake, "github_read_file", { repository: "acme/api", ref: "main", path: "src" });
  assert.match(out.text, /is a folder/);
  out = await run(fake, "github_read_file", { repository: "acme/api", ref: "main", path: "logo.png" });
  assert.match(out.text, /not a text file/);
  out = await run(fake, "github_read_file", { repository: "acme/api", ref: "main", path: "src/../../x" });
  assert.match(out.text, /path is not a valid value/);
});

test("the archive takes a full SHA, follows only codeload's redirect, without the token", async () => {
  let fake = fakeGitHub();
  let out = await run(fake, "github_download_archive", { repository: "acme/api", commit: SHA });
  assert.equal(out.result.isError, undefined);
  const meta = JSON.parse(out.text);
  assert.equal(meta.size, ARCHIVE.length);
  assert.equal(meta.sha256, createHash("sha256").update(ARCHIVE).digest("hex"));
  assert.equal(meta.mimeType, "application/x-gzip");
  assert.deepEqual(Buffer.from(out.result.content[1].resource!.blob, "base64"), ARCHIVE);
  const download = fake.seen.at(-1)!;
  assert.match(download.url, /^https:\/\/codeload\.github\.com\//);
  assert.equal(download.headers.authorization, undefined);
  assert.deepEqual(out.audit.at(-1), { decision: "allowed", target: `acme/api@${SHA}`, tool: "github_download_archive", reason: null, vendorStatus: 200 });

  const calls = fake.seen.length;
  out = await run(fake, "github_download_archive", { repository: "acme/api", commit: "main" });
  assert.match(out.text, /full 40-character commit SHA/);
  assert.equal(fake.seen.length, calls);

  fake = fakeGitHub({ location: "https://evil.example/x.tar.gz" });
  out = await run(fake, "github_download_archive", { repository: "acme/api", commit: SHA });
  assert.match(out.text, /redirected the download somewhere unexpected/);
  assert.ok(fake.seen.every((s) => !s.url.includes("evil")));

  fake = fakeGitHub({ archive: Buffer.alloc(MAX_BINARY_BYTES + 1) });
  out = await run(fake, "github_download_archive", { repository: "acme/api", commit: SHA });
  assert.equal(out.result.isError, true);
  assert.match(out.text, /larger than/);
});

test("a key kept to some repositories reaches only those, in any case", async () => {
  const fake = fakeGitHub();
  let out = await run(fake, "github_resolve_ref", { repository: "acme/other", ref: "main" }, { repositories: ["acme/api"] });
  assert.match(out.text, /outside this agent's limits/);
  out = await run(fake, "github_resolve_ref", { repository: "ACME/API", ref: "release/1.2" }, { repositories: ["acme/api"] });
  assert.equal(out.result.isError, true);
  assert.doesNotMatch(out.text, /limits/);
  out = await run(fake, "github_list_tree", { repository: "../x", ref: "main" });
  assert.match(out.text, /owner\/name/);
});
