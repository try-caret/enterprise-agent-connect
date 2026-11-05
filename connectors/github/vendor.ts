import { createHash, sign } from "node:crypto";
import { MAX_BINARY_BYTES } from "../../packages/connector-kit/src/common.js";
import { vendorFetch, vendorUrl } from "../../packages/connector-kit/src/fetch.js";
import { rsaKey, vendorError } from "../../packages/connector-kit/src/oauth.js";
import {
  ConnectError,
  TokenCache,
  type Fetch,
  type Vendor,
  type VendorConnection,
  type VendorRequest,
  type VendorResponse,
} from "../../packages/connector-kit/src/types.js";
import { guardedFetch } from "../http/index.js";

// A GitHub App the organization owns, installed on its account. The service signs a short JWT
// with the app's private key (RS256, iss = the app's client ID, 9 minutes, iat 60 s back for
// clock drift) and trades it at POST /app/installations/{id}/access_tokens for an installation
// token (an hour, only the installation's repositories and permissions), kept until a minute
// before it expires and minted again once after a 401. GitHub.com only (api.github.com).
// A repository archive answers 302 to codeload.github.com with its own short-lived token in the
// address; that one redirect is followed, without the installation token, and nowhere else.
// Sources: docs.github.com apps/creating-github-apps/authenticating-with-a-github-app
// (generating-a-json-web-token-jwt-for-a-github-app, generating-an-installation-access-token-
// for-a-github-app), rest/apps/apps (get an installation), rest/repos/contents (download a
// repository archive (tar)), read November 5, 2026.

const API = "https://api.github.com";
const CODELOAD = "https://codeload.github.com";
const CLIENT_ID = /^[A-Za-z0-9.]{1,64}$/;
const INSTALLATION_ID = /^\d{1,20}$/;
const HEADERS = {
  accept: "application/vnd.github+json",
  "x-github-api-version": "2022-11-28",
  "user-agent": "enterprise-agent-connect",
};

export function appJwt(clientId: string, privateKey: string, now = Date.now()): string {
  const part = (v: object) => Buffer.from(JSON.stringify(v)).toString("base64url");
  const iat = Math.floor(now / 1000) - 60;
  const signed = `${part({ alg: "RS256", typ: "JWT" })}.${part({ iat, exp: iat + 600, iss: clientId })}`;
  return `${signed}.${sign("sha256", Buffer.from(signed), rsaKey(privateKey)).toString("base64url")}`;
}

function parse(raw: string): unknown {
  try {
    return JSON.parse(raw);
  } catch {
    return raw;
  }
}

const githubError = (status: number, data: unknown) => {
  const detail = vendorError(data);
  return `GitHub answered ${status}${detail ? `: ${detail}` : ""}`;
};

export function githubVendor(unguarded: Fetch = fetch): Vendor {
  const fetchImpl = vendorFetch("GitHub", unguarded);
  const codeload = guardedFetch([CODELOAD], unguarded, { types: /(?:)/, maxBytes: MAX_BINARY_BYTES });
  const tokens = new TokenCache();

  const asApp = async (c: VendorConnection, method: "GET" | "POST", path: string) => {
    const res = await fetchImpl(vendorUrl(API, path), {
      method,
      headers: { ...HEADERS, authorization: `Bearer ${appJwt(c.config.client_id, c.secret.private_key)}` },
    });
    const raw = await res.text();
    return { status: res.status, data: raw ? parse(raw) : null };
  };

  async function mint(c: VendorConnection) {
    const res = await asApp(c, "POST", `/app/installations/${c.config.installation_id}/access_tokens`);
    const body = res.data as { token?: unknown; expires_at?: unknown } | null;
    if (res.status < 200 || res.status >= 300 || typeof body?.token !== "string")
      throw new ConnectError(`GitHub refused the installation token. ${githubError(res.status, res.data)}`);
    const expiresIn = Math.floor((Date.parse(String(body.expires_at)) - Date.now()) / 1000);
    if (!(expiresIn > 60)) throw new ConnectError("GitHub answered a token without a usable lifetime");
    return { token: body.token, expiresIn };
  }

  const cacheKey = (c: VendorConnection) =>
    createHash("sha256")
      .update([c.config.client_id, c.config.installation_id, c.secret.private_key].join("\n"))
      .digest("hex");

  async function send(token: string, request: VendorRequest): Promise<VendorResponse> {
    const init = {
      method: request.method,
      headers: { ...HEADERS, authorization: `Bearer ${token}`, ...request.headers },
    };
    if (request.responseType !== "bytes") {
      const res = await fetchImpl(vendorUrl(API, request.path), init);
      const raw = await res.text();
      const data = request.responseType === "text" && res.ok ? raw : raw ? parse(raw) : null;
      return { status: res.status, data, headers: Object.fromEntries(res.headers) };
    }
    const res = await unguarded(vendorUrl(API, request.path), { ...init, redirect: "manual" });
    if (res.status < 300 || res.status >= 400) {
      const raw = await res.text();
      if (res.ok) throw new ConnectError("GitHub answered the download without its redirect");
      return { status: res.status, data: raw ? parse(raw) : null };
    }
    await res.body?.cancel();
    const to = URL.parse(res.headers.get("location") ?? "");
    if (to?.origin !== CODELOAD) throw new ConnectError("GitHub redirected the download somewhere unexpected");
    const file = await codeload(to.href);
    if (!file.ok) return { status: file.status, data: parse(await file.text()) };
    return {
      status: file.status,
      data: new Uint8Array(await file.arrayBuffer()),
      headers: Object.fromEntries(file.headers),
    };
  }

  return {
    async connect(fields) {
      const config = {
        client_id: fields.client_id?.trim() ?? "",
        installation_id: fields.installation_id?.trim() ?? "",
      };
      const secret = { private_key: fields.private_key?.trim() ?? "" };
      if (!CLIENT_ID.test(config.client_id))
        throw new ConnectError("Enter the app's Client ID, shown on its General page");
      if (!INSTALLATION_ID.test(config.installation_id))
        throw new ConnectError("Enter the installation ID, the number at the end of the installation's address");
      rsaKey(secret.private_key);
      const conn = { config, secret };
      const res = await asApp(conn, "GET", `/app/installations/${config.installation_id}`);
      if (res.status === 401)
        throw new ConnectError("GitHub did not accept the private key for this Client ID");
      if (res.status === 404)
        throw new ConnectError("GitHub found no installation of this app with that installation ID");
      if (res.status !== 200) throw new ConnectError(`GitHub sign-in failed. ${githubError(res.status, res.data)}`);
      const install = res.data as {
        account?: { login?: unknown };
        permissions?: { contents?: unknown };
      };
      const owner = install.account?.login;
      if (typeof owner !== "string") throw new ConnectError("GitHub answered the installation without its account");
      if (install.permissions?.contents !== "read" && install.permissions?.contents !== "write")
        throw new ConnectError(
          "The app cannot read code. Give it the Contents repository permission (Read-only), then accept the new permission on the installation.",
        );
      const connected = { config: { ...config, owner }, secret };
      await tokens.get(cacheKey(connected), () => mint(connected));
      return { ...connected, accountLabel: owner };
    },

    async call(c, request) {
      const key = cacheKey(c);
      const token = () => tokens.get(key, () => mint(c));
      let res = await send(await token(), request);
      if (res.status === 401) {
        tokens.forget(key);
        res = await send(await token(), request);
      }
      return res;
    },
  };
}
