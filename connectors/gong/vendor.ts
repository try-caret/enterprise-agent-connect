import { retryAfter, vendorFetch } from "../../packages/connector-kit/src/fetch.js";
import {
  ConnectError,
  type Fetch,
  type Vendor,
  type VendorConnection,
  type VendorRequest,
  type VendorResponse,
} from "../../packages/connector-kit/src/types.js";

// A Gong API access key and secret, sent as Basic auth to the company's own API address
// (https://<id>.api.gong.io, or the shared https://api.gong.io). Gong allows 3 calls a second
// and 10,000 a day per company and answers 429 with Retry-After past either; the kit's
// vendorFetch retries short waits, longer ones go back to the agent. Errors come as
// {requestId, errors: [...]}.
// Sources: help.gong.io "Receive access to the API" and gong.app.gong.io/settings/api/
// documentation (read October 27, 2026).

const HOST = /^([a-z0-9][a-z0-9-]*\.)?api\.gong\.io$/;

export function normalizeBaseUrl(input: string): string {
  const raw = input.trim().toLowerCase();
  const url = URL.parse(/^[a-z]+:\/\//.test(raw) ? raw : `https://${raw}`);
  if (
    !url ||
    url.protocol !== "https:" ||
    url.port ||
    url.username ||
    !HOST.test(url.hostname) ||
    !/^\/?$/.test(url.pathname.replace(/\/v2\/?$/, ""))
  )
    throw new ConnectError(
      "Enter your Gong API address, such as https://us-12345.api.gong.io (an address ending in .api.gong.io)",
    );
  return url.hostname;
}

export function gongError(res: VendorResponse): string {
  if (res.status === 429) {
    const after = (res.data as { retryAfter?: unknown })?.retryAfter;
    return `Gong is limiting how often the app may call it.${typeof after === "number" ? ` Try again in ${after} seconds.` : ""}`;
  }
  const errors = (res.data as { errors?: unknown })?.errors;
  const detail = Array.isArray(errors)
    ? errors.filter((e) => typeof e === "string").join("; ")
    : "";
  return `Gong answered ${res.status}${detail ? `: ${detail.slice(0, 300)}` : ""}`;
}

export function gongVendor(
  unguarded: Fetch = fetch,
  wait?: (seconds: number) => Promise<void>,
): Vendor {
  const fetchImpl = vendorFetch("Gong", unguarded, wait);

  async function send(host: string, secret: Record<string, string>, request: VendorRequest) {
    if (!HOST.test(host)) throw new Error("Gong connection has no API address");
    const auth = Buffer.from(`${secret.access_key}:${secret.access_key_secret}`).toString("base64");
    const res = await fetchImpl(`https://${host}${request.path}`, {
      method: request.method,
      headers: {
        authorization: `Basic ${auth}`,
        accept: "application/json",
        ...(request.body === undefined ? {} : { "content-type": "application/json" }),
      },
      body: request.body === undefined ? undefined : JSON.stringify(request.body),
    });
    const raw = await res.text();
    const headers = Object.fromEntries(res.headers);
    if (res.status === 429)
      return {
        status: 429,
        data: { retryAfter: retryAfter(res.headers.get("retry-after")) },
        headers,
      };
    return { status: res.status, data: raw ? parse(raw) : null, headers };
  }

  return {
    async connect(fields) {
      const host = normalizeBaseUrl(fields.base_url ?? "");
      const secret = {
        access_key: fields.access_key?.trim() ?? "",
        access_key_secret: fields.access_key_secret?.trim() ?? "",
      };
      if (!secret.access_key || !secret.access_key_secret)
        throw new ConnectError("Enter the access key and its secret from Gong's API page");
      const res = await send(host, secret, {
        method: "GET",
        path: "/v2/workspaces",
      });
      if (res.status === 401)
        throw new ConnectError("Gong did not accept the access key and secret");
      if (res.status !== 200) throw new ConnectError(`Gong sign-in failed. ${gongError(res)}`);
      return { config: { host }, secret, accountLabel: host };
    },

    call: (connection: VendorConnection, request) =>
      send(connection.config.host ?? "", connection.secret, request),
  };
}

function parse(raw: string): unknown {
  try {
    return JSON.parse(raw);
  } catch {
    return raw;
  }
}
