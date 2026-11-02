import { retryAfter, vendorFetch } from "../../packages/connector-kit/src/fetch.js";
import { vendorError } from "../../packages/connector-kit/src/oauth.js";
import {
  ConnectError,
  type Fetch,
  type Vendor,
  type VendorRequest,
  type VendorResponse,
} from "../../packages/connector-kit/src/types.js";

// A Lusha API key in the api_key header, against Lusha's v3 API. v2 (GET /v2/person and
// /v2/company) stops answering on November 30, 2026, and its replacements take a JSON list of
// people or companies, which is why this app has code rather than an http block.
// Sources: docs.lusha.com/openapi.json and the V3 migration guide (read November 1, 2026).

export const LUSHA_HOST = "api.lusha.com";
const BASE = `https://${LUSHA_HOST}`;

export function lushaError(res: VendorResponse): string {
  if (res.status === 402) return "Lusha says the key is out of credits";
  if (res.status === 429) {
    const after = (res.data as { retryAfter?: unknown })?.retryAfter;
    const wait = typeof after === "number" ? ` Try again in ${after} seconds.` : "";
    return `Lusha is limiting how often the app may call it.${wait}`;
  }
  const detail = vendorError(res.data);
  return `Lusha answered ${res.status}${detail ? `: ${detail}` : ""}`;
}

export function lushaVendor(unguarded: Fetch = fetch, wait?: (s: number) => Promise<void>): Vendor {
  const fetchImpl = vendorFetch("Lusha", unguarded, wait);

  async function send(key: string, request: VendorRequest): Promise<VendorResponse> {
    const res = await fetchImpl(`${BASE}${request.path}`, {
      method: request.method,
      headers: {
        api_key: key,
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
      const key = fields.api_key?.trim() ?? "";
      if (!key) throw new ConnectError("Enter the Lusha API key");
      const res = await send(key, { method: "GET", path: "/v3/account/usage" });
      if (res.status === 401) throw new ConnectError("Lusha did not accept the API key");
      if (res.status === 403)
        throw new ConnectError(
          `Lusha refused the key. Ask Lusha to turn on API access for your plan. (${lushaError(res)})`,
        );
      if (res.status !== 200) throw new ConnectError(`Lusha sign-in failed. ${lushaError(res)}`);
      return {
        config: { host: LUSHA_HOST },
        secret: { api_key: key },
        accountLabel: `API key ending ${key.slice(-4)}`,
      };
    },

    call: (connection, request) => send(connection.secret.api_key ?? "", request),
  };
}

function parse(raw: string): unknown {
  try {
    return JSON.parse(raw);
  } catch {
    return raw;
  }
}
