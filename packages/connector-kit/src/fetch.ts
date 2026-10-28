import { ConnectError, type Fetch } from "./types.js";

// How every coded connector reaches its app. A redirect is refused, never followed, so a token
// or body cannot be carried to another host. A 429 means the app did not run the call, so when
// its Retry-After is short the call is sent again (at most RETRIES times); a longer wait is
// passed back as the 429.

const RETRIES = 2;
const MAX_WAIT_SECONDS = 10;
const guarded = new WeakSet<Fetch>();
const sleep = (seconds: number) => new Promise((r) => setTimeout(r, seconds * 1000));

/** Seconds to wait from a Retry-After header (seconds or an HTTP date), or null. */
export function retryAfter(header: string | null, now = Date.now()): number | null {
  if (!header) return null;
  const seconds = /^\d+$/.test(header.trim()) ? Number(header) : (Date.parse(header) - now) / 1000;
  return Number.isFinite(seconds) ? Math.max(0, Math.ceil(seconds)) : null;
}

export function vendorFetch(name: string, fetchImpl: Fetch = fetch, wait = sleep): Fetch {
  if (guarded.has(fetchImpl)) return fetchImpl;
  const safe = (async (input: string | URL | Request, init?: RequestInit) => {
    for (let attempt = 0; ; attempt++) {
      const res = await fetchImpl(input, { ...init, redirect: "manual" });
      if (res.status >= 300 && res.status < 400) {
        await res.body?.cancel();
        throw new ConnectError(
          `${name} answered with a redirect (${res.status}), which is not followed. Check the address it was connected with.`,
        );
      }
      const after = res.status === 429 ? retryAfter(res.headers.get("retry-after")) : null;
      if (after === null || attempt >= RETRIES || after > MAX_WAIT_SECONDS) return res;
      await res.body?.cancel();
      await wait(Math.max(after, 1));
    }
  }) as Fetch;
  guarded.add(safe);
  return safe;
}

/** `${base}${path}`, refused when the path would leave base's origin (e.g. "@other.host"). */
export function vendorUrl(base: string, path: string): string {
  const url = `${base}${path}`;
  if (URL.parse(url)?.origin !== new URL(base).origin)
    throw new Error(`Refused a call leaving ${new URL(base).host}`);
  return url;
}
