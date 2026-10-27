import { createPublicKey, verify, type KeyObject } from "node:crypto";
import type { Fetch } from "../../../connector-kit/src/types.js";

// Checks the token Teams sends with every request to a bot's endpoint, as Microsoft documents
// for the Bot Connector (docs/teams.md 2.2). Every check is required; none can be turned off.

const METADATA = "https://login.botframework.com/v1/.well-known/openidconfiguration";
export const BOT_ISSUER = "https://api.botframework.com";
const SKEW_S = 300;
const REFRESH_MS = 24 * 3600_000;
const REFETCH_MIN_MS = 60_000;

// A request that fails a check. Answered 403 with nothing recorded but a log line.
export class Unauthorized extends Error {}

type Jwk = { kid?: string; kty?: string; n?: string; e?: string; endorsements?: unknown };
type SigningKey = { key: KeyObject; endorsements: string[] };

const isObject = (v: unknown): v is Record<string, unknown> =>
  typeof v === "object" && v !== null && !Array.isArray(v);

const b64json = (part: string) => JSON.parse(Buffer.from(part, "base64url").toString("utf8"));

export function botFrameworkAuth(fetchImpl: Fetch = fetch, now = () => Date.now()) {
  let keys = new Map<string, SigningKey>();
  let algorithms: string[] = [];
  let loadedAt = 0;
  let lastFetch = -Infinity;
  let inflight: Promise<void> | null = null;

  // One fetch at a time, shared by every request waiting on it.
  function refresh(): Promise<void> {
    inflight ??= load().finally(() => (inflight = null));
    return inflight;
  }

  async function load() {
    lastFetch = now();
    const meta = (await (await fetchImpl(METADATA)).json()) as {
      jwks_uri?: string;
      id_token_signing_alg_values_supported?: string[];
    };
    if (!meta.jwks_uri?.startsWith("https://login.botframework.com/"))
      throw new Error("Bot Framework metadata has an unexpected jwks_uri");
    const set = (await (await fetchImpl(meta.jwks_uri)).json()) as { keys?: Jwk[] };
    const next = new Map<string, SigningKey>();
    for (const jwk of set.keys ?? []) {
      if (!jwk.kid || jwk.kty !== "RSA" || !jwk.n || !jwk.e) continue;
      next.set(jwk.kid, {
        key: createPublicKey({ key: { kty: "RSA", n: jwk.n, e: jwk.e }, format: "jwk" }),
        endorsements: Array.isArray(jwk.endorsements)
          ? jwk.endorsements.filter((e): e is string => typeof e === "string")
          : [],
      });
    }
    keys = next;
    algorithms = meta.id_token_signing_alg_values_supported ?? [];
    loadedAt = now();
  }

  // Refreshed daily; an unknown key id refetches at most once a minute (key rollover). A
  // failed daily refresh keeps the keys already held (Microsoft rolls them over weeks apart)
  // and is retried at most once a minute.
  async function signingKey(kid: string): Promise<SigningKey> {
    if (now() - loadedAt > REFRESH_MS && now() - lastFetch >= REFETCH_MIN_MS) {
      try {
        await refresh();
      } catch (error) {
        if (!keys.size) throw error;
        console.error(
          "[agent-connect] Bot Framework key refresh failed; using the keys held",
          error,
        );
      }
    }
    let found = keys.get(kid);
    if (!found && now() - lastFetch >= REFETCH_MIN_MS) {
      await refresh();
      found = keys.get(kid);
    }
    if (!found) throw new Unauthorized("unknown signing key");
    return found;
  }

  /** Throws Unauthorized unless the bearer token is Microsoft's and for this bot. Returns the
   *  check that the activity, read only after this, is the one the token was issued for. */
  async function token(
    authorization: string | null,
    appId: string,
  ): Promise<(activity: { serviceUrl: string; channelId: string }) => void> {
    if (!authorization?.startsWith("Bearer ")) throw new Unauthorized("no bearer token");
    const parts = authorization.slice(7).split(".");
    if (parts.length !== 3) throw new Unauthorized("malformed token");
    let header: { alg?: string; kid?: string };
    let claims: Record<string, unknown>;
    try {
      header = b64json(parts[0]);
      claims = b64json(parts[1]);
      if (!isObject(header) || !isObject(claims)) throw new Error("not an object");
    } catch {
      throw new Unauthorized("malformed token");
    }
    if (header.alg !== "RS256" || !header.kid) throw new Unauthorized("unexpected algorithm");
    const { key, endorsements } = await signingKey(header.kid);
    if (!algorithms.includes(header.alg)) throw new Unauthorized("unexpected algorithm");
    const signed = verify(
      "sha256",
      Buffer.from(`${parts[0]}.${parts[1]}`),
      key,
      Buffer.from(parts[2], "base64url"),
    );
    if (!signed) throw new Unauthorized("bad signature");
    if (claims.iss !== BOT_ISSUER) throw new Unauthorized("wrong issuer");
    if (claims.aud !== appId) throw new Unauthorized("wrong audience");
    const t = Math.floor(now() / 1000);
    if (typeof claims.exp !== "number" || claims.exp + SKEW_S < t)
      throw new Unauthorized("expired");
    if (typeof claims.nbf === "number" && claims.nbf - SKEW_S > t)
      throw new Unauthorized("not yet valid");
    return (activity) => {
      // Microsoft names this claim in lowercase.
      if (claims.serviceurl !== activity.serviceUrl)
        throw new Unauthorized("serviceUrl does not match the token");
      if (activity.channelId !== "msteams" || !endorsements.includes(activity.channelId))
        throw new Unauthorized("key not endorsed for this channel");
    };
  }

  return {
    token,
    /** Both checks at once: the token, then the activity it came with. */
    async verify(
      authorization: string | null,
      expect: { appId: string; serviceUrl: string; channelId: string },
    ): Promise<void> {
      (await token(authorization, expect.appId))(expect);
    },
  };
}

export type BotAuth = ReturnType<typeof botFrameworkAuth>;
