import { createHash, createHmac, randomBytes } from "node:crypto";
import { tokensEqual } from "./secrets.js";
import { ConnectError } from "../../connector-kit/src/types.js";

// The `state` an admin carries to the app's approval page and back. It is signed, names the
// organization, app, capability ceiling and return address, expires, and carries a nonce the admin
// app also keeps in the admin's browser, so an approval only completes in the session that
// started it. After the approval, the admin signs in to the organization that approved; that
// second leg's state also names the organization (tenant) the sign-in must prove.

export type ConsentState = {
  org: string;
  app: string;
  ceiling: string[];
  redirectUri: string;
  tenant?: string;
  nonce: string;
  exp: number;
};

// The OpenID Connect nonce for the sign-in leg. Derived, so the browser-held nonce itself
// never appears in a URL.
export const signInNonce = (nonce: string) =>
  createHash("sha256").update(`sign-in:${nonce}`).digest("base64url");

const TTL_MS = 15 * 60 * 1000;

const mac = (key: Buffer, body: string) =>
  createHmac("sha256", key).update(body).digest("base64url");

export function signState(
  key: Buffer,
  input: { org: string; app: string; ceiling: string[]; redirectUri: string; tenant?: string },
  now = Date.now(),
  nonce = randomBytes(16).toString("base64url"),
): { state: string; nonce: string } {
  const body = Buffer.from(JSON.stringify({ ...input, nonce, exp: now + TTL_MS })).toString(
    "base64url",
  );
  return { state: `${body}.${mac(key, body)}`, nonce };
}

export function verifyState(
  key: Buffer,
  state: string,
  expect: { org: string; app: string; nonce: string },
  now = Date.now(),
): ConsentState {
  const [body, sig, extra] = state.split(".");
  if (!body || !sig || extra !== undefined || !tokensEqual(sig, mac(key, body)))
    throw new ConnectError("This approval did not start here. Start again from Integrations.");
  const parsed = JSON.parse(Buffer.from(body, "base64url").toString("utf8")) as ConsentState;
  if (
    parsed.org !== expect.org ||
    parsed.app !== expect.app ||
    !tokensEqual(parsed.nonce, expect.nonce)
  )
    throw new ConnectError("This approval did not start here. Start again from Integrations.");
  if (!(parsed.exp > now)) throw new ConnectError("This approval took too long. Start again.");
  return parsed;
}

/** The app a state names, unverified: only to choose what verifyState then checks it against. */
export function stateApp(state: string): string | null {
  try {
    const app = JSON.parse(Buffer.from(state.split(".")[0], "base64url").toString("utf8")).app;
    return typeof app === "string" ? app : null;
  } catch {
    return null;
  }
}
