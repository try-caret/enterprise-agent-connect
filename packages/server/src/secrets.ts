import {
  createCipheriv,
  createDecipheriv,
  createHash,
  hkdfSync,
  randomBytes,
  timingSafeEqual,
} from "node:crypto";

// Vendor credentials at rest: AES-256-GCM, bound to their row by the AAD so a ciphertext
// copied onto another organization's or app's row fails to open. Each row records the
// id of the key that sealed it, so a rotation can tell old rows from new.

export type Sealed = { ciphertext: Buffer; keyId: string };

const keyId = (key: Buffer) => createHash("sha256").update(key).digest("hex").slice(0, 16);

function parseKey(base64: string, name: string): Buffer {
  const key = Buffer.from(base64, "base64");
  if (key.length !== 32) throw new Error(`${name} must be 32 bytes, base64`);
  return key;
}

// Rotation: set the new key as CONNECT_ENCRYPTION_KEY and the old one as `previous`. Rows
// sealed with the old key still open, and are re-sealed with the new one when next checked.
export class CredentialKey {
  readonly id: string;
  readonly #key: Buffer;
  readonly #previous: Map<string, Buffer>;

  constructor(base64: string, previousBase64?: string) {
    this.#key = parseKey(base64, "CONNECT_ENCRYPTION_KEY");
    this.id = keyId(this.#key);
    const previous = previousBase64
      ? parseKey(previousBase64, "CONNECT_ENCRYPTION_KEY_PREVIOUS")
      : null;
    this.#previous = new Map(previous ? [[keyId(previous), previous]] : []);
  }

  /** Ids of the previous key(s) still loaded while rotating. */
  get previousIds(): string[] {
    return [...this.#previous.keys()];
  }

  // A separate key for another purpose (e.g. signing consent state), never the sealing key.
  derive(purpose: string): Buffer {
    return Buffer.from(hkdfSync("sha256", this.#key, "", `agent-connect:${purpose}`, 32));
  }

  seal(secret: Record<string, string>, aad: string): Sealed {
    const iv = randomBytes(12);
    const cipher = createCipheriv("aes-256-gcm", this.#key, iv);
    cipher.setAAD(Buffer.from(aad));
    const body = Buffer.concat([cipher.update(JSON.stringify(secret), "utf8"), cipher.final()]);
    return { ciphertext: Buffer.concat([iv, cipher.getAuthTag(), body]), keyId: this.id };
  }

  open(sealed: Sealed, aad: string): Record<string, string> {
    const key = sealed.keyId === this.id ? this.#key : this.#previous.get(sealed.keyId);
    if (!key)
      throw new Error(`Credential was sealed with key ${sealed.keyId}, which is not loaded`);
    const iv = sealed.ciphertext.subarray(0, 12);
    const tag = sealed.ciphertext.subarray(12, 28);
    const decipher = createDecipheriv("aes-256-gcm", key, iv);
    decipher.setAAD(Buffer.from(aad));
    decipher.setAuthTag(tag);
    const plain = Buffer.concat([
      decipher.update(sealed.ciphertext.subarray(28)),
      decipher.final(),
    ]);
    return JSON.parse(plain.toString("utf8"));
  }
}

export const credentialAad = (organizationId: string, connectorId: string) =>
  `agent-connect:${organizationId}:${connectorId}`;

// Both shapes are accepted, forever (the shape check is not the authentication; the hash
// lookup is). New keys keep arb_agent_ for this release, because a rollback to the previous one,
// which accepts only that, must still take keys made meanwhile. Switch to eac_agent_ in the
// next release.
const AGENT_KEY_PREFIX = "arb_agent_";

export function newAgentKey(): string {
  return AGENT_KEY_PREFIX + randomBytes(32).toString("base64url");
}

// Agent keys carry 256 bits of randomness, so an unsalted SHA-256 is enough to make a
// database read useless for calling the service.
export function hashAgentKey(key: string): string {
  return createHash("sha256").update(key).digest("hex");
}

export function looksLikeAgentKey(key: string): boolean {
  return /^(eac|arb)_agent_[A-Za-z0-9_-]{43}$/.test(key);
}

export function tokensEqual(a: string, b: string): boolean {
  const ha = createHash("sha256").update(a).digest();
  const hb = createHash("sha256").update(b).digest();
  return timingSafeEqual(ha, hb);
}
