import {
  constants,
  createHash,
  generateKeyPairSync,
  randomBytes,
  randomUUID,
  sign,
  X509Certificate,
} from "node:crypto";

// A key pair the service makes for one connection: the private key is sealed with the connection's
// credential and never leaves the service; the self-signed certificate is what the admin
// uploads to their own app registration. Node can parse X.509 but not create it, so the
// certificate's DER is written here (RFC 5280, no extensions).

const VALID_DAYS = 730;

function der(tag: number, ...parts: Buffer[]): Buffer {
  const body = Buffer.concat(parts);
  let length: Buffer;
  if (body.length < 0x80) length = Buffer.from([body.length]);
  else {
    const bytes: number[] = [];
    for (let n = body.length; n > 0; n = Math.floor(n / 256)) bytes.unshift(n % 256);
    length = Buffer.from([0x80 | bytes.length, ...bytes]);
  }
  return Buffer.concat([Buffer.from([tag]), length, body]);
}

const seq = (...parts: Buffer[]) => der(0x30, ...parts);

function oid(dotted: string): Buffer {
  const [a, b, ...rest] = dotted.split(".").map(Number);
  const bytes = [40 * a + b];
  for (const n of rest) {
    const chunk = [n & 0x7f];
    for (let v = n >> 7; v > 0; v >>= 7) chunk.unshift((v & 0x7f) | 0x80);
    bytes.push(...chunk);
  }
  return der(0x06, Buffer.from(bytes));
}

// UTCTime, YYMMDDHHMMSSZ: valid for dates before 2050.
const utcTime = (d: Date) =>
  der(0x17, Buffer.from(d.toISOString().replace(/[-:T]/g, "").slice(2, 14) + "Z"));

const name = (cn: string) =>
  seq(der(0x31, seq(oid("2.5.4.3"), der(0x0c, Buffer.from(cn, "utf8")))));

export type NewCertificate = { certificate: string; privateKey: string };

export function newCertificate(commonName: string, now = new Date()): NewCertificate {
  const { privateKey, publicKey } = generateKeyPairSync("rsa", { modulusLength: 2048 });
  const serial = randomBytes(16);
  serial[0] = (serial[0] & 0x7f) | 0x01;
  const sha256WithRsa = seq(oid("1.2.840.113549.1.1.11"), der(0x05));
  const tbs = seq(
    der(0xa0, der(0x02, Buffer.from([2]))),
    der(0x02, serial),
    sha256WithRsa,
    name(commonName),
    seq(utcTime(now), utcTime(new Date(now.getTime() + VALID_DAYS * 86_400_000))),
    name(commonName),
    publicKey.export({ type: "spki", format: "der" }),
  );
  const signature = sign("sha256", tbs, privateKey);
  const cert = seq(tbs, sha256WithRsa, der(0x03, Buffer.from([0]), signature));
  const lines = cert.toString("base64").match(/.{1,64}/g)!;
  return {
    certificate: `-----BEGIN CERTIFICATE-----\n${lines.join("\n")}\n-----END CERTIFICATE-----\n`,
    privateKey: privateKey.export({ type: "pkcs8", format: "pem" }).toString(),
  };
}

/** What the admin sees of a certificate: Entra lists it by its SHA-1 thumbprint. */
export function describeCertificate(pem: string): { thumbprint: string; expiresAt: string } {
  const cert = new X509Certificate(pem);
  return {
    thumbprint: cert.fingerprint.replaceAll(":", ""),
    expiresAt: new Date(cert.validTo).toISOString(),
  };
}

const b64url = (v: Buffer | string) => Buffer.from(v).toString("base64url");

/** A signed client assertion (RFC 7523) for Microsoft's token endpoint, valid five minutes.
 *  Microsoft pairs the x5t#S256 header with PS256 (RSA-PSS). */
export function clientAssertion(input: {
  clientId: string;
  tokenUrl: string;
  certificate: string;
  privateKey: string;
  now?: number;
}): string {
  const raw = new X509Certificate(input.certificate).raw;
  const now = Math.floor((input.now ?? Date.now()) / 1000);
  const header = b64url(
    JSON.stringify({
      alg: "PS256",
      typ: "JWT",
      "x5t#S256": b64url(createHash("sha256").update(raw).digest()),
    }),
  );
  const payload = b64url(
    JSON.stringify({
      aud: input.tokenUrl,
      iss: input.clientId,
      sub: input.clientId,
      jti: randomUUID(),
      nbf: now,
      iat: now,
      exp: now + 300,
    }),
  );
  const signature = sign("sha256", Buffer.from(`${header}.${payload}`), {
    key: input.privateKey,
    padding: constants.RSA_PKCS1_PSS_PADDING,
    saltLength: constants.RSA_PSS_SALTLEN_DIGEST,
  });
  return `${header}.${payload}.${b64url(signature)}`;
}
