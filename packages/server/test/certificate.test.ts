import assert from "node:assert/strict";
import { constants, createHash, createPrivateKey, verify, X509Certificate } from "node:crypto";
import { test } from "node:test";
import {
  clientAssertion,
  describeCertificate,
  newCertificate,
} from "../../connector-kit/src/certificate.ts";

test("a new certificate is a valid self-signed X.509 for its own key, two years long", () => {
  const now = new Date("2026-10-21T10:00:00Z");
  const made = newCertificate("Example Agents", now);
  const cert = new X509Certificate(made.certificate);
  assert.equal(cert.subject, "CN=Example Agents");
  assert.equal(cert.issuer, "CN=Example Agents");
  assert.ok(cert.verify(cert.publicKey));
  assert.ok(cert.checkPrivateKey(createPrivateKey(made.privateKey)));
  assert.equal(new Date(cert.validFrom).toISOString(), "2026-10-21T10:00:00.000Z");
  assert.equal(describeCertificate(made.certificate).expiresAt, "2028-10-20T10:00:00.000Z");
  assert.match(describeCertificate(made.certificate).thumbprint, /^[0-9A-F]{40}$/);
  assert.match(made.certificate, /^-----BEGIN CERTIFICATE-----\n/);
  assert.ok(!made.certificate.includes("PRIVATE"));
});

test("each certificate has its own key and serial", () => {
  const a = new X509Certificate(newCertificate("A").certificate);
  const b = new X509Certificate(newCertificate("A").certificate);
  assert.notEqual(a.serialNumber, b.serialNumber);
  assert.notEqual(a.fingerprint256, b.fingerprint256);
  assert.ok(!a.verify(b.publicKey));
});

test("the client assertion is PS256 (RSA-PSS) with the certificate's x5t#S256", () => {
  const made = newCertificate("A");
  const now = Date.parse("2026-10-21T10:00:00Z");
  const tokenUrl = "https://login.microsoftonline.com/t/oauth2/v2.0/token";
  const jwt = clientAssertion({ clientId: "c", tokenUrl, ...made, now });
  const [h, p, s] = jwt.split(".");
  const header = JSON.parse(Buffer.from(h, "base64url").toString());
  const claims = JSON.parse(Buffer.from(p, "base64url").toString());
  const raw = new X509Certificate(made.certificate).raw;
  assert.equal(header.alg, "PS256");
  assert.equal(header["x5t#S256"], createHash("sha256").update(raw).digest("base64url"));
  assert.deepEqual(
    { aud: claims.aud, iss: claims.iss, sub: claims.sub, exp: claims.exp - claims.iat },
    { aud: tokenUrl, iss: "c", sub: "c", exp: 300 },
  );
  const pss = {
    key: new X509Certificate(made.certificate).publicKey,
    padding: constants.RSA_PKCS1_PSS_PADDING,
    saltLength: constants.RSA_PSS_SALTLEN_DIGEST,
  };
  assert.ok(verify("sha256", Buffer.from(`${h}.${p}`), pss, Buffer.from(s, "base64url")));
});
