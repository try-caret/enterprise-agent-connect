import assert from "node:assert/strict";
import { randomBytes } from "node:crypto";
import { test } from "node:test";
import {
  credentialAad,
  CredentialKey,
  hashAgentKey,
  looksLikeAgentKey,
  newAgentKey,
  tokensEqual,
} from "../src/secrets.ts";

const key = new CredentialKey(randomBytes(32).toString("base64"));
const aad = credentialAad("11111111-1111-1111-1111-111111111111", "jira");
const secret = { client_id: "abc", client_secret: "s3cr3t" };

test("sealed credentials open with the same key and row", () => {
  const sealed = key.seal(secret, aad);
  assert.equal(sealed.keyId, key.id);
  assert.ok(!sealed.ciphertext.toString("utf8").includes("s3cr3t"));
  assert.deepEqual(key.open(sealed, aad), secret);
});

test("every seal uses a fresh IV", () => {
  assert.notDeepEqual(key.seal(secret, aad).ciphertext, key.seal(secret, aad).ciphertext);
});

test("tampering with any part of the ciphertext is detected", () => {
  const sealed = key.seal(secret, aad);
  for (const at of [0, 12, 28, sealed.ciphertext.length - 1]) {
    const bad = Buffer.from(sealed.ciphertext);
    bad[at] ^= 1;
    assert.throws(() => key.open({ ...sealed, ciphertext: bad }, aad), `byte ${at}`);
  }
});

test("a credential moved to another organization or app does not open", () => {
  const sealed = key.seal(secret, aad);
  assert.throws(() =>
    key.open(sealed, credentialAad("22222222-2222-2222-2222-222222222222", "jira")),
  );
  assert.throws(() =>
    key.open(sealed, credentialAad("11111111-1111-1111-1111-111111111111", "salesforce")),
  );
});

test("a different key is refused by id before decrypting", () => {
  const other = new CredentialKey(randomBytes(32).toString("base64"));
  assert.throws(() => other.open(key.seal(secret, aad), aad), /sealed with key/);
  assert.throws(() => other.open({ ...key.seal(secret, aad), keyId: other.id }, aad));
});

test("the key must be exactly 32 bytes", () => {
  assert.throws(() => new CredentialKey(randomBytes(16).toString("base64")), /32 bytes/);
  assert.throws(() => new CredentialKey(""), /32 bytes/);
});

test("agent keys are random, recognizable and stored only as SHA-256", () => {
  const a = newAgentKey();
  assert.ok(looksLikeAgentKey(a));
  assert.notEqual(a, newAgentKey());
  assert.match(hashAgentKey(a), /^[0-9a-f]{64}$/);
  assert.equal(hashAgentKey(a), hashAgentKey(a));
  assert.notEqual(hashAgentKey(a), hashAgentKey(newAgentKey()));
  // The previous release's shape, so a rollback still takes it; the new one is accepted too.
  assert.match(a, /^arb_agent_[A-Za-z0-9_-]{43}$/);
  assert.ok(looksLikeAgentKey(`eac_agent_${"A".repeat(43)}`));
  assert.ok(!looksLikeAgentKey("arb_agent_short"));
  assert.ok(!looksLikeAgentKey(`x${a}`));
});

test("token comparison", () => {
  assert.ok(tokensEqual("same-token", "same-token"));
  assert.ok(!tokensEqual("same-token", "same-tokeN"));
  assert.ok(!tokensEqual("", "x"));
});

test("after a rotation, rows sealed with the previous key still open; new seals use the new key", () => {
  const oldB64 = randomBytes(32).toString("base64");
  const old = new CredentialKey(oldB64);
  const sealedOld = old.seal(secret, aad);
  const rotated = new CredentialKey(randomBytes(32).toString("base64"), oldB64);
  assert.deepEqual(rotated.open(sealedOld, aad), secret);
  assert.equal(rotated.seal(secret, aad).keyId, rotated.id);
  assert.notEqual(rotated.id, old.id);
  assert.throws(
    () => new CredentialKey(randomBytes(32).toString("base64"), "c2hvcnQ="),
    /PREVIOUS/,
  );
});

// Sealed by this code with a fixed key and committed: whatever the code becomes, what an
// existing install stored must keep opening with its key, and derived keys must not move.
test("credentials sealed before keep opening, and derived keys stay the same", () => {
  const fixed = new CredentialKey(Buffer.alloc(32, 7).toString("base64"));
  const sealed = {
    ciphertext: Buffer.from(
      "d6+dzw/2adxP3yQxx0JmWYO5S9jWtucqkvmkfpLc5Mij1O4/CQHQaZ8HeV/ZTQmB8JYyU2Ed",
      "base64",
    ),
    keyId: "4bb06f8e4e3a7715",
  };
  assert.equal(fixed.id, sealed.keyId);
  assert.deepEqual(
    fixed.open(sealed, credentialAad("11111111-1111-1111-1111-111111111111", "provider:microsoft")),
    { client_secret: "golden" },
  );
  assert.equal(
    fixed.derive("consent").toString("base64"),
    "u506qhO4NMtzJlOwk3SKhMVHH4wGY5SSmkeoGFdOcX8=",
  );
});
