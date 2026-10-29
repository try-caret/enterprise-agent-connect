import assert from "node:assert/strict";
import { generateKeyPairSync, verify } from "node:crypto";
import { test } from "node:test";
import { ConnectError, type VendorConnection } from "../../packages/connector-kit/src/types.ts";
import { nextAfter, normalizeOrg, oktaVendor, scopeFor } from "./vendor.ts";

// The Okta sign-in and calls against a fake Okta: the client assertion Okta would verify, a
// token per scope, the Link header's cursor, and the scope check.

const { privateKey, publicKey } = generateKeyPairSync("rsa", { modulusLength: 2048 });
const PEM = privateKey.export({ type: "pkcs8", format: "pem" }).toString();
const TOKEN_URL = "https://acme.okta.com/oauth2/v1/token";
const FIELDS = {
  org: "https://acme-admin.okta.com/admin",
  client_id: "0oa1234567890abcdef",
  key_id: "kid-1",
  private_key: PEM,
};

type Options = { granted?: string[]; tokenError?: { status: number; body: object } };

function fakeOkta(options: Options = {}) {
  const granted = options.granted ?? ["okta.users.read", "okta.groups.read", "okta.users.manage"];
  const tokens: { scope: string; assertion: string }[] = [];
  const calls: { url: string; init: RequestInit }[] = [];
  let reject = 0;
  const fetchImpl = (async (input: string | URL, init: RequestInit = {}) => {
    const url = String(input);
    if (url === TOKEN_URL) {
      const form = new URLSearchParams(String(init.body));
      tokens.push({ scope: form.get("scope")!, assertion: form.get("client_assertion")! });
      if (options.tokenError)
        return Response.json(options.tokenError.body, { status: options.tokenError.status });
      if (!granted.includes(form.get("scope")!))
        return Response.json(
          { error: "invalid_scope", error_description: "The following scopes are not allowed" },
          { status: 400 },
        );
      return Response.json({
        access_token: `at-${tokens.length}`,
        expires_in: 3600,
        scope: form.get("scope"),
      });
    }
    calls.push({ url, init });
    if (reject > 0) {
      reject--;
      return new Response("", { status: 401 });
    }
    if (url.includes("/api/v1/users?"))
      return Response.json([{ id: "00u1" }], {
        headers: {
          link: `<${url}>; rel="self", <https://acme.okta.com/api/v1/users?after=00u1&limit=1>; rel="next"`,
        },
      });
    return Response.json({ id: "00u1" });
  }) as typeof fetch;
  return { fetchImpl, tokens, calls, rejectNext: () => (reject = 1) };
}

function claims(jwt: string) {
  const [h, p, s] = jwt.split(".");
  const ok = verify("sha256", Buffer.from(`${h}.${p}`), publicKey, Buffer.from(s, "base64url"));
  return {
    ok,
    header: JSON.parse(Buffer.from(h, "base64url").toString()),
    payload: JSON.parse(Buffer.from(p, "base64url").toString()),
  };
}

test("an Okta address is the org's own host and nothing else", () => {
  assert.equal(normalizeOrg("acme"), "acme.okta.com");
  assert.equal(normalizeOrg("https://ACME-admin.okta.com/admin/dashboard"), "acme.okta.com");
  assert.equal(normalizeOrg("dev-1.oktapreview.com"), "dev-1.oktapreview.com");
  assert.equal(normalizeOrg("acme.okta-emea.com"), "acme.okta-emea.com");
  for (const bad of ["evil.com", "acme.okta.com.evil.com", "a.b.okta.com", "localhost:8080", ""])
    assert.throws(() => normalizeOrg(bad), ConnectError, bad);
});

test("each call's scope follows its resource and method", () => {
  assert.equal(scopeFor({ method: "GET", path: "/api/v1/users/00u1/groups" }), "okta.users.read");
  assert.equal(
    scopeFor({ method: "POST", path: "/api/v1/users/00u1/lifecycle/unlock" }),
    "okta.users.manage",
  );
  assert.equal(
    scopeFor({ method: "DELETE", path: "/api/v1/users/00u1/sessions?oauthTokens=true" }),
    "okta.users.manage",
  );
  assert.equal(
    scopeFor({ method: "PUT", path: "/api/v1/groups/00g1/users/00u1" }),
    "okta.groups.manage",
  );
  assert.equal(scopeFor({ method: "GET", path: "/api/v1/apps?q=x" }), "okta.apps.read");
  assert.equal(scopeFor({ method: "GET", path: "/api/v1/logs?since=x" }), "okta.logs.read");
  assert.throws(() => scopeFor({ method: "GET", path: "/api/v1/authorizationServers" }));
  assert.throws(() => scopeFor({ method: "GET", path: "/oauth2/v1/clients" }));
});

test("nextAfter reads the next page's cursor from a Link header", () => {
  assert.equal(
    nextAfter(
      '<https://a.okta.com/api/v1/users?limit=2>; rel="self", <https://a.okta.com/api/v1/users?after=abc%3D&limit=2>; rel="next"',
    ),
    "abc=",
  );
  assert.equal(nextAfter('<https://a.okta.com/api/v1/users?limit=2>; rel="self"'), null);
  assert.equal(nextAfter(null), null);
});

test("connect signs in with an RS256 client assertion Okta can verify", async () => {
  const okta = fakeOkta();
  const conn = await oktaVendor(okta.fetchImpl).connect(FIELDS);
  assert.deepEqual(conn.config, {
    org: "acme.okta.com",
    clientId: FIELDS.client_id,
    keyId: "kid-1",
  });
  assert.equal(conn.accountLabel, "acme.okta.com");
  assert.match(conn.secret.private_key, /^-----BEGIN PRIVATE KEY-----/);
  assert.equal(okta.tokens.length, 1);
  assert.equal(okta.tokens[0].scope, "okta.users.read");
  const jwt = claims(okta.tokens[0].assertion);
  assert.ok(jwt.ok);
  assert.deepEqual(jwt.header, { alg: "RS256", typ: "JWT", kid: "kid-1" });
  assert.equal(jwt.payload.iss, FIELDS.client_id);
  assert.equal(jwt.payload.sub, FIELDS.client_id);
  assert.equal(jwt.payload.aud, TOKEN_URL);
  assert.equal(jwt.payload.exp - jwt.payload.iat, 300);
  assert.ok(jwt.payload.jti);
});

test("connect takes the key as a JWK too, and refuses what is not an RSA private key", async () => {
  const okta = fakeOkta();
  const vendor = oktaVendor(okta.fetchImpl);
  const jwk = JSON.stringify({ ...privateKey.export({ format: "jwk" }), kid: "kid-1" });
  assert.ok((await vendor.connect({ ...FIELDS, private_key: jwk })).secret.private_key);
  const ec = generateKeyPairSync("ec", { namedCurve: "P-256" })
    .privateKey.export({ type: "pkcs8", format: "pem" })
    .toString();
  await assert.rejects(vendor.connect({ ...FIELDS, private_key: ec }), /RSA key/);
  await assert.rejects(vendor.connect({ ...FIELDS, private_key: "nope" }), /could not be read/);
  await assert.rejects(
    vendor.connect({
      ...FIELDS,
      private_key: publicKey.export({ type: "spki", format: "pem" }).toString(),
    }),
    /could not be read/,
  );
  await assert.rejects(vendor.connect({ ...FIELDS, client_id: "" }), /Client ID/);
  await assert.rejects(vendor.connect({ ...FIELDS, key_id: "a b" }), /KID/);
  await assert.rejects(vendor.connect({ ...FIELDS, org: "evil.com" }), /Okta address/);
});

test("connect explains Okta's refusals, DPoP included", async () => {
  const dpop = fakeOkta({
    tokenError: {
      status: 400,
      body: {
        error: "invalid_dpop_proof",
        error_description: "The DPoP proof JWT header is missing.",
      },
    },
  });
  await assert.rejects(
    oktaVendor(dpop.fetchImpl).connect(FIELDS),
    /clear "Require Demonstrating Proof of Possession/,
  );
  const bad = fakeOkta({
    tokenError: {
      status: 401,
      body: {
        error: "invalid_client",
        error_description: "The client_assertion signature is invalid.",
      },
    },
  });
  await assert.rejects(
    oktaVendor(bad.fetchImpl).connect(FIELDS),
    (e: Error) =>
      e instanceof ConnectError &&
      e.message === "Okta refused the app's sign-in: The client_assertion signature is invalid.",
  );
});

test("calls get a token for their own scope, kept, and minted again once after a 401", async () => {
  const okta = fakeOkta();
  const vendor = oktaVendor(okta.fetchImpl);
  const conn: VendorConnection = await vendor.connect(FIELDS);
  okta.tokens.length = 0;
  await vendor.call(conn, { method: "GET", path: "/api/v1/users/00u1" });
  await vendor.call(conn, { method: "POST", path: "/api/v1/users/00u1/lifecycle/unlock" });
  await vendor.call(conn, { method: "GET", path: "/api/v1/users/00u2" });
  // The read token from connect is reused; the manage one is minted once.
  assert.deepEqual(
    okta.tokens.map((t) => t.scope),
    ["okta.users.manage"],
  );
  assert.equal(okta.calls[0].url, "https://acme.okta.com/api/v1/users/00u1");
  assert.equal((okta.calls[0].init.headers as Record<string, string>).authorization, "Bearer at-1");
  assert.equal(okta.calls[0].init.redirect, "manual");

  okta.rejectNext();
  const res = await vendor.call(conn, { method: "GET", path: "/api/v1/users/00u1" });
  assert.equal(res.status, 200);
  assert.deepEqual(
    okta.tokens.map((t) => t.scope),
    ["okta.users.manage", "okta.users.read"],
  );
});

test("a list answer comes back with its next page's cursor", async () => {
  const okta = fakeOkta();
  const vendor = oktaVendor(okta.fetchImpl);
  const conn = await vendor.connect(FIELDS);
  const res = await vendor.call(conn, { method: "GET", path: "/api/v1/users?limit=1" });
  assert.deepEqual(res.data, { items: [{ id: "00u1" }], after: "00u1" });
});

test("a call for a scope Okta has not granted fails as a sign-in refusal", async () => {
  const okta = fakeOkta({ granted: ["okta.users.read"] });
  const vendor = oktaVendor(okta.fetchImpl);
  const conn = await vendor.connect(FIELDS);
  await assert.rejects(
    vendor.call(conn, { method: "GET", path: "/api/v1/logs?since=x" }),
    /Okta has not granted the app okta\.logs\.read/,
  );
});

test("the scope check names only the missing scopes the ceiling needs", async () => {
  const okta = fakeOkta({ granted: ["okta.users.read", "okta.groups.read", "okta.apps.read"] });
  const vendor = oktaVendor(okta.fetchImpl);
  const conn = await vendor.connect(FIELDS);
  const expected = ["okta.users.read", "okta.groups.read", "okta.logs.read", "okta.users.manage"];
  assert.deepEqual(await vendor.checks!.scopes(conn, { expected }), {
    values: ["okta.users.read", "okta.groups.read"],
    warning:
      "Missing: okta.logs.read, okta.users.manage. Grant them on the app's Okta API Scopes tab, or agents cannot do what you ticked.",
  });
  // okta.apps.read is granted but not needed: not mentioned.
  assert.deepEqual(await vendor.checks!.scopes(conn, { expected: ["okta.users.read"] }), {
    values: ["okta.users.read"],
  });
  const broken = fakeOkta({ tokenError: { status: 401, body: { error: "invalid_client" } } });
  await assert.rejects(
    oktaVendor(broken.fetchImpl).checks!.scopes(conn, { expected }),
    /refused the app's sign-in/,
  );
});
