import assert from "node:assert/strict";
import { randomBytes } from "node:crypto";
import { test } from "node:test";
import { signState, verifyState } from "../src/consent.ts";
import { ConnectError } from "../../connector-kit/src/types.ts";
import { ADMIN_TOKEN, harness, microsoft, ORG, TENANT, withBy } from "./helpers.ts";

const KEY = randomBytes(32);
const OTHER_ORG = "44444444-4444-4444-4444-444444444444";
const start = (now = 1_000_000) =>
  signState(
    KEY,
    { org: ORG, app: "outlook-mail", ceiling: ["read"], redirectUri: "https://a.example/cb" },
    now,
  );
const expect = (nonce: string) => ({ org: ORG, app: "outlook-mail", nonce });

test("a signed state verifies and carries the ceiling", () => {
  const { state, nonce } = start();
  assert.deepEqual(verifyState(KEY, state, expect(nonce), 1_000_000 + 60_000).ceiling, ["read"]);
});

test("a state expires after 15 minutes", () => {
  const { state, nonce } = start();
  assert.throws(() => verifyState(KEY, state, expect(nonce), 1_000_000 + 15 * 60_000), /too long/);
});

test("a tampered, re-keyed or misdirected state is refused", () => {
  const { state, nonce } = start();
  const [body, sig] = state.split(".");
  const forged = Buffer.from(
    JSON.stringify({
      ...JSON.parse(Buffer.from(body, "base64url").toString()),
      ceiling: ["read", "send"],
    }),
  ).toString("base64url");
  for (const bad of [`${forged}.${sig}`, `${body}.${sig.slice(1)}x`, `${body}.${sig}.x`, body, ""])
    assert.throws(() => verifyState(KEY, bad, expect(nonce), 1_000_000), ConnectError);
  assert.throws(() => verifyState(randomBytes(32), state, expect(nonce), 1_000_000), ConnectError);
  assert.throws(
    () => verifyState(KEY, state, { ...expect(nonce), org: OTHER_ORG }, 1_000_000),
    ConnectError,
  );
  assert.throws(
    () => verifyState(KEY, state, { ...expect(nonce), app: "jira" }, 1_000_000),
    ConnectError,
  );
  assert.throws(() => verifyState(KEY, state, expect(start().nonce), 1_000_000), ConnectError);
});

function admin(h: ReturnType<typeof harness>, method: string, path: string, body?: unknown) {
  return h.app.fetch(
    new Request(`http://svc/agent-connect/admin/v1${path}`, {
      method,
      headers: { authorization: `Bearer ${ADMIN_TOKEN}`, "content-type": "application/json" },
      body: body === undefined ? undefined : JSON.stringify(withBy(body)),
    }),
  );
}

const REDIRECT = "https://admin.example/admin/integrations/outlook-mail/callback";

async function begin(h: ReturnType<typeof harness>, ceiling = ["send"]) {
  const res = await admin(h, "POST", `/orgs/${ORG}/connections/outlook-mail/consent`, {
    capabilities: ceiling,
    redirectUri: REDIRECT,
  });
  assert.equal(res.status, 200);
  const { url, nonce } = (await res.json()) as any;
  return { state: new URL(url).searchParams.get("state")!, nonce, url: url as string };
}

const finish = (h: ReturnType<typeof harness>, body: Record<string, unknown>, org = ORG) =>
  admin(h, "POST", `/orgs/${org}/connections/outlook-mail/consent/callback`, {
    connectedBy: "admin@x.com",
    ...body,
  });

const VICTIM = "55555555-5555-5555-5555-555555555555";

// Microsoft's approval came back naming `tenant`; the service sends the admin on to sign in.
async function approved(h: ReturnType<typeof harness>, tenant: string, ceiling = ["send"]) {
  const { state, nonce } = await begin(h, ceiling);
  const res = await finish(h, { state, nonce, params: { tenant, admin_consent: "True", state } });
  assert.equal(res.status, 200);
  const { redirect } = (await res.json()) as any;
  const signIn = new URL(redirect);
  assert.equal(signIn.pathname, `/${tenant}/authorize`);
  assert.notEqual(signIn.searchParams.get("nonce"), nonce);
  return { nonce, signInState: signIn.searchParams.get("state")! };
}

test("approval, then sign-in, connects the proven tenant at the ceiling chosen before leaving", async () => {
  const h = harness({ shared: true });
  h.msp.signIns["code-1"] = TENANT;
  const { nonce, signInState } = await approved(h, TENANT, ["organize"]);
  assert.deepEqual(h.msp.connected, []);
  const res = await finish(h, {
    state: signInState,
    nonce,
    // A tenant in the sign-in's answer is ignored: the signed state names it.
    params: { code: "code-1", state: signInState, tenant: VICTIM },
  });
  assert.equal(res.status, 200);
  // Outlook mail needs no settings of its own, so it is connected with the approval.
  assert.deepEqual(((await res.json()) as any).checks, [
    { label: microsoft.verify.label, values: [] },
  ]);
  assert.deepEqual(h.msp.connected, [TENANT]);
  // Read is required, so it is always part of the ceiling.
  assert.deepEqual(h.state.saved[0].capabilities, ["read", "organize"]);
  assert.equal(h.state.saved[0].providerId, "p1");
  assert.equal(h.state.providers[0].config.auth, "arbium_app");
  assert.equal(h.state.providers[0].config.tenantId, TENANT);
});

test("approving for an app with settings of its own connects only the Microsoft connection", async () => {
  const h = harness({ shared: true });
  h.state.providers.length = 0;
  h.state.connections = h.state.connections.filter((c) => !c.providerId);
  h.msp.signIns["code-1"] = TENANT;
  const begun = await admin(h, "POST", `/orgs/${ORG}/connections/sharepoint-files/consent`, {
    capabilities: ["read_files"],
    redirectUri: REDIRECT,
  });
  const { url, nonce } = (await begun.json()) as any;
  const state = new URL(url).searchParams.get("state")!;
  const call = (body: Record<string, unknown>) =>
    admin(h, "POST", `/orgs/${ORG}/connections/sharepoint-files/consent/callback`, {
      connectedBy: "a",
      ...body,
    });
  const first = (await (
    await call({ state, nonce, params: { tenant: TENANT, admin_consent: "True" } })
  ).json()) as any;
  const signInState = new URL(first.redirect).searchParams.get("state")!;
  const done = (await (
    await call({ state: signInState, nonce, params: { code: "code-1" } })
  ).json()) as any;
  assert.equal(done.provider.config.auth, "arbium_app");
  assert.deepEqual(h.state.saved, []);
});

test("an approval answer edited to name another organization is refused at sign-in", async () => {
  const h = harness({ shared: true });
  // The admin can only sign in to their own organization.
  h.msp.signIns["code-1"] = TENANT;
  const { nonce, signInState } = await approved(h, VICTIM);
  const res = await finish(h, {
    state: signInState,
    nonce,
    params: { code: "code-1", state: signInState },
  });
  assert.equal(res.status, 422);
  const noCode = await finish(h, { state: signInState, nonce, params: { state: signInState } });
  assert.equal(noCode.status, 422);
  assert.deepEqual(h.msp.connected, []);
  assert.deepEqual(h.state.saved, []);
});

test("a declined, forged, or other-session approval connects nothing", async () => {
  const h = harness({ shared: true });
  const { state, nonce } = await begin(h);
  const declined = await finish(h, {
    state,
    nonce,
    params: {
      error: "access_denied",
      error_description: "AADSTS65004: User declined.\r\nTrace ID: x",
    },
  });
  assert.equal(declined.status, 422);
  assert.match(((await declined.json()) as any).error, /User declined\.$/);
  const notApproved = await finish(h, { state, nonce, params: { tenant: TENANT } });
  assert.equal(notApproved.status, 422);
  const otherSession = await finish(h, {
    state,
    nonce: (await begin(h)).nonce,
    params: { tenant: TENANT, admin_consent: "True" },
  });
  assert.equal(otherSession.status, 422);
  await admin(h, "PUT", "/orgs/44444444-4444-4444-4444-444444444444", {});
  const otherOrg = await finish(
    h,
    { state, nonce, params: { tenant: TENANT, admin_consent: "True" } },
    "44444444-4444-4444-4444-444444444444",
  );
  assert.equal(otherOrg.status, 422);
  assert.deepEqual(h.msp.connected, []);
});

test("approving the deployment's shared app is refused unless the deployment offers it", async () => {
  const h = harness();
  const off = await admin(h, "POST", `/orgs/${ORG}/connections/outlook-mail/consent`, {
    capabilities: ["read"],
    redirectUri: REDIRECT,
  });
  assert.equal(off.status, 400);
  const on = harness({ shared: true });
  const jira = await admin(on, "POST", `/orgs/${ORG}/connections/jira/consent`, {
    capabilities: ["read"],
    redirectUri: REDIRECT,
  });
  assert.equal(jira.status, 400);
  const plainHttp = await admin(on, "POST", `/orgs/${ORG}/connections/outlook-mail/consent`, {
    capabilities: ["read"],
    redirectUri: "http://evil.example/cb",
  });
  assert.equal(plainHttp.status, 400);
  assert.deepEqual(on.msp.connected, []);
});
