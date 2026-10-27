import assert from "node:assert/strict";
import { test } from "node:test";
import { compareRoles } from "../../../connectors/microsoft/client.ts";
import { credentialAad } from "../src/secrets.ts";
import { sameSignIn } from "../src/service.ts";
import {
  ADMIN_TOKEN,
  AGENT,
  FINANCE,
  harness,
  microsoftProvider,
  ORG,
  TENANT,
  withBy,
} from "./helpers.ts";

type H = ReturnType<typeof harness>;
const OTHER_TENANT = "66666666-6666-6666-6666-666666666666";
const CLIENT = "77777777-7777-7777-7777-777777777777";
const BY = { by: "admin@x" };

async function admin(h: H, method: string, path: string, body?: unknown) {
  const res = await h.app.fetch(
    new Request(`http://svc/agent-connect/admin/v1/orgs/${ORG}${path}`, {
      method,
      headers: { authorization: `Bearer ${ADMIN_TOKEN}`, "content-type": "application/json" },
      body: body === undefined ? undefined : JSON.stringify(withBy(body)),
    }),
  );
  const text = await res.text();
  return { status: res.status, body: text ? JSON.parse(text) : null, text };
}

// No Microsoft connection yet, nor any Microsoft app.
function fresh(): H {
  const h = harness();
  h.state.providers.length = 0;
  h.state.connections = h.state.connections.filter((c) => !c.providerId);
  return h;
}

const connectMicrosoft = (h: H, fields: Record<string, string>) =>
  admin(h, "PUT", "/providers/microsoft", {
    fields: { tenant_id: TENANT, client_id: CLIENT, ...fields },
    connectedBy: "admin@x",
  });

const connectApp = (h: H, app: string, capabilities: string[], fields = {}) =>
  admin(h, "PUT", `/connections/${app}`, { capabilities, fields, connectedBy: "admin@x" });

test("an app of a provider cannot connect before the provider", async () => {
  const h = fresh();
  const res = await connectApp(h, "outlook-mail", ["read"]);
  assert.equal(res.status, 422);
  assert.match(res.body.error, /Connect Microsoft 365 first/);
});

test("the second Microsoft app reuses the connection: credentials are entered once", async () => {
  const h = fresh();
  assert.equal((await connectMicrosoft(h, { client_secret: "s3cret" })).status, 200);
  assert.equal(
    (await connectApp(h, "sharepoint-files", ["read_files"], { sites: FINANCE })).status,
    200,
  );
  assert.equal((await connectApp(h, "outlook-mail", ["read", "send"])).status, 200);
  assert.deepEqual(h.msp.connected, [TENANT]);
  const provider = h.state.providers[0];
  for (const saved of h.state.saved) {
    assert.equal(saved.providerId, provider.id);
    assert.equal(saved.credential, null);
    assert.ok(!("tenantId" in saved.config));
  }
  const listed = await admin(h, "GET", "/providers");
  assert.deepEqual(listed.body[0].apps.sort(), ["outlook-mail", "sharepoint-files"]);
  assert.ok(!listed.text.includes("s3cret"));
  assert.ok(!("credential" in listed.body[0]));
});

test("a certificate is made before connecting, and the connection signs in with it", async () => {
  const h = fresh();
  const made = await admin(h, "POST", "/providers/microsoft/certificate", BY);
  assert.equal(made.status, 200);
  assert.equal(made.body.status, "pending");
  assert.match(made.body.certificate.thumbprint, /^[0-9A-F]{40}$/);
  assert.match(made.body.config.certificate, /BEGIN CERTIFICATE/);
  assert.ok(!made.text.includes("PRIVATE KEY"));
  // Not connected yet: no app can use it.
  assert.equal((await connectApp(h, "outlook-mail", ["read"])).status, 422);
  const connected = await connectMicrosoft(h, {});
  assert.equal(connected.status, 200);
  assert.equal(connected.body.status, "connected");
  assert.equal(connected.body.config.auth, "certificate");
  assert.equal(connected.body.certificate.thumbprint, made.body.certificate.thumbprint);
  const used = h.msp.verified.at(-1)!;
  assert.equal(used.config.certificate, made.body.config.certificate);
  assert.match(used.secret.private_key, /BEGIN PRIVATE KEY/);
});

test("without a certificate or a secret nothing connects", async () => {
  const h = fresh();
  const res = await connectMicrosoft(h, {});
  assert.equal(res.status, 422);
  assert.deepEqual(h.state.providers, []);
});

test("rotating: the new certificate waits until Microsoft accepts it, then replaces the old", async () => {
  const h = fresh();
  await admin(h, "POST", "/providers/microsoft/certificate", BY);
  const first = (await connectMicrosoft(h, {})).body;
  const next = (await admin(h, "POST", "/providers/microsoft/certificate", BY)).body;
  assert.equal(next.certificate.thumbprint, first.certificate.thumbprint);
  assert.notEqual(next.nextCertificate.thumbprint, first.certificate.thumbprint);
  // Calls keep using the old certificate meanwhile.
  assert.equal(h.state.providers[0].config.certificate, first.config.certificate);

  h.msp.refuse = (c) => c.config.certificate === next.config.nextCertificate;
  const refused = await admin(h, "POST", "/providers/microsoft/certificate/activate", BY);
  assert.equal(refused.status, 422);
  assert.equal(h.state.providers[0].config.certificate, first.config.certificate);

  h.msp.refuse = () => false;
  // The old certificate stopped working meanwhile; switching proves the new one at once.
  h.state.providers[0].status = "needs_reconnect";
  const switched = await admin(h, "POST", "/providers/microsoft/certificate/activate", BY);
  assert.equal(switched.status, 200);
  assert.equal(h.state.providers[0].status, "connected");
  assert.equal(switched.body.certificate.thumbprint, next.nextCertificate.thumbprint);
  assert.equal(switched.body.nextCertificate, null);
  const row = h.state.providers[0];
  const secret = h.key.open(
    { ciphertext: row.credential, keyId: row.credentialKeyId },
    credentialAad(ORG, row.credentialScope),
  );
  assert.deepEqual(Object.keys(secret), ["private_key"]);
  assert.equal(
    (await admin(h, "POST", "/providers/microsoft/certificate/activate", BY)).status,
    400,
  );
});

test("each change to the Microsoft sign-in is audited with who made it", async () => {
  const h = fresh();
  await admin(h, "POST", "/providers/microsoft/certificate", BY);
  await connectMicrosoft(h, {});
  await connectApp(h, "outlook-mail", ["read"]);
  await admin(h, "POST", "/providers/microsoft/certificate", BY);
  await admin(h, "POST", "/providers/microsoft/certificate/activate", BY);
  await admin(h, "DELETE", "/providers/microsoft", BY);
  const rows = h.state.audit.filter(
    (a) => a.tool.startsWith("admin:") && a.connectorId === "microsoft",
  );
  assert.deepEqual(
    rows.map((r) => [r.tool, r.agentId, r.agentName, r.connectorId]),
    [
      ["admin:create_certificate", null, "admin@x", "microsoft"],
      ["admin:connect", null, "admin@x", "microsoft"],
      ["admin:create_certificate", null, "admin@x", "microsoft"],
      ["admin:switch_certificate", null, "admin@x", "microsoft"],
      ["admin:disconnect", null, "admin@x", "microsoft"],
    ],
  );
  assert.match(rows[0].target!, /^first certificate [0-9A-F]{40}$/);
  assert.match(
    rows[3].target!,
    /certificate [0-9A-F]{40} -> certificate app .* certificate [0-9A-F]{40}$/,
  );
  assert.match(rows[4].target!, /disconnected outlook-mail$/);
  // The admin is named in each request.
  assert.equal((await admin(h, "POST", "/providers/microsoft/certificate")).status, 400);
});

test("a connection on a client secret switches to a certificate, dropping the secret", async () => {
  const h = harness();
  await admin(h, "POST", "/providers/microsoft/certificate", BY);
  const switched = await admin(h, "POST", "/providers/microsoft/certificate/activate", BY);
  assert.equal(switched.body.config.auth, "certificate");
  const row = h.state.providers[0];
  const secret = h.key.open(
    { ciphertext: row.credential, keyId: row.credentialKeyId },
    credentialAad(ORG, row.credentialScope),
  );
  assert.ok(!("client_secret" in secret));
});

test("disconnecting Microsoft disconnects every Microsoft app and its keys' access", async () => {
  const h = harness();
  h.addAgent(["read"], {}, "active", "jira");
  h.state.grants.push(
    {
      agentId: AGENT,
      connectorId: "outlook-mail",
      capabilities: ["read"],
      limits: {},
      updatedAt: "",
    },
    {
      agentId: AGENT,
      connectorId: "sharepoint-files",
      capabilities: ["read_files"],
      limits: {},
      updatedAt: "",
    },
  );
  assert.equal((await admin(h, "DELETE", "/providers/microsoft", BY)).status, 204);
  assert.deepEqual(
    h.state.connections.map((c) => c.connectorId),
    ["jira"],
  );
  assert.deepEqual(
    h.state.grants.map((g) => g.connectorId),
    ["jira"],
  );
  assert.equal((await admin(h, "DELETE", "/providers/microsoft", BY)).status, 404);
});

test("connecting Microsoft to another organization disconnects its apps; the same one keeps them", async () => {
  const h = harness();
  // The fixture's app is client id TENANT.
  assert.equal(
    (await connectMicrosoft(h, { client_secret: "new", client_id: TENANT })).status,
    200,
  );
  assert.equal(h.state.connections.length, 5);
  assert.equal(
    (await connectMicrosoft(h, { client_secret: "x", tenant_id: OTHER_TENANT })).status,
    200,
  );
  assert.deepEqual(
    h.state.savedProviders.map((p) => p.removeApps),
    [false, true],
  );
  assert.deepEqual(
    h.state.connections.map((c) => c.connectorId),
    ["jira"],
  );
});

test("switching to another app in the same organization disconnects the apps too", async () => {
  const h = harness();
  assert.equal((await connectMicrosoft(h, { client_secret: "new" })).status, 200);
  assert.deepEqual(
    h.state.savedProviders.map((p) => p.removeApps),
    [true],
  );
  assert.deepEqual(
    h.state.connections.map((c) => c.connectorId),
    ["jira"],
  );
  assert.ok(!h.state.grants.some((g) => g.connectorId !== "jira"));
});

test("sameSignIn: account and app must match; a key missing on old rows is not a change", () => {
  const sp = { auth: "client_secret", tenantId: TENANT, clientId: CLIENT, servicePrincipalId: "a" };
  assert.ok(sameSignIn(microsoftProvider, sp, { ...sp, auth: "certificate" }));
  assert.ok(!sameSignIn(microsoftProvider, sp, { ...sp, clientId: OTHER_TENANT }));
  assert.ok(!sameSignIn(microsoftProvider, sp, { ...sp, servicePrincipalId: "b" }));
  assert.ok(!sameSignIn(microsoftProvider, sp, { ...sp, tenantId: OTHER_TENANT }));
  // Moved from microsoft-graph (no service principal) or microsoft-365 (no client id).
  const { servicePrincipalId: _s, ...graph } = sp;
  assert.ok(sameSignIn(microsoftProvider, graph, sp));
  const shared = { auth: "arbium_app", tenantId: TENANT, servicePrincipalId: "shared" };
  assert.ok(!sameSignIn(microsoftProvider, shared, sp));
});

test("no Microsoft tools while the Microsoft connection needs reconnecting", async () => {
  const h = harness();
  const key = h.addAgent(["read"], { mailbox: "agent@x.com" }, "active", "outlook-mail");
  const names = async () =>
    ((await h.rpc(key, "tools/list")).body.result.tools as { name: string }[]).length;
  assert.equal(await names(), 2);
  h.msp.refuse = () => true;
  const checked = await admin(h, "POST", "/providers/microsoft/check");
  assert.equal(checked.body.ok, false);
  assert.equal(h.state.providers[0].status, "needs_reconnect");
  assert.equal(await names(), 0);
});

test("a check that fails after a reconnect leaves the new credential connected", async () => {
  const h = harness();
  h.addAgent(["read"], { mailbox: "agent@x.com" }, "active", "outlook-mail");
  h.msp.refuse = () => {
    // An admin reconnects while this check's sign-in with the old credential is in flight.
    h.state.providers[0] = { ...h.state.providers[0], credential: Buffer.from("reconnected") };
    return true;
  };
  const checked = await admin(h, "POST", "/providers/microsoft/check");
  assert.equal(checked.body.ok, false);
  assert.equal(h.state.providers[0].status, "connected");
});

test("a credential moved from an app's row is re-sealed under the provider on the next check", async () => {
  const h = harness();
  const row = h.state.providers[0];
  const sealed = h.key.seal({ client_secret: "x" }, credentialAad(ORG, "sharepoint-files"));
  Object.assign(row, {
    credential: sealed.ciphertext,
    credentialKeyId: sealed.keyId,
    credentialScope: "sharepoint-files",
  });
  const checked = await admin(h, "POST", "/connections/sharepoint-files/check", BY);
  assert.equal(checked.body.ok, true);
  assert.equal(row.credentialScope, "provider:microsoft");
  assert.equal(
    h.key.open(
      { ciphertext: row.credential, keyId: row.credentialKeyId },
      credentialAad(ORG, "provider:microsoft"),
    ).client_secret,
    "x",
  );
});

test("a shared-app row moved from microsoft-365 shows the shared app's client id", async () => {
  const h = harness({ shared: true });
  h.state.providers[0].config = { auth: "arbium_app", tenantId: TENANT, servicePrincipalId: "s" };
  const listed = await admin(h, "GET", "/providers");
  assert.equal(listed.body[0].config.clientId, "app-123");
  // Only rows without one, and never the customer's own app.
  const own = harness();
  own.state.providers[0].config.clientId = CLIENT;
  assert.equal((await admin(own, "GET", "/providers")).body[0].config.clientId, CLIENT);
});

test("setup steps an admin confirms are kept for the organization, and only known ones", async () => {
  const h = fresh();
  assert.equal((await admin(h, "PUT", "/confirmed/mailbox-group", BY)).status, 204);
  assert.equal((await admin(h, "PUT", "/confirmed/mailbox-group", BY)).status, 204);
  assert.equal((await admin(h, "PUT", "/confirmed/microsoft-app-registered", BY)).status, 204);
  const listed = await admin(h, "GET", "/confirmed");
  assert.deepEqual(
    listed.body.map((c: { step: string; by: string }) => [c.step, c.by]),
    [
      ["mailbox-group", "admin@x"],
      ["microsoft-app-registered", "admin@x"],
    ],
  );
  // A step two apps share is audited under the first of them.
  assert.equal(
    h.state.audit.find((r) => r.target === "mailbox-group")!.connectorId,
    "outlook-mail",
  );
  assert.equal((await admin(h, "PUT", "/confirmed/not-a-step", BY)).status, 400);
  assert.equal((await admin(h, "DELETE", "/confirmed/mailbox-group", BY)).status, 204);
  assert.deepEqual(
    (await admin(h, "GET", "/confirmed")).body.map((c: { step: string }) => c.step),
    ["microsoft-app-registered"],
  );
});

test("settings an app's steps ask for are kept on the provider, checked and audited", async () => {
  const h = fresh();
  assert.equal((await connectMicrosoft(h, { client_secret: "s3cret" })).status, 200);
  const sub = "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa";
  const saved = await admin(h, "PUT", "/providers/microsoft/settings", {
    values: { azureSubscriptionId: ` ${sub} `, azureResourceGroup: "example-agents" },
    ...BY,
  });
  assert.equal(saved.status, 200);
  assert.equal(h.state.providers[0].config.azureSubscriptionId, sub);
  assert.equal(h.state.audit.at(-1)!.tool, "admin:settings");
  for (const values of [
    { tenantId: "x" }, // not a setting: the sign-in's own config is never set this way
    { azureSubscriptionId: "not-a-guid" },
    { azureResourceGroup: "a group; rm -rf" },
    { azureResourceGroup: "ends-with-dot." },
  ]) {
    const res = await admin(h, "PUT", "/providers/microsoft/settings", { values, ...BY });
    assert.equal(res.status, 400, JSON.stringify(values));
  }
  // The same sign-in again keeps them.
  assert.equal((await connectMicrosoft(h, { client_secret: "s3cret" })).status, 200);
  assert.equal(h.state.providers[0].config.azureResourceGroup, "example-agents");
});

test("the Microsoft permission check names only what is missing", () => {
  const expected = ["User.Read.All", "GroupMember.Read.All"];
  assert.equal(
    compareRoles(["User.Read.All", "GroupMember.Read.All", "Sites.Read.All"], { expected }),
    undefined,
  );
  assert.match(
    compareRoles(["User.Read.All", "Sites.Read.All"], { expected }) ?? "",
    /^Missing: GroupMember\.Read\.All\./,
  );
});
