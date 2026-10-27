import assert from "node:assert/strict";
import { randomUUID } from "node:crypto";
import { after, before, test } from "node:test";
import { migrate } from "../src/migrate.ts";
import { assertSchemaName, connectDatabase, createStore, ProviderChanged } from "../src/store.ts";
import { ConnectError } from "../../connector-kit/src/types.ts";

// Runs against the scratch database scripts/test-databases.sh makes, migrated here:
// TEST_DATABASE_URL=postgres://…/ac_test DATABASE_SSL=disable npm test. Skipped without one.
const url = process.env.TEST_DATABASE_URL;
const sql = url ? connectDatabase(url) : null;
const SCHEMA = "agent_connect";
before(async () => {
  if (sql) await migrate(sql, SCHEMA, () => {});
});
after(() => sql?.end());

const P = Buffer.from("p");
const Q = Buffer.from("q");
const hash = () => randomUUID().replaceAll("-", "").repeat(2);
// The audit row an admin's key change is written with.
const keyAudit = (org: string, target = "key") => ({
  organizationId: org,
  agentId: null,
  agentName: "admin",
  connectorId: "agents",
  tool: "admin:create_key",
  target,
  decision: "allowed" as const,
  reason: null,
  vendorStatus: null,
});
const made = (a: { organizationId: string; id: string }) => keyAudit(a.organizationId, a.id);

async function org() {
  const id = randomUUID();
  await createStore(sql!, SCHEMA).ensureOrganization(id);
  return id;
}

const connection = (
  organizationId: string,
  config: Record<string, string>,
  clearGrants = false,
) => ({
  organizationId,
  connectorId: "outlook-mail",
  capabilities: ["read"],
  config,
  credential: Buffer.from("x"),
  credentialKeyId: "k",
  providerId: null,
  accountLabel: "X",
  connectedBy: "admin",
  clearGrants,
});

const provider = (
  organizationId: string,
  config: Record<string, string>,
  removeApps = false,
  replaces: Buffer | null = null,
) => ({
  organizationId,
  provider: "microsoft",
  config,
  credential: Buffer.from("p"),
  credentialKeyId: "k",
  credentialScope: "provider:microsoft",
  accountLabel: "T",
  connectedBy: "admin",
  removeApps,
  replaces,
});

test(
  "a tenant approved for the shared app binds to one org; a customer's own app does not",
  { skip: !sql },
  async () => {
    const store = createStore(sql!, SCHEMA);
    const [a, b] = [await org(), await org()];
    const tenant = randomUUID();
    await store.saveProvider(provider(a, { auth: "arbium_app", tenantId: tenant }));
    await assert.rejects(
      store.saveProvider(provider(b, { auth: "arbium_app", tenantId: tenant })),
      ConnectError,
    );
    await store.saveProvider(provider(b, { auth: "client_secret", tenantId: tenant }));
    assert.equal((await store.getProvider(b, "microsoft"))!.status, "connected");
  },
);

test("a provider's apps share its credential, and go with it", { skip: !sql }, async () => {
  const store = createStore(sql!, SCHEMA);
  const a = await org();
  // A certificate made before connecting is kept, but never replaces a connected row's.
  assert.equal(
    await store.savePendingProvider({
      ...provider(a, { certificate: "c1" }),
      credentialScope: "s",
    }),
    true,
  );
  assert.equal((await store.getProvider(a, "microsoft"))!.status, "pending");
  // Saving over a row needs the credential read from it.
  await assert.rejects(
    store.saveProvider(provider(a, { auth: "certificate", tenantId: "t1" })),
    ProviderChanged,
  );
  await store.saveProvider(provider(a, { auth: "certificate", tenantId: "t1" }, false, P));
  assert.equal(
    await store.savePendingProvider({
      ...provider(a, { certificate: "c2" }),
      credentialScope: "s",
    }),
    false,
  );
  const p = (await store.getProvider(a, "microsoft"))!;
  assert.equal(p.config.tenantId, "t1");
  assert.equal(p.credentialScope, "provider:microsoft");

  // An app row holds either its own credential or its provider, never both or neither.
  const tile = { ...connection(a, {}), credential: null, credentialKeyId: null, providerId: p.id };
  await assert.rejects(
    store.saveConnection({ ...tile, credential: Buffer.from("x"), credentialKeyId: "k" }),
  );
  await assert.rejects(store.saveConnection({ ...tile, providerId: null }));
  await store.saveConnection(tile);
  await store.saveConnection({ ...tile, connectorId: "sharepoint-files", config: { sites: "s" } });
  const listed = await store.listConnections(a);
  assert.deepEqual(
    listed.map((c) => [c.connectorId, c.providerId, c.credential]),
    [
      ["outlook-mail", p.id, null],
      ["sharepoint-files", p.id, null],
    ],
  );

  const agent = await store.createAgent(
    {
      organizationId: a,
      name: "Both",
      keyHash: hash(),
      createdBy: "admin",
    },
    made,
  );
  await store.saveConnection({ ...connection(a, {}), connectorId: "jira" });
  await store.setGrant(a, agent!.id, "outlook-mail", ["read"], {});
  await store.setGrant(a, agent!.id, "jira", ["read"], {});
  // Above the app's ceiling, or for an app not connected, nothing is saved.
  assert.equal(await store.setGrant(a, agent!.id, "jira", ["read", "comment"], {}), false);
  assert.equal(await store.setGrant(a, agent!.id, "fedex", ["track"], {}), false);
  const rotate = (replaces: Buffer) =>
    store.updateProviderCredential(
      p.id,
      { ...p.config, nextCertificate: "n" },
      Buffer.from("q"),
      "k2",
      "provider:microsoft",
      replaces,
    );
  assert.equal(await rotate(P), true);
  // A second write from the same stale read changes nothing.
  assert.equal(await rotate(P), false);
  assert.equal((await store.getProvider(a, "microsoft"))!.config.nextCertificate, "n");
  // A check of the credential that was replaced changes nothing either.
  await store.markProviderChecked(p.id, false, P);
  assert.equal((await store.getProvider(a, "microsoft"))!.status, "connected");
  await store.markProviderChecked(p.id, false, Buffer.from("q"));
  assert.equal((await store.getProvider(a, "microsoft"))!.status, "needs_reconnect");
  // A credential just proven by a sign-in (a certificate switched to) connects it again.
  const now = (await store.getProvider(a, "microsoft"))!;
  assert.equal(
    await store.updateProviderCredential(
      p.id,
      now.config,
      now.credential,
      "k2",
      "provider:microsoft",
      now.credential,
      true,
    ),
    true,
  );
  assert.equal((await store.getProvider(a, "microsoft"))!.status, "connected");
  await store.markProviderChecked(p.id, false, now.credential);

  // Reaching another tenant disconnects the provider's apps.
  // A stale save neither replaces the credential nor disconnects the apps.
  await assert.rejects(
    store.saveProvider(provider(a, { auth: "certificate", tenantId: "t2" }, true, P)),
    ProviderChanged,
  );
  assert.equal((await store.listConnections(a)).length, 3);
  await store.saveProvider(provider(a, { auth: "certificate", tenantId: "t2" }, true, Q));
  assert.deepEqual(
    (await store.listConnections(a)).map((c) => c.connectorId),
    ["jira"],
  );
  assert.deepEqual(
    (await store.grantsFor(agent!.id)).map((g) => g.connectorId),
    ["jira"],
  );

  await store.saveConnection(tile);
  await store.setGrant(a, agent!.id, "outlook-mail", ["read"], {});
  assert.equal(await store.removeProvider(a, "microsoft"), true);
  assert.deepEqual(
    (await store.listConnections(a)).map((c) => c.connectorId),
    ["jira"],
  );
  assert.deepEqual(
    (await store.grantsFor(agent!.id)).map((g) => g.connectorId),
    ["jira"],
  );
  assert.equal(await store.getProvider(a, "microsoft"), null);
});

test("store clears grants with the account, and keeps audit names", { skip: !sql }, async () => {
  const store = createStore(sql!, SCHEMA);
  const a = await org();
  const tenant = randomUUID();
  await store.saveConnection(connection(a, { tenantId: tenant }));

  const conn = await store.getConnection(a, "outlook-mail");
  assert.match(conn!.connectedAt, /^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/);
  assert.deepEqual(conn!.capabilities, ["read"]);
  assert.equal(await store.setCeiling(a, "outlook-mail", ["read", "delete"]), true);
  assert.deepEqual((await store.getConnection(a, "outlook-mail"))!.capabilities, [
    "read",
    "delete",
  ]);

  const agent = await store.createAgent(
    {
      organizationId: a,
      name: "Mailer",
      keyHash: hash(),
      createdBy: "admin",
    },
    made,
  );
  assert.match(agent!.createdAt, /Z$/);
  await store.setGrant(a, agent!.id, "outlook-mail", ["read"], { mailbox: "x@y.com" });
  // Same tenant again: grants stay. A different one: they go.
  await store.saveConnection(connection(a, { tenantId: tenant }, false));
  assert.equal((await store.grantsFor(agent!.id)).length, 1);
  assert.match((await store.grantsFor(agent!.id))[0].updatedAt, /Z$/);
  assert.deepEqual((await store.grantsFor(agent!.id))[0].capabilities, ["read"]);
  // Lowering the ceiling takes the capability off the agent; raising it gives nothing back.
  await store.setCeiling(a, "outlook-mail", ["read", "send", "delete"]);
  await store.setGrant(a, agent!.id, "outlook-mail", ["read", "send", "delete"], {});
  await store.setCeiling(a, "outlook-mail", ["read", "send"]);
  await store.setCeiling(a, "outlook-mail", ["read", "send", "delete"]);
  assert.deepEqual((await store.grantsFor(agent!.id))[0].capabilities, ["read", "send"]);
  await store.saveConnection(connection(a, { tenantId: tenant }, false));
  assert.deepEqual((await store.grantsFor(agent!.id))[0].capabilities, ["read"]);
  await store.saveConnection(connection(a, { tenantId: randomUUID() }, true));
  assert.equal((await store.grantsFor(agent!.id)).length, 0);

  await store.setGrant(a, agent!.id, "outlook-mail", ["read"], { mailbox: "x@y.com" });
  assert.equal(await store.removeConnection(a, "outlook-mail"), true);
  assert.equal((await store.grantsFor(agent!.id)).length, 0);

  await store.audit({
    organizationId: a,
    agentId: agent!.id,
    agentName: "Mailer",
    connectorId: "outlook-mail",
    tool: "mail_send",
    target: null,
    decision: "failed",
    reason: "Microsoft 365 answered 500",
    vendorStatus: 500,
  });
  const [listed] = await store.listAgents(a);
  assert.match(listed.lastUsedAt!, /Z$/);
  const other = await store.createAgent(
    {
      organizationId: a,
      name: "Idle",
      keyHash: hash(),
      createdBy: "admin",
    },
    made,
  );
  assert.equal((await store.listAgents(a)).find((x) => x.id === other!.id)!.lastUsedAt, null);
  await store.deleteAgent(a, agent!.id);
  const [row] = await store.recentAudit(a, 10, { agentId: null, connectorId: "outlook-mail" });
  assert.equal(row.agentName, "Mailer");
  assert.equal(row.decision, "failed");
  assert.match(row.at, /Z$/);
});

test("sites added after connecting change only the app's settings", { skip: !sql }, async () => {
  const store = createStore(sql!, SCHEMA);
  const a = await org();
  await store.saveConnection(connection(a, { access: "selected" }));
  assert.equal(
    await store.updateConfig(
      a,
      "outlook-mail",
      { access: "selected", sites: "s" },
      {
        access: "selected",
      },
    ),
    true,
  );
  const row = (await store.getConnection(a, "outlook-mail"))!;
  assert.deepEqual(row.config, { access: "selected", sites: "s" });
  assert.equal(row.status, "connected");
  assert.equal(await store.updateConfig(a, "fedex", { sites: "s" }, {}), false);
  // Written over a config that changed since it was read (a reconnect): refused.
  assert.equal(
    await store.updateConfig(a, "outlook-mail", { sites: "t" }, { access: "selected" }),
    false,
  );
  assert.deepEqual((await store.getConnection(a, "outlook-mail"))!.config, {
    access: "selected",
    sites: "s",
  });
  // A re-seal only replaces the credential it read.
  const stale = Buffer.from("old");
  assert.equal(await store.reseal(row.id, stale, Buffer.from("new"), "k2"), false);
  assert.equal(await store.reseal(row.id, row.credential!, Buffer.from("new"), "k2"), true);
  assert.equal((await store.getConnection(a, "outlook-mail"))!.credentialKeyId, "k2");
});

test("confirmed setup steps are kept once per organization", { skip: !sql }, async () => {
  const store = createStore(sql!, SCHEMA);
  const o = await org();
  await store.setConfirmed(o, "mailbox-group", "a", true);
  await store.setConfirmed(o, "mailbox-group", "b", true);
  const rows = await store.listConfirmed(o);
  assert.deepEqual(
    rows.map((r) => [r.step, r.by]),
    [["mailbox-group", "a"]],
  );
  await store.setConfirmed(o, "mailbox-group", "a", false);
  assert.equal((await store.listConfirmed(o)).length, 0);
});

test("a key change whose audit row fails is not kept", { skip: !sql }, async () => {
  const store = createStore(sql!, SCHEMA);
  const a = await org();
  // An audit row for an organization that does not exist is refused by its foreign key.
  const refused = keyAudit(randomUUID());
  const input = { organizationId: a, name: "Lost", keyHash: hash(), createdBy: "admin" };
  await assert.rejects(store.createAgent(input, () => refused));
  assert.deepEqual(await store.listAgents(a), []);

  const agent = (await store.createAgent(input, made))!;
  const fresh = hash();
  await assert.rejects(store.updateAgent(a, agent.id, { keyHash: fresh }, refused));
  assert.equal((await store.agentByKeyHash(input.keyHash))?.id, agent.id);
  assert.equal(await store.agentByKeyHash(fresh), null);
  assert.equal(await store.updateAgent(a, agent.id, { keyHash: fresh }, keyAudit(a)), true);
  assert.equal((await store.agentByKeyHash(fresh))?.id, agent.id);
  assert.deepEqual(
    (await store.recentAudit(a, 10, { agentId: null, connectorId: "agents" })).length,
    2,
  );
});

test(
  "organizations are made by id or by name, and a name is kept once given",
  { skip: !sql },
  async () => {
    const store = createStore(sql!, SCHEMA);
    const id = randomUUID();
    assert.equal(await store.getOrganization(id), null);
    assert.equal((await store.ensureOrganization(id)).name, null);
    assert.equal((await store.ensureOrganization(id, "Acme")).name, "Acme");
    assert.equal((await store.ensureOrganization(id)).name, "Acme");
    // Without a name, an existing row is only read (the legacy prefix ensures on every request).
    const version = async () =>
      (await sql!`select xmin::text from agent_connect.organizations where id = ${id}`)[0].xmin;
    const was = await version();
    await store.ensureOrganization(id);
    assert.equal(await version(), was);
    const named = await store.createOrganization("Beta");
    assert.match(named.createdAt, /Z$/);
    assert.deepEqual(await store.getOrganization(named.id), named);
    const ids = (await store.listOrganizations()).map((o) => o.id);
    assert.ok(ids.includes(id) && ids.includes(named.id));
  },
);

test("the schema is the service's own, never public or a system one", () => {
  assertSchemaName("agent_connect");
  for (const schema of ["public", "pg_catalog", "pg_x", "Agent", "a-b"])
    assert.throws(() => assertSchemaName(schema), /DATABASE_SCHEMA/);
});
