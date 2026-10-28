import assert from "node:assert/strict";
import { test } from "node:test";
import { microsoftClient } from "../microsoft/client.ts";
import { ConnectError } from "../../packages/connector-kit/src/types.ts";
import { dynamicsVendor, normalizeEnvironment } from "./vendor.ts";

// The Dynamics 365 sign-in through the Microsoft connection, against a fake Microsoft: a token
// for the environment itself, WhoAmI to prove the application user, and the security role check.

const TENANT = "11111111-1111-1111-1111-111111111111";
const CLIENT = "22222222-2222-2222-2222-222222222222";
const USER = "33333333-3333-3333-3333-333333333333";
const ENV = "https://contoso.crm4.dynamics.com";
const PROVIDER = {
  config: { auth: "client_secret", tenantId: TENANT, clientId: CLIENT },
  secret: { client_secret: "s3cret" },
  accountLabel: TENANT,
};

type Options = { whoAmI?: number; privileges?: number; held?: string[] };

function fakeMicrosoft(options: Options = {}) {
  const scopes: string[] = [];
  const calls: { url: string; headers: Record<string, string> }[] = [];
  const fetchImpl = (async (input: string | URL, init: RequestInit = {}) => {
    const url = String(input);
    if (url.startsWith("https://login.microsoftonline.com/")) {
      scopes.push(new URLSearchParams(String(init.body)).get("scope")!);
      return Response.json({ access_token: "at", expires_in: 3600 });
    }
    calls.push({ url, headers: init.headers as Record<string, string> });
    const path = new URL(url).pathname;
    if (path.endsWith("/WhoAmI"))
      return options.whoAmI
        ? Response.json(
            {
              error: {
                code: "0x80072560",
                message: "The user is not a member of the organization.",
              },
            },
            { status: options.whoAmI },
          )
        : Response.json({ UserId: USER, BusinessUnitId: USER });
    if (path.endsWith("RetrieveUserPrivileges()"))
      return options.privileges
        ? Response.json(
            { error: { message: "missing prvReadUser" } },
            { status: options.privileges },
          )
        : Response.json({
            RolePrivileges: (
              options.held ?? ["prvReadUser", "prvReadAccount", "prvDeleteAccount"]
            ).map((p) => ({ PrivilegeName: p, Depth: "Global" })),
          });
    if (path.endsWith(`/systemusers(${USER})`))
      return Response.json({
        fullname: "# Agents",
        systemuserroles_association: [{ name: "Agents" }, { name: "Basic User" }],
      });
    return Response.json({}, { status: 404 });
  }) as typeof fetch;
  const vendor = dynamicsVendor(microsoftClient({ sharedApp: null }, fetchImpl));
  return { vendor, scopes, calls };
}

test("normalizeEnvironment keeps the environment's origin and refuses other hosts", () => {
  for (const [input, want] of [
    ["contoso.crm.dynamics.com", "https://contoso.crm.dynamics.com"],
    ["https://Contoso.crm4.dynamics.com/main.aspx?appid=1", "https://contoso.crm4.dynamics.com"],
    ["https://contoso.api.crm11.dynamics.com/api/data/v9.2/", "https://contoso.crm11.dynamics.com"],
    ["https://gov-org.crm9.dynamics.com", "https://gov-org.crm9.dynamics.com"],
  ])
    assert.equal(normalizeEnvironment(input), want, input);
  for (const bad of [
    "",
    "http://contoso.crm.dynamics.com",
    "https://contoso.crm.dynamics.com:8443",
    "https://contoso.crm.dynamics.com.evil.example",
    "https://evil.example/contoso.crm.dynamics.com",
    "https://contoso.sharepoint.com",
    "https://crm.dynamics.com",
  ])
    assert.throws(() => normalizeEnvironment(bad), ConnectError, bad);
  for (const sovereign of [
    "https://agency.crm.microsoftdynamics.us",
    "https://agency.crm.appsplatform.us",
    "https://contoso.crm.dynamics.cn",
  ])
    assert.throws(() => normalizeEnvironment(sovereign), /not supported/, sovereign);
});

test("connect signs in to the environment itself and proves the application user", async () => {
  const { vendor, scopes, calls } = fakeMicrosoft();
  const conn = await vendor.connect({ environment: "contoso.crm4.dynamics.com/" }, PROVIDER);
  assert.deepEqual(conn, {
    config: { environment: ENV },
    secret: {},
    accountLabel: "contoso.crm4.dynamics.com",
  });
  assert.deepEqual(scopes, [`${ENV}/.default`]);
  assert.equal(calls[0].url, `${ENV}/api/data/v9.2/WhoAmI`);
  assert.equal(calls[0].headers["OData-Version"], "4.0");
});

test("connect explains a missing application user", async () => {
  const { vendor } = fakeMicrosoft({ whoAmI: 403 });
  await assert.rejects(vendor.connect({ environment: ENV }, PROVIDER), {
    name: "Error",
    message:
      "Dynamics 365 answered 403: The user is not a member of the organization. Add the app to contoso.crm4.dynamics.com as an application user with the agents' security role (the steps above), then try again. A new application user can take a few minutes to apply.",
  });
  await assert.rejects(
    vendor.connect({ environment: "https://example.com" }, PROVIDER),
    ConnectError,
  );
});

test("calls go only to the connected environment", async () => {
  const { vendor, calls } = fakeMicrosoft();
  const conn = { config: { ...PROVIDER.config, environment: ENV }, secret: PROVIDER.secret };
  await vendor.call(conn, { method: "GET", path: "/api/data/v9.2/accounts" });
  assert.equal(calls[0].url, `${ENV}/api/data/v9.2/accounts`);
  await assert.rejects(
    vendor.call(
      { ...conn, config: { ...conn.config, environment: "https://evil.example" } },
      { method: "GET", path: "/api/data/v9.2/accounts" },
    ),
    ConnectError,
  );
});

test("the access check lists the roles and warns only about missing privileges", async () => {
  const conn = { config: { ...PROVIDER.config, environment: ENV }, secret: PROVIDER.secret };
  const { vendor } = fakeMicrosoft();
  const ok = await vendor.checks!.access(conn, { expected: ["prvReadUser", "prvReadAccount"] });
  assert.deepEqual(ok, { values: ["Agents", "Basic User"] });
  const missing = await vendor.checks!.access(conn, {
    expected: ["prvReadUser", "prvReadAccount", "prvWriteAccount", "prvReadContact"],
  });
  assert.equal(
    missing.warning,
    "Its security roles lack prvWriteAccount, prvReadContact. Add them to the role, or agents cannot do what you ticked.",
  );
  const unreadable = await fakeMicrosoft({ privileges: 403 }).vendor.checks!.access(conn, {
    expected: ["prvReadUser"],
  });
  assert.deepEqual(unreadable.values, []);
  assert.match(unreadable.warning!, /^Dynamics 365 answered 403: missing prvReadUser when reading/);
});
