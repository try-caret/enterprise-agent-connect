import assert from "node:assert/strict";
import { generateKeyPairSync } from "node:crypto";
import { test } from "node:test";
import { audited, type ToolContext } from "../../connector-kit/src/common.ts";
import { loadConnectors } from "../../connector-kit/src/connector.ts";
import { retryAfter, vendorFetch, vendorUrl } from "../../connector-kit/src/fetch.ts";
import { clientCredentials } from "../../connector-kit/src/oauth.ts";
import {
  ConnectError,
  type VendorConnection,
  type VendorRequest,
} from "../../connector-kit/src/types.ts";
import { microsoftClient } from "../../../connectors/microsoft/client.ts";
import { builtinConnectors } from "../src/builtin.ts";

const GUID = "11111111-2222-3333-4444-555555555555";
const microsoftConn = {
  config: { auth: "client_secret", tenantId: GUID, clientId: GUID },
  secret: { client_secret: "s" },
};
const okta = generateKeyPairSync("rsa", { modulusLength: 2048 }).privateKey.export({
  type: "pkcs8",
  format: "pem",
}) as string;

// A connection and a request each built-in app accepts. A new app must be added here.
const calls: Record<string, { conn: VendorConnection; request: VendorRequest }> = {
  jira: {
    conn: { config: { cloudId: GUID }, secret: { client_id: "i", client_secret: "s" } },
    request: { method: "GET", path: "/rest/api/3/myself" },
  },
  salesforce: {
    conn: {
      config: { myDomain: "acme.my.salesforce.com", instanceUrl: "https://acme.my.salesforce.com" },
      secret: { client_id: "i", client_secret: "s" },
    },
    request: { method: "GET", path: "/services/data/v62.0/limits" },
  },
  servicenow: {
    conn: {
      config: { instance: "acme.service-now.com", auth: "oauth" },
      secret: { client_id: "i", client_secret: "s" },
    },
    request: { method: "GET", path: "/api/now/table/incident" },
  },
  fedex: {
    conn: { config: { environment: "sandbox" }, secret: { api_key: "k", secret_key: "s" } },
    request: { method: "POST", path: "/track/v1/trackingnumbers", body: {} },
  },
  "dell-techdirect": {
    conn: { config: {}, secret: { client_id: "i", client_secret: "s" } },
    request: { method: "GET", path: "/x" },
  },
  okta: {
    conn: {
      config: { org: "acme.okta.com", clientId: "0oa1", keyId: "k1" },
      secret: { private_key: okta },
    },
    request: { method: "GET", path: "/api/v1/users" },
  },
  slack: {
    conn: { config: {}, secret: { bot_token: "xoxb-1" } },
    request: { method: "GET", path: "/auth.test" },
  },
  jamf: {
    conn: {
      config: { url: "https://acme.jamfcloud.com" },
      secret: { client_id: "i", client_secret: "s" },
    },
    request: { method: "GET", path: "/api/v1/sites" },
  },
  "sap-s4hana": {
    conn: {
      config: { host: "sap.acme.example", user: "u", auth: "oauth" },
      secret: { password: "p" },
    },
    request: { method: "GET", path: "/sap/opu/odata/sap/API_BUSINESS_PARTNER/A_BusinessPartner" },
  },
  workday: {
    conn: {
      config: { host: "wd2-impl-services1.workday.com", tenant: "acme" },
      secret: { client_id: "i", client_secret: "s", refresh_token: "r" },
    },
    request: { method: "GET", path: "/ccx/api/staffing/v7/acme/workers" },
  },
  netsuite: {
    conn: {
      config: { accountId: "1234567", auth: "tba" },
      secret: { client_id: "i", client_secret: "s", token_id: "t", token_secret: "u" },
    },
    request: { method: "GET", path: "/services/rest/record/v1/customer" },
  },
  "dynamics-365": {
    conn: {
      config: { ...microsoftConn.config, environment: "https://acme.crm.dynamics.com" },
      secret: microsoftConn.secret,
    },
    request: { method: "GET", path: "/api/data/v9.2/WhoAmI" },
  },
};
for (const id of [
  "outlook-mail",
  "outlook-calendar",
  "sharepoint-files",
  "teams",
  "entra",
  "intune",
])
  calls[id] = { conn: microsoftConn, request: { method: "GET", path: "/v1.0/organization" } };

const TOKEN = {
  access_token: "t",
  expires_in: 3600,
  instance_url: "https://acme.my.salesforce.com",
  id: "https://login.salesforce.com/id/00D/005000000000001AAA",
};

// Token requests are answered; every other call is redirected to another host.
function redirecting(token: boolean) {
  const seen: { url: string; redirect?: RequestInit["redirect"] }[] = [];
  const fake = (async (input: string | URL, init?: RequestInit) => {
    const url = String(input);
    seen.push({ url, redirect: init?.redirect });
    if (token && /token/.test(new URL(url).pathname)) return Response.json(TOKEN);
    return new Response(null, { status: 302, headers: { location: "https://evil.example/" } });
  }) as typeof fetch;
  return { seen, fake };
}

test("no built-in app follows a redirect, at sign-in or on a call", async () => {
  const connectors = loadConnectors();
  const real = globalThis.fetch;
  try {
    for (const module of builtinConnectors) {
      const call = calls[module.id];
      assert.ok(call, `Add ${module.id} to this test's calls`);
      for (const token of [false, true]) {
        const { seen, fake } = redirecting(token);
        globalThis.fetch = fake;
        const { vendor } = module.create({
          microsoft: microsoftClient({ sharedApp: null }, fake),
          connector: connectors.find((c) => c.id === module.id)!,
        });
        await assert.rejects(
          vendor.call(call.conn, call.request),
          (e: Error) => e instanceof ConnectError && /redirect/.test(e.message),
          `${module.id} (token answered: ${token})`,
        );
        assert.ok(seen.length, module.id);
        assert.ok(
          seen.every((s) => s.redirect === "manual" && !s.url.includes("evil")),
          `${module.id}: ${JSON.stringify(seen)}`,
        );
      }
    }
  } finally {
    globalThis.fetch = real;
  }
});

test("vendorFetch refuses redirects and waits out a short Retry-After", async () => {
  const waits: number[] = [];
  const answers = [
    new Response(null, { status: 429, headers: { "retry-after": "3" } }),
    new Response("ok"),
    new Response(null, { status: 429, headers: { "retry-after": "60" } }),
    new Response(null, { status: 301, headers: { location: "https://evil.example/" } }),
  ];
  const inits: RequestInit[] = [];
  const fake = (async (_: string, init: RequestInit) => {
    inits.push(init);
    return answers.shift()!;
  }) as typeof fetch;
  const safe = vendorFetch("Acme", fake, async (s) => void waits.push(s));
  assert.equal(vendorFetch("Acme", safe), safe);
  assert.equal(await (await safe("https://acme.example/a")).text(), "ok");
  assert.deepEqual(waits, [3]);
  assert.equal((await safe("https://acme.example/b")).status, 429);
  await assert.rejects(safe("https://acme.example/c"), /Acme answered with a redirect \(301\)/);
  assert.ok(inits.every((i) => i.redirect === "manual"));
  assert.ok((retryAfter(new Date(Date.now() + 5000).toUTCString()) ?? 0) >= 4);
  assert.equal(retryAfter("soon"), null);
});

test("calls stay on the app's own origin", async () => {
  assert.equal(vendorUrl("https://acme.example", "/a?b"), "https://acme.example/a?b");
  for (const path of ["@evil.example/x", ".evil.example/x", ":443@evil.example"])
    assert.throws(() => vendorUrl("https://acme.example", path), /leaving acme\.example/);
  const { fake, seen } = redirecting(true);
  const oauth = clientCredentials("Acme", fake);
  const creds = { tokenUrl: "https://acme.example/token", clientId: "i", clientSecret: "s" };
  await assert.rejects(
    oauth.call(creds, "https://acme.example", { method: "GET", path: "@evil.example" }),
    /leaving/,
  );
  await assert.rejects(
    oauth.call(creds, "https://acme.example", { method: "GET", path: "/x" }),
    /redirect \(302\)/,
  );
  assert.ok(seen.every((s) => s.redirect === "manual"));
});

test("an app's ConnectError reaches the agent as a failed call, audited", async () => {
  const audit: { decision: string; reason: string | null }[] = [];
  const ctx = {
    audit: async (e: (typeof audit)[number]) => void audit.push(e),
  } as unknown as ToolContext;
  const r = await audited(ctx, "acme_read", null, async () => {
    throw new ConnectError("Acme answered with a redirect (302), which is not followed.");
  });
  assert.equal(r.isError, true);
  assert.match(JSON.stringify(r.content), /redirect \(302\)/);
  assert.deepEqual(audit, [
    {
      tool: "acme_read",
      target: null,
      decision: "failed",
      reason: "Acme answered with a redirect (302), which is not followed.",
      vendorStatus: 502,
    },
  ]);
});

test("a file's content follows Graph's download address only to SharePoint, without the token", async () => {
  const seen: { url: string; auth?: string }[] = [];
  const to = { location: "https://acme.sharepoint.com/_layouts/15/download.aspx?x=1" };
  const fake = (async (input: string, init?: RequestInit) => {
    seen.push({ url: input, auth: (init?.headers as Record<string, string>)?.authorization });
    if (input.includes("/oauth2/")) return Response.json(TOKEN);
    if (input.includes("sharepoint.com")) return new Response("file text");
    return new Response(null, { status: 302, headers: to });
  }) as typeof fetch;
  const ms = microsoftClient({ sharedApp: null }, fake);
  const read = {
    method: "GET",
    path: "/v1.0/drives/d/items/i/content",
    responseType: "text",
  } as const;
  assert.equal((await ms.call(microsoftConn, read)).data, "file text");
  assert.deepEqual(seen.at(-1), { url: to.location, auth: undefined });
  to.location = "https://evil.example/x";
  await assert.rejects(ms.call(microsoftConn, read), /not SharePoint: evil\.example/);
  await assert.rejects(
    ms.call(microsoftConn, { method: "GET", path: "/v1.0/me" }),
    /redirect \(302\)/,
  );
});
