import assert from "node:assert/strict";
import { test } from "node:test";
import { jamfVendor, normalizeUrl } from "./vendor.ts";
import { ConnectError } from "../../packages/connector-kit/src/types.ts";

// The Jamf Pro sign-in and calls against a fake fetch.

function fakeFetch(token: { status?: number; body?: unknown } = {}) {
  const calls: { url: string; init: RequestInit }[] = [];
  const impl = (async (url: string, init: RequestInit) => {
    calls.push({ url, init });
    if (url.endsWith("/api/v1/oauth/token"))
      return new Response(
        JSON.stringify(
          token.body ?? { access_token: "tok", expires_in: 1800, token_type: "Bearer" },
        ),
        { status: token.status ?? 200 },
      );
    if (url.endsWith("/api/v1/sites"))
      return new Response(
        JSON.stringify([
          { id: "1", name: "North" },
          { id: "2", name: "South" },
        ]),
      );
    return new Response("", { status: 201 });
  }) as typeof fetch;
  return { impl, calls };
}

const FIELDS = { url: "https://acme.jamfcloud.com/", client_id: "id-1", client_secret: "s3cret" };

test("normalizeUrl keeps only an https origin, cloud or on premises", () => {
  assert.equal(normalizeUrl(" https://acme.jamfcloud.com/ "), "https://acme.jamfcloud.com");
  assert.equal(normalizeUrl("acme.jamfcloud.com"), "https://acme.jamfcloud.com");
  assert.equal(normalizeUrl("https://jamf.acme.com:8443"), "https://jamf.acme.com:8443");
  for (const bad of [
    "http://acme.jamfcloud.com",
    "https://acme.jamfcloud.com/api",
    "https://acme.jamfcloud.com/?x=1",
    "https://user:pw@acme.jamfcloud.com",
    "ftp://acme.jamfcloud.com",
    "",
  ])
    assert.throws(() => normalizeUrl(bad), ConnectError, bad);
});

test("connect signs in with client credentials at /api/v1/oauth/token and names the host", async () => {
  const { impl, calls } = fakeFetch();
  const conn = await jamfVendor(impl).connect(FIELDS);
  assert.deepEqual(conn, {
    config: { url: "https://acme.jamfcloud.com", host: "acme.jamfcloud.com" },
    secret: { client_id: "id-1", client_secret: "s3cret" },
    accountLabel: "acme.jamfcloud.com",
  });
  assert.equal(calls[0].url, "https://acme.jamfcloud.com/api/v1/oauth/token");
  const form = new URLSearchParams(String(calls[0].init.body));
  assert.equal(form.get("grant_type"), "client_credentials");
  assert.equal(form.get("client_id"), "id-1");
  assert.equal(form.get("client_secret"), "s3cret");
});

test("connect refuses missing credentials and a refused sign-in plainly", async () => {
  const { impl } = fakeFetch({ status: 401, body: { error: "invalid_client" } });
  await assert.rejects(
    jamfVendor(impl).connect({ ...FIELDS, client_secret: " " }),
    /Enter the client ID and secret/,
  );
  await assert.rejects(
    jamfVendor(impl).connect(FIELDS),
    /Jamf Pro refused the sign-in: invalid_client/,
  );
});

test("calls go to the connected server only, JSON as JSON and Classic XML as is", async () => {
  const { impl, calls } = fakeFetch();
  const vendor = jamfVendor(impl);
  const conn = await vendor.connect(FIELDS);
  await vendor.call(conn, { method: "POST", path: "/api/v2/mdm/blank-push", body: { a: 1 } });
  await vendor.call(conn, {
    method: "PUT",
    path: "/JSSResource/computergroups/id/5",
    body: "<computer_group/>",
    headers: { "content-type": "application/xml" },
  });
  const [, json, xml] = calls;
  assert.equal(json.url, "https://acme.jamfcloud.com/api/v2/mdm/blank-push");
  assert.equal(json.init.body, '{"a":1}');
  assert.equal((json.init.headers as Record<string, string>)["content-type"], "application/json");
  assert.equal((json.init.headers as Record<string, string>).authorization, "Bearer tok");
  assert.equal(xml.init.body, "<computer_group/>");
  assert.equal((xml.init.headers as Record<string, string>)["content-type"], "application/xml");
  // A stored address that is not a bare https origin is never called.
  await assert.rejects(
    vendor.call(
      { ...conn, config: { url: "https://evil.example/x" } },
      { method: "GET", path: "/" },
    ),
    /no server address/,
  );
});

test("the sites check lists site ids with their names, and says which privilege is missing", async () => {
  const { impl } = fakeFetch();
  const vendor = jamfVendor(impl);
  const conn = await vendor.connect(FIELDS);
  assert.deepEqual(await vendor.checks!.sites(conn, { expected: [] }), {
    values: ["1", "2"],
    labels: { "1": "North", "2": "South" },
  });
  const refused = jamfVendor((async (url: string) =>
    url.endsWith("/oauth/token")
      ? new Response(JSON.stringify({ access_token: "t", expires_in: 1800 }))
      : new Response("", { status: 403 })) as typeof fetch);
  await assert.rejects(refused.checks!.sites(conn, { expected: [] }), /Add Read Sites/);
});
