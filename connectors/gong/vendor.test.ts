import assert from "node:assert/strict";
import { test } from "node:test";
import { ConnectError } from "../../packages/connector-kit/src/types.ts";
import { gongVendor, normalizeBaseUrl } from "./vendor.ts";

// The Gong sign-in and calls against a fake Gong: Basic auth on the company's own API address
// only, and Retry-After waits.

const FIELDS = {
  base_url: "https://US-12345.api.gong.io/v2/",
  access_key: "AK",
  access_key_secret: "SECRET",
};

function fakeGong(answers: Response[] = []) {
  const calls: { url: string; init: RequestInit }[] = [];
  const waits: number[] = [];
  const fetchImpl = (async (input: string | URL, init: RequestInit = {}) => {
    calls.push({ url: String(input), init });
    return answers.shift() ?? Response.json({ workspaces: [{ id: "10", name: "Sales" }] });
  }) as typeof fetch;
  const vendor = gongVendor(fetchImpl, async (s) => {
    waits.push(s);
  });
  return { vendor, calls, waits };
}

const header = (init: RequestInit, name: string) => (init.headers as Record<string, string>)[name];

test("a Gong API address is a *.api.gong.io https host and nothing else", () => {
  assert.equal(normalizeBaseUrl("https://us-12345.api.gong.io"), "us-12345.api.gong.io");
  assert.equal(normalizeBaseUrl("us-12345.api.gong.io/"), "us-12345.api.gong.io");
  assert.equal(normalizeBaseUrl("https://api.gong.io/v2"), "api.gong.io");
  for (const bad of [
    "http://us-1.api.gong.io",
    "https://us-1.api.gong.io:8443",
    "https://evil.com",
    "https://api.gong.io.evil.com",
    "https://us-1.api.gong.io/other",
    "https://user@us-1.api.gong.io",
    "https://app.gong.io",
    "",
  ])
    assert.throws(() => normalizeBaseUrl(bad), ConnectError, bad);
});

test("connect proves the key on the company's address and keeps it secret", async () => {
  const { vendor, calls } = fakeGong();
  const c = await vendor.connect(FIELDS);
  assert.deepEqual(c, {
    config: { host: "us-12345.api.gong.io" },
    secret: { access_key: "AK", access_key_secret: "SECRET" },
    accountLabel: "us-12345.api.gong.io",
  });
  assert.equal(calls[0].url, "https://us-12345.api.gong.io/v2/workspaces");
  assert.equal(
    header(calls[0].init, "authorization"),
    `Basic ${Buffer.from("AK:SECRET").toString("base64")}`,
  );
  assert.equal(calls[0].init.redirect, "manual");
});

test("connect explains a refused key, a missing one and other failures", async () => {
  const refused = fakeGong([Response.json({ errors: ["bad key"] }, { status: 401 })]);
  await assert.rejects(refused.vendor.connect(FIELDS), /did not accept the access key/);
  await assert.rejects(fakeGong().vendor.connect({ ...FIELDS, access_key: " " }), ConnectError);
  const down = fakeGong([Response.json({ errors: ["Down for maintenance"] }, { status: 503 })]);
  await assert.rejects(down.vendor.connect(FIELDS), /Gong answered 503: Down for maintenance/);
});

test("calls go to the connection's host with a JSON body", async () => {
  const { vendor, calls } = fakeGong([Response.json({ calls: [] })]);
  const res = await vendor.call(
    {
      config: { host: "us-1.api.gong.io" },
      secret: { access_key: "a", access_key_secret: "b" },
    },
    { method: "POST", path: "/v2/calls/extensive", body: { filter: {} } },
  );
  assert.deepEqual(res, {
    status: 200,
    data: { calls: [] },
    headers: { "content-type": "application/json" },
  });
  assert.equal(calls[0].url, "https://us-1.api.gong.io/v2/calls/extensive");
  assert.equal(calls[0].init.body, '{"filter":{}}');
  assert.equal(header(calls[0].init, "content-type"), "application/json");
  await assert.rejects(
    vendor.call(
      { config: { host: "evil.com" }, secret: {} },
      { method: "GET", path: "/v2/workspaces" },
    ),
  );
});

test("a short Retry-After is waited out; a long one goes back as 429", async () => {
  const limited = (after: string) =>
    new Response("", { status: 429, headers: { "retry-after": after } });
  const conn = {
    config: { host: "api.gong.io" },
    secret: { access_key: "a", access_key_secret: "b" },
  };
  const get = { method: "GET", path: "/v2/workspaces" } as const;

  const short = fakeGong([limited("2"), limited("0")]);
  assert.equal((await short.vendor.call(conn, get)).status, 200);
  assert.deepEqual(short.waits, [2, 1]);
  assert.equal(short.calls.length, 3);

  const long = fakeGong([limited("3600")]);
  const refused = await long.vendor.call(conn, get);
  assert.deepEqual([refused.status, refused.data], [429, { retryAfter: 3600 }]);
  assert.equal(refused.headers?.["retry-after"], "3600");
  assert.deepEqual(long.waits, []);

  const persistent = fakeGong([limited("1"), limited("1"), limited("1")]);
  assert.equal((await persistent.vendor.call(conn, get)).status, 429);
  assert.equal(persistent.calls.length, 3);
});
