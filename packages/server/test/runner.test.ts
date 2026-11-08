import assert from "node:assert/strict";
import { createHmac } from "node:crypto";
import { test } from "node:test";
import { CredentialKey } from "../src/secrets.ts";
import { InvalidInput } from "../src/service.ts";
import {
  claudeManagedAgents,
  DELIVERY_HEADER,
  nextAttempt,
  retryAfter,
  RetryLater,
  signature,
  signedWebhook,
  SIGNATURE_HEADER,
  type RunInput,
  type RunnerContext,
} from "../src/teams/runner.ts";
import type { RunRow, TriggerRow } from "../src/teams/store.ts";
import { ConnectError } from "../../connector-kit/src/types.ts";

const ORG = "11111111-1111-1111-1111-111111111111";
const AGENT = "22222222-2222-2222-2222-222222222222";
const SECRET = "s".repeat(40);
const key = new CredentialKey(Buffer.alloc(32, 7).toString("base64"));
const noKey = { claudeKey: async () => null } as unknown as RunnerContext["teams"];
const ctx = { teams: noKey, key, publicUrl: "https://svc.example/agent-connect" };

const trigger = (over: Partial<TriggerRow> = {}): TriggerRow => ({
  agentId: AGENT,
  enabled: true,
  runner: "webhook",
  settings: { url: "https://hooks.example/teams" },
  secret: null,
  secretKeyId: null,
  managedAgentId: null,
  environmentId: null,
  vaultIds: [],
  budgetCents: 50,
  maxRunsPerHour: 20,
  maxCentsPerDay: 500,
  maxConcurrent: 2,
  replyMode: "mentions",
  allowedTeams: [],
  allowedPeople: [],
  anyoneInOrganization: true,
  updatedBy: "ann",
  updatedAt: "",
  ...over,
});

function receiver(status = 200, headers: Record<string, string> = {}) {
  const got: { url: string; body: string; sig: string; delivery?: string }[] = [];
  const fetch = (async (url: string, init: RequestInit) => {
    const sent = init.headers as Record<string, string>;
    got.push({
      url,
      body: String(init.body),
      sig: sent[SIGNATURE_HEADER],
      delivery: sent[DELIVERY_HEADER],
    });
    return new Response(null, { status, headers });
  }) as never;
  return { got, fetch };
}

// What a receiver does: recompute the HMAC over "<t>.<body>" with the shared secret.
function verifies(secret: string, body: string, header: string) {
  const { t, v1 } = Object.fromEntries(header.split(",").map((p) => p.split("=")));
  return createHmac("sha256", secret).update(`${t}.${body}`).digest("hex") === v1;
}

test("the signed webhook pings when saved and posts each message signed", async () => {
  const r = receiver();
  const hook = signedWebhook(r.fetch)(ctx);
  assert.ok(hook.settings.safeParse({ url: "https://hooks.example/teams" }).success);
  assert.ok(!hook.settings.safeParse({ url: "http://hooks.example/teams" }).success);
  assert.equal(await hook.unready(ORG), null);

  await hook.verify(trigger(), ORG, SECRET, "save");
  await hook.verify(trigger(), ORG, SECRET, "answer");
  assert.equal(r.got.length, 1);
  assert.deepEqual(JSON.parse(r.got[0].body), { type: "ping", agent: AGENT });
  assert.ok(verifies(SECRET, r.got[0].body, r.got[0].sig));
  assert.ok(!verifies("other".repeat(10), r.got[0].body, r.got[0].sig));

  const input: RunInput = {
    org: ORG,
    agentId: AGENT,
    runId: "run-1",
    botId: "bot-1",
    requestedBy: "person",
    text: "hello",
    replyToken: "eac_reply_x",
    budgetCents: 50,
    trigger: trigger(),
    secret: SECRET,
    deliveryId: "d-1",
  };
  const session = await hook.start(input);
  await hook.continue(session, { ...input, text: "again" });
  const [first, second] = r.got.slice(1).map((g) => JSON.parse(g.body));
  assert.deepEqual(first, {
    type: "message",
    run: "run-1",
    delivery: "d-1",
    agent: AGENT,
    bot: "bot-1",
    requestedBy: "person",
    text: "hello",
    replyToken: "eac_reply_x",
    mcpUrl: `https://svc.example/agent-connect/mcp/${AGENT}`,
  });
  assert.equal(second.text, "again");
  assert.ok(r.got.every((g) => g.url === "https://hooks.example/teams"));
  // The delivery id rides in a header and in the signed body; a ping has none.
  assert.deepEqual(
    r.got.map((g) => g.delivery),
    [undefined, "d-1", "d-1"],
  );

  // Answered once a reply came back through teams_reply.
  const run = { sessionId: session, replyCount: 0 } as RunRow;
  assert.equal(((await hook.state(run, ORG)) as { status: string }).status, "running");
  assert.equal(
    ((await hook.state({ ...run, replyCount: 1 }, ORG)) as { status: string }).status,
    "idle",
  );
});

test("the signed webhook refuses without a secret and on a refusal", async () => {
  const hook = signedWebhook(receiver(500).fetch)(ctx);
  await assert.rejects(hook.verify(trigger(), ORG, null, "save"), InvalidInput);
  await assert.rejects(hook.verify(trigger(), ORG, SECRET, "save"), (e: unknown) => {
    return e instanceof ConnectError && /500/.test(e.message);
  });
});

test("signatures carry their time", () => {
  assert.match(signature(SECRET, "{}", 1700000000), /^t=1700000000,v1=[0-9a-f]{64}$/);
});

test("Claude Managed Agents needs the organization's Claude key and its agent", async () => {
  const claude = claudeManagedAgents(() => {
    throw new Error("no key, so no client");
  })(ctx);
  assert.equal(await claude.unready(ORG), "organization has no Claude API key");
  await assert.rejects(
    claude.verify(trigger({ runner: claude.id, settings: {} }), ORG, null, "save"),
    InvalidInput,
  );
  await assert.rejects(
    claude.verify(
      trigger({ runner: claude.id, managedAgentId: "a", environmentId: "e" }),
      ORG,
      null,
      "save",
    ),
    /Claude API key/,
  );
  assert.deepEqual(await claude.state({ sessionId: "s" } as RunRow, ORG), {
    gone: "Claude API key removed",
  });
  // Nothing to stop a session with once the key is gone.
  await claude.stop("s", ORG);
});

const input = (deliveryId = "d-1"): RunInput => ({
  org: ORG,
  agentId: AGENT,
  runId: "run-1",
  botId: "bot-1",
  requestedBy: "person",
  text: "hello",
  replyToken: "eac_reply_x",
  budgetCents: 50,
  trigger: trigger(),
  secret: SECRET,
  deliveryId,
});

test("the signed webhook asks for a retry on no answer, 5xx, 408, 409 and 429", async () => {
  for (const status of [500, 502, 503, 504, 408, 409, 429]) {
    const hook = signedWebhook(receiver(status).fetch)(ctx);
    await assert.rejects(hook.start(input()), (e: unknown) => {
      return e instanceof RetryLater && e.message === `The webhook answered ${status}`;
    });
  }
  const limited = signedWebhook(receiver(429, { "retry-after": "30" }).fetch)(ctx);
  await assert.rejects(limited.start(input()), (e: unknown) => {
    return e instanceof RetryLater && e.retryAfterMs === 30_000;
  });
  const down = signedWebhook((async () => {
    throw new TypeError("fetch failed");
  }) as never)(ctx);
  await assert.rejects(down.start(input()), (e: unknown) => {
    return e instanceof RetryLater && e.message === "The webhook could not be reached";
  });
  const slow = signedWebhook((async () => {
    throw new DOMException("The operation was aborted due to timeout", "TimeoutError");
  }) as never)(ctx);
  await assert.rejects(slow.continue("s", input()), (e: unknown) => {
    return e instanceof RetryLater && /did not answer within 10 seconds/.test(e.message);
  });
});

test("the signed webhook fails at once on any other refusal", async () => {
  for (const status of [400, 401, 403, 404, 410, 422]) {
    const hook = signedWebhook(receiver(status).fetch)(ctx);
    await assert.rejects(hook.start(input()), (e: unknown) => {
      return (
        e instanceof ConnectError &&
        !(e instanceof RetryLater) &&
        e.message === `The webhook answered ${status}`
      );
    });
  }
});

test("the signed webhook never follows a redirect, and fails at once on one", async () => {
  for (const status of [301, 302, 307, 308]) {
    const modes: (string | undefined)[] = [];
    const hook = signedWebhook((async (_url: string, init: RequestInit) => {
      modes.push(init.redirect);
      return new Response(null, { status, headers: { location: "https://elsewhere.example/" } });
    }) as never)(ctx);
    await assert.rejects(hook.start(input()), (e: unknown) => {
      return (
        e instanceof ConnectError &&
        !(e instanceof RetryLater) &&
        e.message === `The webhook answered ${status}`
      );
    });
    assert.deepEqual(modes, ["manual"]);
  }
});

test("every attempt of one delivery carries its id, signed", async () => {
  const r = receiver(503);
  const hook = signedWebhook(r.fetch)(ctx);
  for (let i = 0; i < 3; i++) await assert.rejects(hook.start(input("same")), RetryLater);
  assert.deepEqual(
    r.got.map((g) => [g.delivery, JSON.parse(g.body).delivery]),
    [
      ["same", "same"],
      ["same", "same"],
      ["same", "same"],
    ],
  );
  assert.ok(r.got.every((g) => verifies(SECRET, g.body, g.sig)));
});

test("retries back off 5 s, 15 s, 30 s, 1 min, 2 min, then every 4 min, for 15 minutes", () => {
  const WINDOW = 15 * 60_000;
  const mid = () => 0.5;
  const delays: number[] = [];
  let now = 0;
  for (let attempts = 1; ; attempts++) {
    const at = nextAttempt(attempts, 0, WINDOW, null, now, mid);
    if (at === null) break;
    delays.push(at - now);
    now = at;
  }
  assert.deepEqual(delays, [5_000, 15_000, 30_000, 60_000, 120_000, 240_000, 240_000]);
  assert.ok(now <= WINDOW);
  // Within 20% either way, never past the 4 minute step.
  assert.equal(nextAttempt(1, 0, WINDOW, null, 0, () => 0), 4_000);
  assert.equal(nextAttempt(1, 0, WINDOW, null, 0, () => 1), 6_000);
  assert.equal(nextAttempt(9, 0, WINDOW, null, 0, () => 1), 240_000);
  // Retry-After is waited out, up to the 4 minute step.
  assert.equal(nextAttempt(1, 0, WINDOW, 90_000, 0, mid), 90_000);
  assert.equal(nextAttempt(1, 0, WINDOW, 600_000, 0, mid), 240_000);
  assert.equal(nextAttempt(5, 0, WINDOW, 1_000, 0, mid), 120_000);
  // Nothing past the window; a window of 0 never tries again.
  assert.equal(nextAttempt(1, 0, WINDOW, null, WINDOW - 1_000, mid), null);
  assert.equal(nextAttempt(1, 0, 0, null, 0, mid), null);
});

test("Retry-After is read in seconds or as a date", () => {
  const now = Date.parse("2026-01-01T00:00:00Z");
  assert.equal(retryAfter("120", now), 120_000);
  assert.equal(retryAfter("Thu, 01 Jan 2026 00:01:00 GMT", now), 60_000);
  assert.equal(retryAfter("Wed, 31 Dec 2025 23:00:00 GMT", now), 0);
  assert.equal(retryAfter("soon", now), null);
  assert.equal(retryAfter(null, now), null);
});
