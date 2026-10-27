import assert from "node:assert/strict";
import { createHmac } from "node:crypto";
import { test } from "node:test";
import { CredentialKey } from "../src/secrets.ts";
import { InvalidInput } from "../src/service.ts";
import {
  claudeManagedAgents,
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

function receiver(status = 200) {
  const got: { url: string; body: string; sig: string }[] = [];
  const fetch = (async (url: string, init: RequestInit) => {
    got.push({
      url,
      body: String(init.body),
      sig: (init.headers as Record<string, string>)[SIGNATURE_HEADER],
    });
    return new Response(null, { status });
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
  };
  const session = await hook.start(input);
  await hook.continue(session, { ...input, text: "again" });
  const [first, second] = r.got.slice(1).map((g) => JSON.parse(g.body));
  assert.deepEqual(first, {
    type: "message",
    run: "run-1",
    agent: AGENT,
    bot: "bot-1",
    requestedBy: "person",
    text: "hello",
    replyToken: "eac_reply_x",
    mcpUrl: `https://svc.example/agent-connect/mcp/${AGENT}`,
  });
  assert.equal(second.text, "again");
  assert.ok(r.got.every((g) => g.url === "https://hooks.example/teams"));

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
