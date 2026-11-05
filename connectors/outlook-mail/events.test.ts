import assert from "node:assert/strict";
import { test } from "node:test";
import { fill, outsideLimits } from "../../packages/server/src/events/deliver.ts";
import { graphItem } from "../../packages/server/src/events/graph.ts";
import { microsoft as mail } from "../../packages/server/test/helpers.ts";

// Outlook mail's events as its connector.json declares them, on a message shaped like Graph's
// (learn.microsoft.com/graph/api/resources/message, read November 5, 2026). Never run against
// a real tenant.

const events = mail.events!;
const graph = events.graph!;
const MAILBOX = "helper@contoso.com";
const message = {
  id: "AAMkAD=",
  conversationId: "AAQkAD=",
  internetMessageId: "<abc@contoso.com>",
  from: { emailAddress: { name: "Ann", address: "ann@contoso.com" } },
  toRecipients: [{ emailAddress: { name: "Helper", address: MAILBOX } }],
  ccRecipients: [{ emailAddress: { name: "Bob", address: "bob@contoso.com" } }],
  subject: "Quarterly numbers",
  receivedDateTime: "2026-11-05T10:00:00Z",
  bodyPreview: "Can you pull the figures?",
};

test("a received message keeps the fields an agent and its conditions need", () => {
  const { data, own } = graphItem(graph, events.fields, MAILBOX, message);
  assert.deepEqual(data, {
    mailbox: MAILBOX,
    message_id: "AAMkAD=",
    conversation_id: "AAQkAD=",
    internet_message_id: "<abc@contoso.com>",
    from: "ann@contoso.com",
    to: [MAILBOX],
    cc: ["bob@contoso.com"],
    subject: "Quarterly numbers",
    received: "2026-11-05T10:00:00Z",
    preview: "Can you pull the figures?",
    only_mailbox: false,
  });
  assert.equal(own, false);
  assert.equal(events.thread, "conversation_id");
  // Each agent hears only its own mailbox.
  assert.equal(outsideLimits(events.limits, data, { mailbox: MAILBOX }), null);
  assert.ok(outsideLimits(events.limits, data, { mailbox: "other@contoso.com" }));
});

test("mail the mailbox sent itself is its own; mail to it alone says so", () => {
  const self = { ...message, from: { emailAddress: { address: "Helper@Contoso.com" } } };
  assert.equal(graphItem(graph, events.fields, MAILBOX, self).own, true);
  const alone = { ...message, ccRecipients: [] };
  assert.equal(graphItem(graph, events.fields, MAILBOX, alone).data.only_mailbox, true);
});

test("the subscription, read and catch-up address the mailbox's inbox", () => {
  assert.equal(
    fill(graph.resource, { value: MAILBOX }),
    `users/${MAILBOX}/mailFolders('inbox')/messages`,
  );
  assert.equal(graph.changeType, "created");
  // Graph's limit for message subscriptions is 10080 minutes; a little under it.
  assert.ok(graph.maxMinutes < 10080);
  const delta = String(fill(graph.delta.path, { value: "v", since: "2026-11-05T10:00:00.000Z" }));
  assert.match(
    delta,
    /^\/v1\.0\/users\/v\/mailFolders\/inbox\/messages\/delta\?changeType=created&\$select=[^&]*conversationId[^&]*&\$filter=receivedDateTime%20ge%202026-11-05T10:00:00\.000Z$/,
  );
});
