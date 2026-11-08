import assert from "node:assert/strict";
import { test } from "node:test";
import { microsoftTools } from "./tools.ts";
import type { ToolContext } from "../../packages/connector-kit/src/common.ts";
import type { VendorRequest, VendorResponse } from "../../packages/connector-kit/src/types.ts";

// The mail tools against a fake Graph, for the agent mailbox agent@contoso.com.

const BOX = "/v1.0/users/agent%40contoso.com";
const who = (address: string) => ({ emailAddress: { address } });

function setup(answer: (req: VendorRequest) => VendorResponse = () => ({ status: 202, data: null })) {
  const requests: VendorRequest[] = [];
  const ctx = {
    limits: { mailbox: "agent@contoso.com" },
    config: {},
    call: async (req: VendorRequest) => {
      requests.push(req);
      return answer(req);
    },
    async audit() {},
  } as unknown as ToolContext;
  const tools = new Map<string, (args: unknown) => Promise<any>>();
  const server = {
    registerTool: (name: string, def: any, fn: (args: unknown) => Promise<any>) =>
      tools.set(name, (args: unknown) => fn(def.inputSchema.parse(args))),
  };
  microsoftTools.register(server as never, ctx, new Set(microsoftTools.names));
  const run = async (name: string, args: object) => {
    const r = await tools.get(name)!(args);
    const text = r.content[0].text as string;
    return { isError: Boolean(r.isError), text, body: r.isError ? null : JSON.parse(text) };
  };
  return { run, requests };
}

test("mail_reply answers all with reply_all, in the same conversation, and escapes plain text", async () => {
  const { run, requests } = setup();
  let r = await run("mail_reply", { message_id: "M1", body: "Tue <3pm> works\nThanks" });
  assert.equal(r.isError, false, r.text);
  assert.equal(requests[0].path, `${BOX}/messages/M1/reply`);
  assert.deepEqual(requests[0].body, { comment: "Tue &lt;3pm&gt; works<br>Thanks" });

  r = await run("mail_reply", {
    message_id: "M1",
    reply_all: true,
    body_html: '<p>Tue?</p><script>x</script><a href="javascript:x">y</a>',
  });
  assert.equal(r.body.reply_all, true);
  assert.equal(requests[1].path, `${BOX}/messages/M1/replyAll`);
  assert.deepEqual(requests[1].body, { comment: "<p>Tue?</p><a>y</a>" });
});

test("mail_reply and mail_send take exactly one of body and body_html", async () => {
  const { run, requests } = setup();
  for (const [tool, extra] of [
    ["mail_reply", { message_id: "M1" }],
    ["mail_send", { to: ["a@example.com"], subject: "s" }],
  ] as const) {
    assert.match((await run(tool, { ...extra })).text, /Give body or body_html/);
    assert.match(
      (await run(tool, { ...extra, body: "x", body_html: "<p>x</p>" })).text,
      /not both/,
    );
    assert.match((await run(tool, { ...extra, body_html: "<script>x</script>" })).text, /nothing left/);
  }
  assert.equal(requests.length, 0);
});

test("mail_send sends sanitized HTML", async () => {
  const { run, requests } = setup();
  const r = await run("mail_send", {
    to: ["a@example.com"],
    subject: "Times",
    body_html: '<table><tr><td style="color: red; position: absolute">Mon</td></tr></table><iframe></iframe>',
  });
  assert.equal(r.isError, false, r.text);
  assert.deepEqual((requests[0].body as any).message.body, {
    contentType: "HTML",
    content: '<table><tr><td style="color: red">Mon</td></tr></table>',
  });
});

test("mail_get_message gives what a thread needs: ids, recipients, new text and whether it is to this mailbox alone", async () => {
  const { run, requests } = setup(() => ({
    status: 200,
    data: {
      id: "M1",
      subject: "Re: meeting",
      from: who("pat@example.com"),
      toRecipients: [who("Agent@Contoso.com")],
      ccRecipients: [],
      receivedDateTime: "2026-11-05T10:00:00Z",
      conversationId: "CONV1",
      internetMessageId: "<abc@example.com>",
      body: { content: "Tuesday works.\n\nOn Monday Pat wrote: ..." },
      uniqueBody: { content: "Tuesday works." },
    },
  }));
  const r = await run("mail_get_message", { message_id: "M1" });
  assert.match(requests[0].path, /uniqueBody/);
  assert.match(requests[0].path, /internetMessageId/);
  assert.equal(requests[0].headers?.prefer, 'outlook.body-content-type="text"');
  assert.equal(r.body.conversation, "CONV1");
  assert.equal(r.body.internet_message_id, "<abc@example.com>");
  assert.deepEqual([r.body.to, r.body.cc], [["Agent@Contoso.com"], []]);
  assert.equal(r.body.only_to_me, true);
  assert.equal(r.body.new_text, "Tuesday works.");
});

test("mail_list_messages lists a conversation and says when the mailbox is only copied", async () => {
  const { run, requests } = setup(() => ({
    status: 200,
    data: {
      value: [
        {
          id: "M1",
          toRecipients: [who("pat@example.com")],
          ccRecipients: [who("agent@contoso.com")],
          conversationId: "CONV1",
          isRead: false,
        },
      ],
    },
  }));
  const r = await run("mail_list_messages", { conversation: "CONV1", unread_only: true });
  assert.equal(r.isError, false, r.text);
  const q = new URL(`https://graph.microsoft.com${requests[0].path}`).searchParams;
  assert.equal(q.get("$filter"), "conversationId eq 'CONV1' and isRead eq false");
  assert.equal(r.body.messages[0].only_to_me, false);
  assert.deepEqual(r.body.messages[0].cc, ["agent@contoso.com"]);

  assert.match((await run("mail_list_messages", { conversation: "x' or 1 eq 1" })).text, /conversation value/);
  assert.match((await run("mail_list_messages", { conversation: "C", search: "hi" })).text, /not both/);
  assert.equal(requests.length, 1);
});

test("a draft addressed to nobody yet is not only to this mailbox", async () => {
  const { run } = setup(() => ({
    status: 200,
    data: { value: [{ id: "D1", toRecipients: [], ccRecipients: [], isDraft: true }] },
  }));
  const r = await run("mail_list_messages", { folder: "drafts" });
  assert.equal(r.isError, false, r.text);
  assert.equal(r.body.messages[0].only_to_me, false);
});
