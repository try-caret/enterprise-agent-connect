import assert from "node:assert/strict";
import { test } from "node:test";
import { calendarTools, MAX_PEOPLE } from "./tools.ts";
import { loadConnectors } from "../../packages/connector-kit/src/connector.ts";
import type { ToolContext } from "../../packages/connector-kit/src/common.ts";
import type { VendorRequest, VendorResponse } from "../../packages/connector-kit/src/types.ts";

// The calendar tools against a fake Graph, for the agent mailbox agent@contoso.com.

const BOX = "/v1.0/users/agent%40contoso.com";
const JOIN = "https://teams.microsoft.com/l/meetup-join/abc";

function setup(answer: (req: VendorRequest) => VendorResponse | undefined) {
  const requests: VendorRequest[] = [];
  const ctx = {
    limits: { mailbox: "agent@contoso.com" },
    config: {},
    call: async (req: VendorRequest) => {
      requests.push(req);
      return answer(req) ?? { status: 500, data: { error: { message: "unexpected" } } };
    },
    async audit() {},
  } as unknown as ToolContext;
  const tools = new Map<string, (args: unknown) => Promise<any>>();
  const server = {
    registerTool: (name: string, def: any, fn: (args: unknown) => Promise<any>) =>
      tools.set(name, (args: unknown) => fn(def.inputSchema.parse(args))),
  };
  calendarTools.register(server as never, ctx, new Set(calendarTools.names));
  const run = async (name: string, args: object) => {
    const r = await tools.get(name)!(args);
    const text = r.content[0].text as string;
    return { isError: Boolean(r.isError), text, body: r.isError ? null : JSON.parse(text) };
  };
  return { run, requests };
}

const slot = (dateTime: string, timeZone = "UTC") => ({ dateTime, timeZone });

test("the availability tool is its own capability, on the agent mailbox's read role", () => {
  const cal = loadConnectors().find((c) => c.id === "outlook-calendar")!;
  const cap = cal.capabilities.find((c) => c.id === "availability")!;
  assert.deepEqual(cap.tools, ["calendar_get_availability"]);
  assert.deepEqual(cap.scopes, ["Application Calendars.Read"]);
});

test("calendar_get_availability asks as the agent's mailbox in UTC and returns only busy times and working hours", async () => {
  const { run, requests } = setup((req) =>
    req.path.endsWith("/calendar/getSchedule")
      ? {
          status: 200,
          data: {
            value: [
              {
                scheduleId: "pat@contoso.com",
                availabilityView: "0220",
                scheduleItems: [
                  {
                    status: "busy",
                    subject: "Salary review with HR",
                    location: "Room 4",
                    isPrivate: false,
                    start: slot("2026-11-10T15:00:00.0000000"),
                    end: slot("2026-11-10T16:00:00.0000000"),
                  },
                  { status: "free", subject: "x", start: slot("2026-11-10T14:00:00"), end: slot("2026-11-10T15:00:00") },
                  { status: "oof", start: slot("2026-11-11T00:00:00"), end: slot("2026-11-12T00:00:00") },
                ],
                workingHours: {
                  daysOfWeek: ["monday", "tuesday"],
                  startTime: "08:00:00.0000000",
                  endTime: "17:00:00.0000000",
                  timeZone: { name: "Central Standard Time" },
                },
              },
              {
                scheduleId: "lee@example.com",
                error: { message: "Unable to resolve the schedule", responseCode: "ErrorMailRecipientNotFound" },
              },
              {
                scheduleId: "kim@contoso.com",
                scheduleItems: [],
                workingHours: { daysOfWeek: ["monday"], startTime: "09:00:00", endTime: "17:00:00", timeZone: { "@odata.type": "#microsoft.graph.customTimeZone", name: "Customized Time Zone" } },
              },
            ],
          },
        }
      : undefined,
  );
  const r = await run("calendar_get_availability", {
    people: ["pat@contoso.com", "lee@example.com", "kim@contoso.com"],
    start: "2026-11-10T09:00:00-05:00",
    end: "2026-11-12T00:00:00Z",
    interval_minutes: 60,
  });
  assert.equal(r.isError, false, r.text);
  assert.equal(requests[0].method, "POST");
  assert.equal(requests[0].path, `${BOX}/calendar/getSchedule`);
  assert.equal(requests[0].headers?.prefer, 'outlook.timezone="UTC"');
  assert.deepEqual(requests[0].body, {
    schedules: ["pat@contoso.com", "lee@example.com", "kim@contoso.com"],
    startTime: { dateTime: "2026-11-10T14:00:00", timeZone: "UTC" },
    endTime: { dateTime: "2026-11-12T00:00:00", timeZone: "UTC" },
    availabilityViewInterval: 60,
  });
  const [pat, lee, kim] = r.body.people;
  assert.deepEqual(pat, {
    email: "pat@contoso.com",
    error: null,
    busy: [
      { start: "2026-11-10T15:00:00.000Z", end: "2026-11-10T16:00:00.000Z", status: "busy" },
      { start: "2026-11-11T00:00:00.000Z", end: "2026-11-12T00:00:00.000Z", status: "oof" },
    ],
    slots: "0220",
    working_hours: {
      days: ["monday", "tuesday"],
      start: "08:00",
      end: "17:00",
      time_zone: "Central Standard Time",
      time_zone_iana: "America/Chicago",
    },
  });
  assert.ok(!/Salary|Room 4/.test(r.text), "no subject or location reaches the agent");
  assert.deepEqual([lee.error, lee.busy, lee.slots], ["Unable to resolve the schedule", null, null]);
  assert.equal(kim.working_hours.time_zone_iana, null);
});

test("calendar_get_availability refuses a long window and too many people, and fails on a time not in UTC", async () => {
  const { run, requests } = setup((req) =>
    req.path.endsWith("/calendar/getSchedule")
      ? {
          status: 200,
          data: {
            value: [
              {
                scheduleId: "pat@contoso.com",
                scheduleItems: [{ status: "busy", start: slot("2026-11-10T09:00:00", "Pacific Standard Time"), end: slot("2026-11-10T10:00:00", "Pacific Standard Time") }],
              },
            ],
          },
        }
      : undefined,
  );
  const window = { start: "2026-11-01T00:00:00Z", end: "2026-11-02T00:00:00Z" };
  assert.match(
    (await run("calendar_get_availability", { people: ["a@contoso.com"], start: "2026-11-01T00:00:00Z", end: "2027-01-05T00:00:00Z" })).text,
    /at most 62 days/,
  );
  await assert.rejects(() =>
    run("calendar_get_availability", {
      ...window,
      people: Array.from({ length: MAX_PEOPLE + 1 }, (_, i) => `p${i}@contoso.com`),
    }),
  );
  assert.equal(requests.length, 0);
  const r = await run("calendar_get_availability", { ...window, people: ["pat@contoso.com"] });
  assert.equal(r.isError, true);
});

test("calendar_get_availability surfaces a Graph error", async () => {
  const { run } = setup(() => ({ status: 403, data: { error: { message: "Access is denied" } } }));
  const r = await run("calendar_get_availability", {
    people: ["pat@contoso.com"],
    start: "2026-11-01T00:00:00Z",
    end: "2026-11-02T00:00:00Z",
  });
  assert.equal(r.isError, true);
  assert.match(r.text, /403/);
});

const created = (extra: object = {}) => ({
  id: "E1",
  subject: "Intro",
  start: slot("2026-11-10T15:00:00.0000000"),
  end: slot("2026-11-10T15:30:00.0000000"),
  attendees: [
    { emailAddress: { address: "pat@contoso.com" }, type: "required" },
    { emailAddress: { address: "lee@contoso.com" }, type: "optional" },
  ],
  ...extra,
});

test("calendar_create_event makes a Teams meeting with required and optional attendees, an HTML body and a transaction id", async () => {
  const { run, requests } = setup((req) => {
    if (req.method === "POST") return { status: 201, data: { id: "E1" } };
    if (req.path.startsWith(`${BOX}/events/E1?`))
      return {
        status: 200,
        data: created({ isOnlineMeeting: true, onlineMeeting: { joinUrl: JOIN }, transactionId: "req-1" }),
      };
  });
  const r = await run("calendar_create_event", {
    subject: "Intro",
    start: "2026-11-10T09:00",
    end: "2026-11-10T09:30",
    time_zone: "America/Chicago",
    attendees: ["pat@contoso.com", { email: "lee@contoso.com", type: "optional" }],
    body_html: '<p>Agenda</p><script>x</script>',
    online_meeting: true,
    transaction_id: "req-1",
  });
  assert.equal(r.isError, false, r.text);
  const body = requests[0].body as any;
  assert.deepEqual(body.attendees, [
    { emailAddress: { address: "pat@contoso.com" }, type: "required" },
    { emailAddress: { address: "lee@contoso.com" }, type: "optional" },
  ]);
  assert.deepEqual(body.body, { contentType: "HTML", content: "<p>Agenda</p>" });
  assert.equal(body.isOnlineMeeting, true);
  assert.equal(body.onlineMeetingProvider, "teamsForBusiness");
  assert.equal(body.transactionId, "req-1");
  assert.equal(r.body.join_url, JOIN);
  assert.equal(r.body.transaction_id, "req-1");
  assert.deepEqual(
    r.body.attendees.map((a: any) => a.type),
    ["required", "optional"],
  );
  assert.equal(r.body.warning, undefined);
  assert.equal(r.body.existing, undefined);
});

test("calendar_create_event warns, without failing, when Microsoft saved the event with no Teams link", async () => {
  const { run } = setup((req) =>
    req.method === "POST"
      ? { status: 201, data: { id: "E1" } }
      : { status: 200, data: created({ isOnlineMeeting: false, onlineMeeting: null }) },
  );
  const r = await run("calendar_create_event", {
    subject: "Intro",
    start: "2026-11-10T09:00",
    end: "2026-11-10T09:30",
    online_meeting: true,
  });
  assert.equal(r.isError, false, r.text);
  assert.equal(r.body.id, "E1");
  assert.match(r.body.warning, /no Teams meeting link/);
});

test("calendar_create_event answers the first event when Microsoft refuses a repeated transaction id", async () => {
  const { run, requests } = setup((req) => {
    if (req.method === "POST")
      return { status: 409, data: { error: { code: "ErrorDuplicateTransactionId", message: "dup" } } };
    if (req.path.startsWith(`${BOX}/events?`))
      return {
        status: 200,
        data: { value: [{ id: "E0", transactionId: "other" }, { id: "E1", transactionId: "req-1" }] },
      };
    if (req.path.startsWith(`${BOX}/events/E1?`))
      return { status: 200, data: created({ transactionId: "req-1" }) };
  });
  const args = { subject: "Pat's intro", start: "2026-11-10T09:00", end: "2026-11-10T09:30" };
  const r = await run("calendar_create_event", { ...args, transaction_id: "req-1" });
  assert.equal(r.isError, false, r.text);
  assert.equal(r.body.id, "E1");
  assert.equal(r.body.existing, true);
  const q = new URL(`https://graph.microsoft.com${requests[1].path}`).searchParams;
  assert.equal(q.get("$filter"), "subject eq 'Pat''s intro'");

  // Without a transaction id, or with no event carrying it, the refusal is the answer.
  const none = await run("calendar_create_event", { ...args, transaction_id: "req-2" });
  assert.equal(none.isError, true);
  assert.match(none.text, /409: dup/);
  const plain = await run("calendar_create_event", args);
  assert.equal(plain.isError, true);
  assert.match((await run("calendar_create_event", { ...args, transaction_id: "bad id!" })).text, /transaction_id/);
});

test("calendar_update_event patches the same event: attendees, time, Teams and HTML", async () => {
  const { run, requests } = setup((req) => {
    if (req.method === "PATCH") return { status: 200, data: {} };
    if (req.path.includes("$select=id,type"))
      return { status: 200, data: { id: "E1", type: "singleInstance" } };
    return { status: 200, data: created({ onlineMeeting: { joinUrl: JOIN } }) };
  });
  const r = await run("calendar_update_event", {
    event_id: "E1",
    start: "2026-11-11T10:00",
    end: "2026-11-11T10:30",
    time_zone: "UTC",
    attendees: [{ email: "lee@contoso.com", type: "optional" }],
    online_meeting: true,
    body_html: "<b>moved</b>",
  });
  assert.equal(r.isError, false, r.text);
  const patch = requests.find((q) => q.method === "PATCH")!;
  assert.equal(patch.path, `${BOX}/events/E1`);
  assert.deepEqual(patch.body, {
    start: { dateTime: "2026-11-11T10:00", timeZone: "UTC" },
    end: { dateTime: "2026-11-11T10:30", timeZone: "UTC" },
    body: { contentType: "HTML", content: "<b>moved</b>" },
    attendees: [{ emailAddress: { address: "lee@contoso.com" }, type: "optional" }],
    isOnlineMeeting: true,
    onlineMeetingProvider: "teamsForBusiness",
  });
  assert.equal(r.body.id, "E1");
  assert.match(
    (await run("calendar_update_event", { event_id: "E1", online_meeting: false })).text,
    /keeps a Teams meeting/,
  );
});
