import { z } from "zod";
import {
  audited,
  Denied,
  json,
  type ToolContext,
  type Toolset,
} from "../../packages/connector-kit/src/common.js";
import { mailboxPath, nextPage, ok, pageParam } from "../outlook-mail/tools.js";

// The calendar of the mailbox in the agent's grant. No tool takes a mailbox argument, and
// event ids are only ever looked up under that mailbox, where Exchange finds no other
// mailbox's events.

export const MAX_RANGE_DAYS = 62;
const EVENT_ID = /^[A-Za-z0-9_=-]{1,1024}$/;
// A moment with its offset, so the window means the same thing to the service and Microsoft.
const INSTANT = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(:\d{2}(\.\d{1,7})?)?(Z|[+-]\d{2}:\d{2})$/;
// A wall-clock time read in `time_zone`, as Microsoft's events take it.
const LOCAL = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(:\d{2})?$/;
const TIME_ZONE = /^[A-Za-z0-9_/+ .()-]{1,100}$/;
const DATE = /^\d{4}-\d{2}-\d{2}$/;
const FIELDS =
  "id,subject,start,end,location,organizer,attendees,isAllDay,isCancelled,showAs,bodyPreview," +
  "type,seriesMasterId,recurrence,isOrganizer,responseStatus";

export const MAX_OCCURRENCES = 999;
const DAYS = [
  "sunday",
  "monday",
  "tuesday",
  "wednesday",
  "thursday",
  "friday",
  "saturday",
] as const;
type Day = (typeof DAYS)[number];
// The tool's names for Microsoft's response actions.
const RESPONSES = {
  accept: "accept",
  tentatively_accept: "tentativelyAccept",
  decline: "decline",
} as const;

export type RecurrenceInput = {
  pattern: "daily" | "weekly" | "monthly" | "yearly";
  interval: number;
  days_of_week?: Day[];
  end_date?: string;
  count?: number;
};

type Address = { emailAddress?: { name?: string; address?: string } };
type When = { dateTime?: string; timeZone?: string };
type Event = {
  id: string;
  subject?: string;
  start?: When;
  end?: When;
  location?: { displayName?: string };
  organizer?: Address;
  attendees?: (Address & { status?: { response?: string } })[];
  isAllDay?: boolean;
  isCancelled?: boolean;
  showAs?: string;
  bodyPreview?: string;
  body?: { content?: string };
  type?: string;
  seriesMasterId?: string;
  recurrence?: unknown;
  isOrganizer?: boolean;
  responseStatus?: { response?: string };
};

function eventPath(ctx: ToolContext, id: string): string {
  if (!EVENT_ID.test(id)) throw new Denied("Give an event id from calendar_list_events");
  return `${mailboxPath(ctx)}/events/${encodeURIComponent(id)}`;
}

// Date.parse rolls a day that does not exist (February 31, 24:00) into the next one; reading it
// back shows whether it did.
const exists = (value: string, iso: string) => {
  const ms = Date.parse(iso);
  return Number.isFinite(ms) && new Date(ms).toISOString().startsWith(value);
};
const isDate = (value: string) => DATE.test(value) && exists(value, `${value}T00:00:00Z`);

/** The listing window as UTC instants; refused when unbounded, reversed or too long. */
export function eventWindow(start: string, end: string): { start: string; end: string } {
  if (!INSTANT.test(start) || !INSTANT.test(end))
    throw new Denied("Give start and end as date-times with an offset, e.g. 2026-10-21T09:00:00Z");
  const from = Date.parse(start);
  const to = Date.parse(end);
  const real = (v: string) => isDate(v.slice(0, 10)) && Number(v.slice(11, 13)) < 24;
  if (!Number.isFinite(from) || !Number.isFinite(to) || !real(start) || !real(end))
    throw new Denied("start or end is not a date");
  if (to <= from) throw new Denied("end must be after start");
  if (to - from > MAX_RANGE_DAYS * 86_400_000)
    throw new Denied(`List at most ${MAX_RANGE_DAYS} days at a time`);
  return { start: new Date(from).toISOString(), end: new Date(to).toISOString() };
}

function when(value: string, timeZone: string): When {
  if (!LOCAL.test(value) || !exists(value, `${value}Z`))
    throw new Denied(`${value} is not a date-time such as 2026-10-21T09:00`);
  if (!TIME_ZONE.test(timeZone)) throw new Denied(`${timeZone} is not a time zone name`);
  return { dateTime: value, timeZone };
}

function checkOrder(start: When, end: When) {
  if (Date.parse(`${end.dateTime}Z`) <= Date.parse(`${start.dateTime}Z`))
    throw new Denied("end must be after start");
}

/**
 * Microsoft's patternedRecurrence for a series whose first event starts on `startDate`
 * (YYYY-MM-DD, read in `timeZone`). Monthly and yearly repeat on the start's day (and month);
 * weekly defaults to the start's weekday. Every series ends, by date or by count.
 * https://learn.microsoft.com/graph/outlook-schedule-recurring-events
 */
export function recurrence(r: RecurrenceInput, startDate: string, timeZone: string) {
  if (!isDate(startDate)) throw new Denied(`${startDate} is not a date`);
  if (!TIME_ZONE.test(timeZone)) throw new Denied(`${timeZone} is not a time zone name`);
  if ((r.end_date === undefined) === (r.count === undefined))
    throw new Denied("Give the series either an end_date or a count, not both");
  if (r.days_of_week && r.pattern !== "weekly")
    throw new Denied("days_of_week is only for a weekly pattern");
  const start = new Date(`${startDate}T00:00:00Z`);
  const pattern =
    r.pattern === "daily"
      ? { type: "daily", interval: r.interval }
      : r.pattern === "weekly"
        ? {
            type: "weekly",
            interval: r.interval,
            daysOfWeek: [...new Set(r.days_of_week ?? [DAYS[start.getUTCDay()]])],
          }
        : r.pattern === "monthly"
          ? { type: "absoluteMonthly", interval: r.interval, dayOfMonth: start.getUTCDate() }
          : {
              type: "absoluteYearly",
              interval: r.interval,
              dayOfMonth: start.getUTCDate(),
              month: start.getUTCMonth() + 1,
            };
  if (r.end_date !== undefined) {
    if (!isDate(r.end_date)) throw new Denied(`${r.end_date} is not a date such as 2026-12-31`);
    if (r.end_date < startDate) throw new Denied("end_date must not be before the first event");
  }
  const range =
    r.end_date !== undefined
      ? { type: "endDate", startDate, endDate: r.end_date, recurrenceTimeZone: timeZone }
      : {
          type: "numbered",
          startDate,
          numberOfOccurrences: r.count,
          recurrenceTimeZone: timeZone,
        };
  return { pattern, range };
}

// The event a change applies to. An occurrence's id changes that occurrence only; with
// `series`, the change goes to its series master, so every occurrence follows. A master's own
// id needs `series`, so a whole series is never changed or deleted by accident.
async function target(
  ctx: ToolContext,
  mark: (s: number) => void,
  id: string,
  series: boolean,
): Promise<{ id: string; recurrence?: any }> {
  const e: Event = await ok(ctx, mark, {
    method: "GET",
    path: `${eventPath(ctx, id)}?$select=id,type,seriesMasterId,recurrence`,
  });
  if (e.type === "seriesMaster") {
    if (!series)
      throw new Denied(
        "This id is a whole series. Pass series: true to change every occurrence, or give one occurrence's id from calendar_list_events",
      );
    return { id, recurrence: e.recurrence };
  }
  if (!series) return { id };
  if (!e.seriesMasterId) throw new Denied("This event is not part of a series");
  const master: Event = await ok(ctx, mark, {
    method: "GET",
    path: `${eventPath(ctx, e.seriesMasterId)}?$select=id,recurrence`,
  });
  return { id: e.seriesMasterId, recurrence: master.recurrence };
}

// Microsoft answers a create or update in the zone it was sent in, whatever the Prefer header
// says; reading the event back gives UTC, as the list and get tools do.
async function inUtc(ctx: ToolContext, mark: (s: number) => void, id: string): Promise<Event> {
  return ok(ctx, mark, {
    method: "GET",
    path: `${eventPath(ctx, id)}?$select=${FIELDS}`,
    headers: { prefer: 'outlook.timezone="UTC"' },
  });
}

const address = (a?: Address) => a?.emailAddress?.address ?? null;
const attendees = (list: string[]) =>
  list.map((address) => ({ emailAddress: { address }, type: "required" }));

function summary(e: Event) {
  return {
    id: e.id,
    subject: e.subject ?? "",
    start: e.start ?? null,
    end: e.end ?? null,
    all_day: e.isAllDay ?? false,
    cancelled: e.isCancelled ?? false,
    show_as: e.showAs ?? null,
    location: e.location?.displayName ?? "",
    organizer: address(e.organizer),
    attendees: (e.attendees ?? []).map((a) => ({
      email: address(a),
      response: a.status?.response ?? null,
    })),
    preview: e.bodyPreview ?? "",
    // singleInstance, occurrence, exception (a changed occurrence) or seriesMaster.
    type: e.type ?? null,
    series_id: e.seriesMasterId ?? null,
    recurrence: e.recurrence ?? null,
    is_organizer: e.isOrganizer ?? null,
    my_response: e.responseStatus?.response ?? null,
  };
}

export const calendarTools: Toolset = {
  names: [
    "calendar_list_events",
    "calendar_get_event",
    "calendar_create_event",
    "calendar_update_event",
    "calendar_delete_event",
    "calendar_respond_to_event",
  ],

  register(server, ctx, allowed) {
    const readOnly = { readOnlyHint: true, openWorldHint: true } as const;
    const writes = { readOnlyHint: false, destructiveHint: false, openWorldHint: true } as const;
    const eventId = z.string().max(1100).describe("Id from calendar_list_events.");
    const emails = z.array(z.email()).max(50);
    const timeZone = z
      .string()
      .max(100)
      .describe("Time zone of start and end, e.g. UTC or Europe/London.");
    // Text and wall-clock times are checked inside audited(), so a refusal is logged.
    const local = z.string().max(40);
    const repeat = z
      .object({
        pattern: z.enum(["daily", "weekly", "monthly", "yearly"]),
        interval: z
          .number()
          .int()
          .min(1)
          .max(99)
          .default(1)
          .describe("Every this many days, weeks, months or years."),
        days_of_week: z
          .array(z.enum(DAYS))
          .min(1)
          .max(7)
          .optional()
          .describe("Weekly only. Defaults to the start's weekday."),
        end_date: z
          .string()
          .max(10)
          .optional()
          .describe("Last date the series may fall on, e.g. 2026-12-31. Give this or count."),
        count: z.number().int().min(1).max(MAX_OCCURRENCES).optional(),
      })
      .describe(
        "Repeat the event. Monthly and yearly repeat on the start's date. Needs end_date or count.",
      );
    const series = z
      .boolean()
      .default(false)
      .describe(
        "true to apply this to the whole series the event belongs to; false for this one occurrence only.",
      );

    if (allowed.has("calendar_list_events"))
      server.registerTool(
        "calendar_list_events",
        {
          description: `List events in this agent's calendar between two date-times (at most ${MAX_RANGE_DAYS} days apart), earliest first. Each occurrence of a repeating event is listed on its own. Times come back in UTC. When \`more\` is true, call again with \`cursor\` set to \`next\`.`,
          inputSchema: z.object({
            start: z.string().max(40).describe("e.g. 2026-10-21T00:00:00Z"),
            end: z.string().max(40).describe("e.g. 2026-10-28T00:00:00Z"),
            max_results: z.number().int().min(1).max(100).default(50),
            cursor: z
              .string()
              .max(4000)
              .optional()
              .describe(
                "The `next` value from the previous call with the same window: the following page.",
              ),
          }),
          annotations: readOnly,
        },
        ({ start, end, max_results, cursor }) =>
          audited(ctx, "calendar_list_events", `${start} → ${end}`, async (mark) => {
            const window = eventWindow(start, end);
            const query = [
              `startDateTime=${encodeURIComponent(window.start)}`,
              `endDateTime=${encodeURIComponent(window.end)}`,
              `$top=${max_results}`,
              `$select=${FIELDS}`,
              "$orderby=start/dateTime",
              ...pageParam(cursor),
            ].join("&");
            const data = await ok(ctx, mark, {
              method: "GET",
              path: `${mailboxPath(ctx)}/calendarView?${query}`,
              headers: { prefer: 'outlook.timezone="UTC"' },
            });
            return json({
              events: (data.value as Event[]).map(summary),
              more: Boolean(data["@odata.nextLink"]),
              next: nextPage(data),
            });
          }),
      );

    if (allowed.has("calendar_get_event"))
      server.registerTool(
        "calendar_get_event",
        {
          description: "Read one event in this agent's calendar, with its description as text.",
          inputSchema: z.object({ event_id: eventId }),
          annotations: readOnly,
        },
        ({ event_id }) =>
          audited(ctx, "calendar_get_event", event_id, async (mark) => {
            const e: Event = await ok(ctx, mark, {
              method: "GET",
              path: `${eventPath(ctx, event_id)}?$select=${FIELDS},body`,
              headers: {
                prefer: 'outlook.timezone="UTC", outlook.body-content-type="text"',
              },
            });
            return json({ ...summary(e), body: e.body?.content ?? "" });
          }),
      );

    if (allowed.has("calendar_create_event"))
      server.registerTool(
        "calendar_create_event",
        {
          description:
            "Add an event to this agent's calendar. Attendees are sent an invitation from this agent's mailbox.",
          inputSchema: z.object({
            subject: z.string().min(1).max(255),
            start: local.describe("Wall-clock start, e.g. 2026-10-21T09:00"),
            end: local.describe("Wall-clock end, e.g. 2026-10-21T09:30"),
            time_zone: timeZone.default("UTC"),
            location: z.string().max(255).optional(),
            body: z.string().max(100000).optional(),
            attendees: emails.default([]),
            recurrence: repeat.optional(),
          }),
          annotations: writes,
        },
        ({ subject, start, end, time_zone, location, body, attendees: invited, recurrence: r }) =>
          audited(ctx, "calendar_create_event", subject, async (mark, retarget) => {
            const from = when(start, time_zone);
            const to = when(end, time_zone);
            checkOrder(from, to);
            const repeats = r ? { recurrence: recurrence(r, start.slice(0, 10), time_zone) } : {};
            const created: Event = await ok(ctx, mark, {
              method: "POST",
              path: `${mailboxPath(ctx)}/events`,
              body: {
                subject,
                start: from,
                end: to,
                ...(location ? { location: { displayName: location } } : {}),
                ...(body ? { body: { contentType: "Text", content: body } } : {}),
                attendees: attendees(invited),
                ...repeats,
              },
            });
            retarget(created.id);
            return json(summary(await inUtc(ctx, mark, created.id)));
          }),
      );

    if (allowed.has("calendar_update_event"))
      server.registerTool(
        "calendar_update_event",
        {
          description:
            "Change an event in this agent's calendar. Only the fields given change; attendees, when given, replace the list and are sent an update. An occurrence's id changes that occurrence only; pass series: true to change the whole series.",
          inputSchema: z.object({
            event_id: eventId,
            subject: z.string().min(1).max(255).optional(),
            start: local.optional(),
            end: local.optional(),
            time_zone: timeZone
              .optional()
              .describe(
                "Time zone of start and end, e.g. Europe/London. Needed with start or end.",
              ),
            location: z.string().max(255).optional(),
            body: z.string().max(100000).optional(),
            attendees: emails.optional(),
            series,
            recurrence: repeat
              .optional()
              .describe("A new repeat pattern for the series. Needs series: true."),
          }),
          annotations: writes,
        },
        ({
          event_id,
          subject,
          start,
          end,
          time_zone,
          location,
          body,
          attendees: invited,
          series: whole,
          recurrence: r,
        }) =>
          audited(ctx, "calendar_update_event", event_id, async (mark, retarget) => {
            if (r && !whole) throw new Denied("Give series: true to change how a series repeats");
            // No default: a time read in the wrong zone would move the event silently.
            if ((start !== undefined || end !== undefined) === (time_zone === undefined))
              throw new Denied(
                "Give time_zone together with start or end, e.g. Europe/London, and not on its own",
              );
            const from = start === undefined ? undefined : when(start, time_zone!);
            const to = end === undefined ? undefined : when(end, time_zone!);
            if (from && to) checkOrder(from, to);
            const patch: Record<string, unknown> = {
              ...(subject !== undefined ? { subject } : {}),
              ...(from ? { start: from } : {}),
              ...(to ? { end: to } : {}),
              ...(location !== undefined ? { location: { displayName: location } } : {}),
              ...(body !== undefined ? { body: { contentType: "Text", content: body } } : {}),
              ...(invited !== undefined ? { attendees: attendees(invited) } : {}),
            };
            if (Object.keys(patch).length === 0 && !r)
              throw new Denied("Give at least one field to change");
            const t = await target(ctx, mark, event_id, whole);
            retarget(t.id);
            if (r) {
              // The series' first date and zone stay unless a new start (in time_zone) moves them.
              const range = t.recurrence?.range ?? {};
              const startDate = start?.slice(0, 10) ?? range.startDate;
              const zone = start !== undefined ? time_zone : range.recurrenceTimeZone;
              if (!startDate || !zone) throw new Denied("Give start and time_zone with recurrence");
              patch.recurrence = recurrence(r, startDate, zone);
            }
            await ok(ctx, mark, { method: "PATCH", path: eventPath(ctx, t.id), body: patch });
            return json(summary(await inUtc(ctx, mark, t.id)));
          }),
      );

    if (allowed.has("calendar_delete_event"))
      server.registerTool(
        "calendar_delete_event",
        {
          description:
            "Delete an event from this agent's calendar. An occurrence's id deletes that occurrence only; pass series: true to delete every occurrence. Attendees are sent a cancellation when this agent organized it.",
          inputSchema: z.object({ event_id: eventId, series }),
          annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
        },
        ({ event_id, series: whole }) =>
          audited(ctx, "calendar_delete_event", event_id, async (mark, retarget) => {
            const t = await target(ctx, mark, event_id, whole);
            retarget(t.id);
            await ok(ctx, mark, { method: "DELETE", path: eventPath(ctx, t.id) });
            return json({ deleted: true, series: whole });
          }),
      );

    if (allowed.has("calendar_respond_to_event"))
      server.registerTool(
        "calendar_respond_to_event",
        {
          description:
            "Accept, tentatively accept or decline an invitation in this agent's calendar. An occurrence's id answers that occurrence only; a series id answers the whole series. A declined event leaves the calendar.",
          inputSchema: z.object({
            event_id: eventId,
            response: z.enum(["accept", "tentatively_accept", "decline"]),
            comment: z.string().max(10000).optional().describe("Text sent with the answer."),
            send_response: z
              .boolean()
              .default(true)
              .describe("false to answer without emailing the organizer."),
          }),
          annotations: writes,
        },
        ({ event_id, response, comment, send_response }) =>
          audited(ctx, "calendar_respond_to_event", event_id, async (mark) => {
            await ok(ctx, mark, {
              method: "POST",
              path: `${eventPath(ctx, event_id)}/${RESPONSES[response]}`,
              body: { ...(comment ? { comment } : {}), sendResponse: send_response },
            });
            return json({ response, sent: send_response });
          }),
      );
  },
};
