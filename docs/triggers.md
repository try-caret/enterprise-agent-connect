# Starting agents from app events

An agent can be started by something that happens in a connected app: a Jira issue created, a
comment added. The organization subscribes to the app once; each agent then says which of its
events start it, under which conditions, with which instructions, and how often at most.

```
app ──webhook──▶ POST /agent-connect/events/<subscription id> ─┐
app ◀──poll─────  every interval, from a cursor ───────────────┤
                                                              ▼
                               queue (deduplicated by delivery id)
                                                              ▼
            each agent's trigger on that app: event type, own changes, access,
                    limits, conditions, runner ready, caps  ─▶  run
```

## How it works

1. **Subscription** (one per organization and app): either a **webhook**, where the app posts
   each event to `https://<host>/agent-connect/events/<subscription id>`, signed with a secret
   made for that subscription, or a **poll**, where the service reads the app every interval
   (60 seconds to a day, 5 minutes by default) from a cursor, for apps that cannot send
   webhooks. Both produce the same events.
2. **Delivery**: the proof is checked on the raw body before anything in it is read (bodies over
   1 MiB are refused). The service then reads the event type, the delivery id and, where the app
   stamps deliveries, the time (refused outside the app's replay window). A delivery id already
   received is ignored (a retry or a replay); deliveries are kept 30 days, longer than any
   window. Only the fields the app's connector keeps are stored, sealed, until processed. The
   app is answered at once; the work happens from the queue.
3. **Decision**, for every enabled trigger on that app in the organization, in this order:
   - the event type is one the trigger chose;
   - the change was not made by the app connection's own account (an agent's comment would
     otherwise start the agent again);
   - the agent is active and holds a grant on the app (with the app connected);
   - the event is within the agent's limits on that app: an agent limited to the Jira project
     KAN is never started by an event in HELP (an event without the field is outside a set
     limit);
   - every condition holds (`{field, in: [...]}` or `{field, notIn: [...]}`; a list field such as
     labels matches when any of its items does);
   - the runner can run now (for Claude, the organization's Claude key is set);
   - the trigger's caps: runs per hour, spending per day (what finished runs cost plus each live
     run's budget), runs at once.
4. **Run**: the same runners as Teams (`claude-managed-agents` or the signed `webhook`, see the
   Trigger runners section of CLAUDE.md), checked again before each run. The agent gets the
   trigger's instructions, then the event's kept fields as untrusted data. It acts with its own
   tools and limits through its own MCP address; nothing is replied to. A run ends when its
   session goes idle, fails at its budget, and is stopped after 30 minutes.

Every delivery (received, refused for its signature, too old, a repeat, an undeclared type) is
audited as `events:delivery`, every decision for every trigger (started, filtered, out of
limits, capped, not started) as `events:run` under the agent, and failed polls as
`events:poll`. `GET /orgs/{org}/audit?app=jira` reads them. Refused signatures are audited at most
once a minute per subscription, since anyone can post to the address.

## Set up Jira

Connect Jira first (its connect steps), give the agent a grant on Jira with the projects it may
work in, and set the organization's Claude key if the agent runs on Claude.

### With a webhook

```http
POST /agent-connect/admin/v1/orgs/{org}/event-subscriptions
{"app": "jira", "source": "webhook", "filter": "project = KAN", "register": true, "by": "ann@example.com"}
```

The answer has the address (`subscription.url`), the **secret, shown only here**
(`POST …/event-subscriptions/{id}/secret` makes a new one), and the setup steps with the address
filled in. With `register: true` the service asks Jira to add the webhook through the
connection. Jira only allows that to a Jira administrator, which the service account is not
meant to be, so `registration.ok` is usually false and a Jira administrator adds it by hand:

- Jira > Settings (gear) > System > WebHooks > Create a WebHook
- URL: the subscription's address. Secret: the secret. Events: Issue created, updated and
  deleted, Comment created. JQL filter: the projects agents should hear about.

Jira signs each delivery with `X-Hub-Signature: sha256=<HMAC-SHA256 of the body>`, names it with
`X-Atlassian-Webhook-Identifier` (the same on every retry) and stamps it with `timestamp`. Jira
retries for over an hour, so its replay window is a day.

### With a poll

```http
POST /agent-connect/admin/v1/orgs/{org}/event-subscriptions
{"app": "jira", "source": "poll", "filter": "project = KAN", "intervalSeconds": 120, "by": "ann@example.com"}
```

Every interval the service searches `(<filter>) AND updated >= -<minutes>m ORDER BY updated ASC`
with the connection's own service account, minutes counted from the cursor (the newest `updated`
seen, starting when the subscription was made). An issue created after the cursor is
`jira:issue_created` (even if it was also changed since), any other `jira:issue_updated`; its id
is `<key>@<updated>`, so the overlap between polls starts nothing twice. A poll cannot see deleted issues or comments, and
cannot tell who made a change, so the own-change guard does not apply: when the agent may also
change issues, give its trigger a condition that its own change no longer meets (for example
`{"field": "status", "in": ["To Do"]}` when the agent moves the issue on) and keep its caps low.

### The agent's trigger

```http
PUT /agent-connect/admin/v1/orgs/{org}/agents/{agent}/event-triggers/jira
{
  "enabled": true,
  "eventTypes": ["jira:issue_created"],
  "conditions": [{"field": "issue_type", "in": ["Bug"]}, {"field": "labels", "notIn": ["no-agent"]}],
  "prompt": "A new {event} arrived in {app}. Triage it: set a priority and comment with next steps.",
  "managedAgentId": "agent_…", "environmentId": "env_…", "vaultIds": ["vlt_…"],
  "budgetCents": 50, "maxRunsPerHour": 20, "maxCentsPerDay": 500, "maxConcurrent": 2,
  "by": "ann@example.com"
}
```

`{app}` and `{event}` are the only placeholders: the app's name and the event's label. The
event's data never goes into the instructions; it follows them as quoted data. The fields a
condition can name are the app's `events.fields` (Jira: key, project, issue_type, status,
priority, labels, summary, assignee, reporter, changed_by, changed, comment, comment_by,
updated). With `"runner": "webhook"`, `"settings": {"url": "https://…"}` and a `signingSecret`,
each run is a signed POST of `{type: "event", run, agent, app, event, deliveryId, data, text,
mcpUrl}` instead (`X-Agent-Connect-Signature` as for Teams). `GET …/event-triggers` lists an
agent's triggers, `DELETE …/event-triggers/jira` removes one (stopping its live runs), and `GET
…/event-runs` shows what they started.

## How an app adds events

Events are data in the app's `connector.json` (schema: `eventsSchema` in
`packages/connector-kit/src/connector.ts`), so a JSON-only `http` or `sql` app, or a built-in one,
adds them the same way. Paths are `a.b`, or `a[].b` for every item of a list.

```jsonc
"events": {
  "types": [{"id": "ticket.created", "label": "Ticket created"}],
  // What agents, conditions and limits see of an event; values are text, numbers, booleans or
  // lists of them, never objects.
  "fields": {"id": "ticket.id", "queue": "ticket.queue", "tags": "ticket.tags[]"},
  // A limit of the connector → the field holding the event's value for it.
  "limits": {"queues": "queue"},
  // Optional: who the connection is, so its own changes start nothing.
  "self": {"request": {"method": "GET", "path": "/me"}, "id": "id", "fields": ["changed_by"]},
  "webhook": {
    // hmac-sha256 (hex of the body, after an optional prefix), token (the secret itself in a
    // header), or none (only with "allowInsecure": true).
    "verify": {"type": "hmac-sha256", "header": "X-Signature", "prefix": "sha256="},
    "type": {"field": "event"},             // or {"header": "X-Event"}
    "deliveryId": {"header": "X-Delivery"}, // optional: the body's SHA-256 otherwise
    "timestamp": {"at": {"field": "sent_at"}, "unit": "iso", "toleranceSeconds": 300},
    "setup": [{"title": "…", "body": "…", "click": ["… URL: {url}"]}],
    // Optional: registering through the app's API with the connection's credentials.
    "register": {"method": "POST", "path": "/hooks", "location": "self",
                 "body": {"url": "{url}", "secret": "{secret}", "events": "{events}"}}
  },
  "poll": {
    "request": {"method": "GET", "path": "/tickets?changed_since_minutes={minutes}{filter}"},
    "filter": "&q={value}",
    "items": "tickets", "cursor": "changed_at", "deliveryId": ["id", "changed_at"],
    "types": [{"type": "ticket.created", "since": "created_at"},
              {"type": "ticket.updated"}]
  }
}
```

A poll's request goes through the app's own client (`vendor.call`), so it reaches only the app's
own hosts with the connection's credentials, as tools do; `{minutes}` is the time since the
cursor rounded up plus one, `{cursor}` the cursor itself, `{filter}` the subscription's filter
put into `filter`. Keep `fields` to what an agent needs: everything kept is shown to it.

A SQL app (`sql` block) polls through one of its own queries: the request's path is `/<tool
name>` and its body the values to bind, for example a tool `orders_changed` with
`select order_no, customer_code, status, changed_at from orders where changed_at > :since order
by changed_at limit 100` and

```json
"poll": {
  "request": {"method": "POST", "path": "/orders_changed",
              "body": {"values": {"since": "{cursor}"}, "answerLimits": []}},
  "items": "rows", "wrap": "order", "cursor": "order.changed_at",
  "deliveryId": ["order.order_no", "order.changed_at"], "types": [{"type": "order.changed"}]
}
```

## Security

- The address is public (add `/agent-connect/events` to what the load balancer exposes) but
  proves nothing by itself: every delivery needs the subscription's secret (HMAC or token,
  compared in constant time) before its body is parsed. A subscription id alone is not a
  credential.
- Secrets are made by the service (32 random bytes), shown once, sealed with
  `CONNECT_ENCRYPTION_KEY` and re-sealed during a key rotation, like other credentials.
- Replays: a delivery id is accepted once, and deliveries are kept longer than any app's replay
  window, so a captured signed delivery cannot start anything again.
- An event starts only agents granted on that app, within their limits, so a trigger cannot
  widen what an agent sees. The agent then acts with its own grant, limits and audit.
- Event content is written by people in the app: it is quoted as untrusted data, never put into
  the instructions, and cut to the fields the connector keeps (text to 2000 characters, lists to
  50 items).
- Runs are bounded by each trigger's caps and budget. Counting them is one process's job, so the
  service runs as one replica, as for Teams.
