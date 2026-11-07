# Starting agents from app events

An agent can be started by something that happens in a connected app: a Jira issue created, a
comment added. The organization subscribes to the app once; each agent then says which of its
events start it, under which conditions, with which instructions, and how often at most.

```
app ──webhook──▶ POST /agent-connect/events/<subscription id> ─┐
Microsoft Graph ─notification─▶ (same address), item read ─────┤
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
   webhooks, or **Microsoft Graph** change notifications (Outlook mail, below). All produce the
   same events.
2. **Delivery**: the proof is checked on the raw body before anything in it is read (bodies over
   1 MiB are refused). The service then reads the event type, the delivery id and, where the app
   stamps deliveries, the time (refused outside the app's replay window). A delivery id already
   received is ignored (a retry or a replay); deliveries are kept 30 days, longer than any
   window. Only the fields the app's connector keeps are stored, sealed, until processed. The
   app is answered at once; the work happens from the queue.
3. **Decision**, for every enabled trigger on that app in the organization, in this order:
   - the event type is one the trigger chose;
   - the change was not made by the app connection's own account (an agent's comment would
     otherwise start the agent again). When the app will not say which account that is (Jira
     without the `read:jira-user` scope), an event that names who made it starts nothing, and
     the audit says why;
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
each run is a signed POST of `{type: "event", run, agent, app, event, deliveryId, thread, data,
text, mcpUrl}` instead (`X-Agent-Connect-Signature` as for Teams). `GET …/event-triggers` lists an
agent's triggers, `DELETE …/event-triggers/jira` removes one (stopping its live runs), and `GET
…/event-runs` shows what they started.

## Set up Outlook mail

An agent with its own mailbox can be started by mail arriving in it: an assistant copied on a
thread, answering in the thread. Microsoft tells the service of each new message in each
agent's mailbox, the service reads the message with the Outlook mail connection, and every
interval it also reads what changed in the inbox since its last look, so a notification Microsoft
dropped, or one sent while the service was down, still starts the agent, and only once.

1. Connect Outlook mail (its connect steps) and give the agent a grant on it with its mailbox.
   Nothing more is needed in Microsoft 365: the events read mail with the same Exchange role
   (`Application Mail.Read`) the read capability already has.
2. Let Microsoft reach the service: `https://<host>/agent-connect/events/` must be public over
   HTTPS (the chart's Ingress serves it). Microsoft checks the address when each mailbox is
   subscribed, so a deployment that cannot be reached from the internet cannot use these events.
3. Subscribe the organization:

   ```http
   POST /agent-connect/admin/v1/orgs/{org}/event-subscriptions
   {"app": "outlook-mail", "source": "graph", "by": "ann@example.com"}
   ```

   There is no secret to paste anywhere: the service gives Microsoft its own check value and
   compares it on every notification. `intervalSeconds` (5 minutes by default) is how often it
   catches up and renews.
4. Give the agent its trigger, as for Jira, on `message_received`:

   ```http
   PUT /agent-connect/admin/v1/orgs/{org}/agents/{agent}/event-triggers/outlook-mail
   {
     "enabled": true, "eventTypes": ["message_received"],
     "conditions": [{"field": "only_mailbox", "in": ["false"]}],
     "prompt": "You were copied on this email. Read the thread and answer in it if you can help.",
     "runner": "webhook", "settings": {"url": "https://…"}, "signingSecret": "…",
     "budgetCents": 50, "by": "ann@example.com"
   }
   ```

   Its fields: `mailbox`, `message_id`, `conversation_id`, `internet_message_id`, `from`, `to`,
   `cc`, `subject`, `received`, `preview` (the start of the text) and `only_mailbox` (true when
   every To and Cc address is the mailbox itself).

`GET …/event-subscriptions/{id}/watches` shows each watched mailbox: whether Microsoft holds a
live subscription for it, when that expires, the last notification and the last catch-up, and
the last error. A mailbox is watched while an active agent holding it has an enabled trigger:
turning a trigger on or off applies at once, a grant's mailbox changing or an agent being turned
off within an interval. Mail the mailbox sent itself starts nothing, nor does mail received
before its watch started or more than 7 days ago. Removing the subscription deletes Microsoft's
subscriptions first.

How it works with Microsoft: one Graph subscription per mailbox on
`users/{mailbox}/mailFolders('inbox')/messages` for created messages, lasting 10070 minutes
(Microsoft allows under 7 days for mail) and renewed once half of that has passed. Each
notification carries only the message's id; the service reads it. Lifecycle notifications come
to the same address: `reauthorizationRequired` renews, `subscriptionRemoved` subscribes again
and catches up, `missed` catches up. The catch-up is a delta query on the inbox from a stored
delta link, so each message is read once whichever way it arrives. Not yet run against a real
tenant: whether Exchange's application roles cover Graph subscriptions and delta queries the way
they cover reading mail is the first thing to check.

## Threads

An app may name the field that holds an event's conversation (`events.thread`: Outlook mail's
`conversation_id`, Jira's issue `key`). Every run started by an event of one conversation, on one
trigger, then carries the same `thread`: a 32-character id made from the trigger and that value,
which says nothing of the value itself. `run` is still one per event. A host keying its session
on the signed webhook's `thread`, falling back to `run`, continues one conversation across
emails; events of one thread can arrive close together, so a host should take them one at a
time. Claude starts a new session for each event (with `thread` in its metadata): a busy session
cannot take the next event, and the agent reads the thread itself with its own tools.

## Set up Slack

Slack starts an agent when someone @mentions the app in a channel (`app_mention`); the agent
answers in the thread with its own Slack tools. Connect Slack first (its connect steps), give the
agent a grant on Slack with the `post` capability and, if it should hear only some channels, its
`channels` limit. The app hears only channels it is a member of.

Slack makes the secret it signs with (the app's **signing secret**), so the admin gives it to the
subscription instead of receiving one:

1. In Slack: api.slack.com/apps > the app > Basic Information > App Credentials > Signing
   Secret > Show, and copy it.
2. Subscribe, with that secret:

   ```http
   POST /agent-connect/admin/v1/orgs/{org}/event-subscriptions
   {"app": "slack", "source": "webhook", "secret": "<the signing secret>", "by": "ann@example.com"}
   ```

   The answer has the address (`subscription.url`) and the setup steps with it filled in;
   `secret` is null (yours is not shown back). Slack's events cannot be polled.
3. In Slack: the app > Event Subscriptions > Enable Events, Request URL: the address. Slack
   sends a signed `url_verification` at once, which the service answers with its challenge, and
   shows Verified. Subscribe to bot events > `app_mention`, Save Changes, and reinstall the app
   when Slack asks (it adds the `app_mentions:read` permission). The steps also give the same as
   lines to add to the app's manifest.
4. The agent's trigger: `PUT …/agents/{agent}/event-triggers/slack` with `"eventTypes":
   ["app_mention"]` and instructions such as "Answer the question you were asked in {app}:
   reply in the thread (channel, thread_ts) with slack_post_message."

When Slack's signing secret is regenerated, give the new one:
`POST …/event-subscriptions/{id}/secret {"secret": "<new>", "by": "…"}`.

Slack signs each delivery with `X-Slack-Signature: v0=<HMAC-SHA256 of
"v0:<X-Slack-Request-Timestamp>:<body>">`, refused when the timestamp is over 5 minutes off. A
delivery is named by its `event_id`, which Slack repeats on its retries (`X-Slack-Retry-Num`),
so a retry starts nothing. The agent gets `channel`, `user`, `text`, `ts`, `thread_ts` (the
thread to reply in: the mention's thread, or the mention itself when it is not in one) and
`team`; conditions can name any of them. The app's own messages (`user` is the bot user that
`auth.test` names) start nothing.

## How an app adds events

Events are data in the app's `connector.json` (schema: `eventsSchema` in
`packages/connector-kit/src/connector.ts`), so a JSON-only `http` or `sql` app, or a built-in one,
adds them the same way. Paths are `a.b`, or `a[].b` for every item of a list.

```jsonc
"events": {
  "types": [{"id": "ticket.created", "label": "Ticket created"}],
  // What agents, conditions and limits see of an event; values are text, numbers, booleans or
  // lists of them, never objects. A list of paths takes the first that has a value.
  "fields": {"id": "ticket.id", "queue": "ticket.queue", "tags": "ticket.tags[]",
             "thread": ["ticket.thread_id", "ticket.id"]},
  // A limit of the connector → the field holding the event's value for it.
  "limits": {"queues": "queue"},
  // Optional: who the connection is, so its own changes start nothing.
  "self": {"request": {"method": "GET", "path": "/me"}, "id": "id", "fields": ["changed_by"]},
  "webhook": {
    // hmac-sha256 (hex of the body, after an optional prefix), token (the secret itself in a
    // header), or none (only with "allowInsecure": true). An HMAC of more than the body names
    // it in "signs": {body} and {header:Name} in text, e.g. Slack's
    // "v0:{header:X-Slack-Request-Timestamp}:{body}".
    "verify": {"type": "hmac-sha256", "header": "X-Signature", "prefix": "sha256="},
    // Optional: the app makes the secret; the admin gives it as "secret" (Slack).
    "secretFromApp": false,
    "type": {"field": "event"},             // or {"header": "X-Event"}
    "deliveryId": {"header": "X-Delivery"}, // optional: the body's SHA-256 otherwise
    "timestamp": {"at": {"field": "sent_at"}, "unit": "iso", "toleranceSeconds": 300},
    // Optional: the app's signed check of the address, answered with the text at "echo" and
    // starting nothing (Slack: {"at": {"field": "type"}, "equals": "url_verification",
    // "echo": "challenge"}).
    "handshake": {"at": {"field": "kind"}, "equals": "ping", "echo": "nonce"},
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
  },
  // Optional: the field holding the event's conversation (see Threads).
  "thread": "id"
}
```

A Microsoft app can use Graph change notifications instead, as Outlook mail does
(`connectors/outlook-mail/connector.json`): `graph.per` names the limit whose value each Graph
subscription watches (it must also be an event limit, so an agent hears only its own), `resource`
and `changeType` are the subscription's, `read` fetches a notified item (`{value}`, `{id}`),
`delta` reads what changed since the watch started (`{since}`), `wrap`, `deliveryId` and
`received` place the item, `own` names fields that make an item the watched account's own, and
`only` adds true/false fields for "every value of these fields is the watched value".

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
  compared in constant time) before its body is parsed, and every Graph notification its
  clientState. A subscription id alone is not a credential. Graph's validation request is
  answered with its own token as plain text and starts nothing.
- Secrets are made by the service (32 random bytes), shown once, sealed with
  `CONNECT_ENCRYPTION_KEY` and re-sealed during a key rotation, like other credentials. An app
  that makes its own (Slack's signing secret) has it given by the admin, sealed the same way and
  never shown back.
- A handshake (Slack's url_verification) is answered only after the same proof, and only with
  the one value the connector names, as text.
- Replays: a delivery id is accepted once, and deliveries are kept longer than any app's replay
  window, so a captured signed delivery cannot start anything again.
- An event starts only agents granted on that app, within their limits, so a trigger cannot
  widen what an agent sees. The agent then acts with its own grant, limits and audit.
- Event content is written by people in the app: it is quoted as untrusted data, never put into
  the instructions, and cut to the fields the connector keeps (text to 2000 characters, lists to
  50 items).
- Runs are bounded by each trigger's caps and budget. Counting them is one process's job, so the
  service runs as one replica, as for Teams.
