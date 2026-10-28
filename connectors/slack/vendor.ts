import {
  ConnectError,
  type Fetch,
  type Vendor,
  type VendorConnection,
  type VendorRequest,
  type VendorResponse,
} from "../../packages/connector-kit/src/types.js";

// A Slack app's bot token (xoxb-), made from the manifest in the setup steps and installed to
// one workspace. Web API methods at slack.com/api/<method>: reads as GET with query arguments,
// writes as POST JSON (docs.slack.dev/apis/web-api). Slack answers most
// failures 200 with {ok: false, error}; call() turns those into a status so the service's
// probes and the tools' VendorFailed treat them like any app's refusal.

const API = "https://slack.com/api";
// Where files.getUploadURLExternal points the file's bytes; the only other host called.
const UPLOAD_HOST = "files.slack.com";
// Waits longer than this are answered as 429 for the agent to retry later.
const MAX_WAIT_SECONDS = 10;
const RETRIES = 2;

const STATUS: Record<string, number> = {
  not_authed: 401,
  invalid_auth: 401,
  account_inactive: 401,
  token_revoked: 401,
  token_expired: 401,
  missing_scope: 403,
  not_allowed_token_type: 403,
  no_permission: 403,
  restricted_action: 403,
  ekm_access_denied: 403,
  channel_not_found: 404,
  message_not_found: 404,
  thread_not_found: 404,
  users_not_found: 404,
  user_not_found: 404,
  ratelimited: 429,
};

// Slack's error codes in words an agent or admin can act on; others are passed as they are.
const WORDS: Record<string, string> = {
  not_in_channel:
    "the app is not a member of that channel. Someone in the channel can add it with /invite",
  channel_not_found: "no channel with that id is visible to the app",
  is_archived: "that channel is archived",
  message_not_found: "no message with that timestamp in that channel",
  thread_not_found: "no thread with that timestamp in that channel",
  cant_delete_message: "the app can delete only its own messages",
  msg_too_long: "the message is too long",
  no_text: "the message is empty",
  invalid_name: "no emoji with that name",
  too_many_reactions: "that message has too many reactions",
  users_not_found: "no person with that email address",
  missing_scope: "the Slack app lacks a permission this needs",
  invalid_auth: "Slack no longer accepts the app's token",
  token_revoked: "the app's token was revoked",
  account_inactive: "the app was uninstalled or its token deactivated",
};

export function slackError(res: VendorResponse): string {
  const d = res.data as { error?: unknown; needed?: unknown; retry_after?: unknown } | null;
  if (res.status === 429) {
    const after =
      typeof d?.retry_after === "number" ? ` Try again in ${d.retry_after} seconds.` : "";
    return `Slack is limiting how often the app may call it.${after}`;
  }
  const code = typeof d?.error === "string" ? d.error : null;
  if (!code) return `Slack answered ${res.status}`;
  const needed = code === "missing_scope" && typeof d?.needed === "string" ? ` (${d.needed})` : "";
  return `Slack answered ${code}: ${WORDS[code] ?? code.replaceAll("_", " ")}${needed}`;
}

const sleep = (seconds: number) => new Promise((r) => setTimeout(r, seconds * 1000));

export function slackVendor(fetchImpl: Fetch = fetch, wait = sleep): Vendor {
  // A 429 means Slack did not run the call, so even a post is safe to send again.
  async function send(token: string, request: VendorRequest) {
    const upload = request.path.startsWith("https://");
    if (upload && new URL(request.path).host !== UPLOAD_HOST)
      throw new Error("Slack upload address on an unexpected host");
    for (let attempt = 0; ; attempt++) {
      const res = await fetchImpl(upload ? request.path : `${API}${request.path}`, {
        method: request.method,
        redirect: "error",
        headers: upload
          ? { "content-type": "application/octet-stream" }
          : {
              authorization: `Bearer ${token}`,
              ...(request.body === undefined
                ? {}
                : { "content-type": "application/json; charset=utf-8" }),
            },
        body:
          request.body === undefined
            ? undefined
            : upload
              ? String(request.body)
              : JSON.stringify(request.body),
      });
      const raw = await res.text();
      if (res.status === 429) {
        const after = Number(res.headers.get("retry-after") ?? "1");
        if (attempt < RETRIES && Number.isFinite(after) && after <= MAX_WAIT_SECONDS) {
          await wait(Math.max(after, 1));
          continue;
        }
        const data = { ok: false, error: "ratelimited", retry_after: after };
        return { headers: res.headers, response: { status: 429, data } };
      }
      const data = raw ? parse(raw) : null;
      const error = !upload && res.ok ? refusal(data) : null;
      const status = error === null ? res.status : (STATUS[error] ?? 400);
      return { headers: res.headers, response: { status, data } as VendorResponse };
    }
  }

  const token = (c: VendorConnection) => {
    const t = c.secret.bot_token;
    if (!t) throw new Error("Slack connection has no bot token");
    return t;
  };
  const call: Vendor["call"] = async (c, request) => (await send(token(c), request)).response;

  return {
    async connect(fields) {
      const botToken = fields.bot_token?.trim() ?? "";
      if (!botToken.startsWith("xoxb-"))
        throw new ConnectError(
          "Enter the app's Bot User OAuth Token, which starts with xoxb- (Install App in the app's settings)",
        );
      const { response } = await send(botToken, { method: "GET", path: "/auth.test" });
      const d = response.data as Record<string, unknown> | null;
      if (response.status === 401) throw new ConnectError("Slack did not accept the token");
      if (response.status !== 200 || typeof d?.team_id !== "string")
        throw new ConnectError(`Slack sign-in failed. ${slackError(response)}`);
      const url = typeof d.url === "string" ? d.url : "";
      const team = typeof d.team === "string" ? d.team : d.team_id;
      return {
        config: { teamId: d.team_id, team, url, botUserId: String(d.user_id ?? "") },
        secret: { bot_token: botToken },
        accountLabel: url ? `${team} (${new URL(url).host})` : team,
      };
    },

    call,

    checks: {
      // auth.test proves the token; Slack names its scopes in x-oauth-scopes on every answer.
      async workspace(connection, { expected }) {
        const { headers, response } = await send(token(connection), {
          method: "GET",
          path: "/auth.test",
        });
        if (response.status !== 200)
          throw new ConnectError(`Slack sign-in failed. ${slackError(response)}`);
        const d = response.data as { team?: string; url?: string };
        const has = new Set((headers.get("x-oauth-scopes") ?? "").split(",").map((s) => s.trim()));
        const missing = expected.filter((s) => !has.has(s));
        return {
          values: [d.url ? `${d.team} (${new URL(d.url).host})` : String(d.team)],
          ...(missing.length
            ? {
                warning: `The Slack app is missing these bot scopes: ${missing.join(", ")}. Add them under OAuth & Permissions, reinstall the app to the workspace, then check again.`,
              }
            : {}),
        };
      },
      async channels(connection) {
        return { values: (await memberChannels(call, connection)).map((c) => c.id) };
      },
    },
  };
}

export type Channel = {
  id: string;
  name?: string;
  is_private?: boolean;
  is_archived?: boolean;
  topic?: { value?: string };
  purpose?: { value?: string };
};

// Every channel the bot is a member of, for the admin's limit choices: at most 5 pages of 200.
export async function memberChannels(call: Vendor["call"], connection: VendorConnection) {
  const out: Channel[] = [];
  let cursor = "";
  for (let page = 0; page < 5; page++) {
    const res = await call(connection, {
      method: "GET",
      path: `/users.conversations?${channelQuery(200, cursor)}`,
    });
    if (res.status !== 200) throw new ConnectError(slackError(res));
    const d = res.data as { channels?: Channel[]; response_metadata?: { next_cursor?: string } };
    out.push(...(d.channels ?? []));
    cursor = d.response_metadata?.next_cursor ?? "";
    if (!cursor) break;
  }
  return out;
}

export const channelQuery = (limit: number, cursor = "") =>
  new URLSearchParams({
    types: "public_channel,private_channel",
    exclude_archived: "true",
    limit: String(limit),
    ...(cursor ? { cursor } : {}),
  }).toString();

function refusal(data: unknown): string | null {
  if (!data || typeof data !== "object") return null;
  const d = data as { ok?: unknown; error?: unknown };
  return d.ok === false ? (typeof d.error === "string" ? d.error : "unknown_error") : null;
}

function parse(raw: string): unknown {
  try {
    return JSON.parse(raw);
  } catch {
    return raw;
  }
}
