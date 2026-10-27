import { createHash } from "node:crypto";
import { brand } from "../../../connector-kit/src/branding.js";
import { clientAssertion } from "../../../connector-kit/src/certificate.js";
import {
  ConnectError,
  TokenCache,
  type Fetch,
  type VendorResponse,
} from "../../../connector-kit/src/types.js";

// An agent's own bot signing in to its customer's tenant with the certificate the service made for
// it: Bot Connector tokens to post, Graph tokens to read where it is installed (resource-
// specific consent, never a tenant-wide permission). docs/teams.md 2.3, 3, 4.

const LOGIN = "https://login.microsoftonline.com";
const GRAPH = "https://graph.microsoft.com";
export const BOT_CONNECTOR_SCOPE = "https://api.botframework.com/.default";
const GRAPH_SCOPE = `${GRAPH}/.default`;
// Teams' Bot Connector addresses. A service address comes from a token-checked activity; this
// keeps the bot's token from ever being sent anywhere else.
const SERVICE_HOST =
  /^https:\/\/smba\.(trafficmanager\.net|infra\.(gcc|gov)\.teams\.microsoft\.(com|us))\//;
// Microsoft asks bots to retry these with backoff [16].
const RETRY = new Set([412, 429, 502, 504]);
const BACKOFF_MS = [500, 1500, 4000];

export type BotCredential = {
  botAppId: string;
  tenantId: string;
  certificate: string;
  privateKey: string;
};

export function validServiceUrl(url: unknown): url is string {
  return typeof url === "string" && SERVICE_HOST.test(url);
}

const join = (serviceUrl: string, path: string) => `${serviceUrl.replace(/\/+$/, "")}${path}`;

async function parse(res: Response): Promise<VendorResponse> {
  const text = await res.text();
  let data: unknown = null;
  if (text)
    try {
      data = JSON.parse(text);
    } catch {
      data = text;
    }
  return { status: res.status, data };
}

export function botClient(
  fetchImpl: Fetch = fetch,
  sleep = (ms: number) => new Promise((r) => setTimeout(r, ms)),
) {
  const tokens = new TokenCache();

  async function mint(bot: BotCredential, scope: string) {
    const tokenUrl = `${LOGIN}/${bot.tenantId}/oauth2/v2.0/token`;
    const res = await fetchImpl(tokenUrl, {
      method: "POST",
      redirect: "error",
      headers: { "content-type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type: "client_credentials",
        client_id: bot.botAppId,
        scope,
        client_assertion_type: "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
        client_assertion: clientAssertion({
          clientId: bot.botAppId,
          tokenUrl,
          certificate: bot.certificate,
          privateKey: bot.privateKey,
        }),
      }),
    });
    const body = (await res.json().catch(() => ({}))) as {
      access_token?: string;
      expires_in?: number;
      error?: string;
      error_description?: string;
    };
    if (!res.ok || !body.access_token) {
      const why = (body.error_description ?? body.error ?? String(res.status)).split(/\r?\n/)[0];
      throw new ConnectError(`Microsoft refused the bot's sign-in: ${why.slice(0, 300)}`);
    }
    return { token: body.access_token, expiresIn: body.expires_in ?? 3600 };
  }

  // Per bot, certificate and scope, so a switched certificate never reuses an old token.
  const key = (bot: BotCredential, scope: string) =>
    createHash("sha256")
      .update([bot.botAppId, bot.tenantId, bot.certificate, scope].join("\n"))
      .digest("hex");

  async function authorized(
    bot: BotCredential,
    scope: string,
    send: (token: string) => Promise<Response>,
  ) {
    const k = key(bot, scope);
    let res = await send(await tokens.get(k, () => mint(bot, scope)));
    if (res.status === 401) {
      tokens.forget(k);
      res = await send(await tokens.get(k, () => mint(bot, scope)));
    }
    return res;
  }

  return {
    /** Signs in once, without caching: proves a bot's app id and certificate, and that the
     *  app has no organization-wide Microsoft Graph permissions (a bot reads only where it was
     *  added, by resource-specific consent). */
    async prove(bot: BotCredential): Promise<void> {
      await mint(bot, BOT_CONNECTOR_SCOPE);
      const { token } = await mint(bot, GRAPH_SCOPE);
      let roles: unknown;
      try {
        roles = JSON.parse(Buffer.from(token.split(".")[1], "base64url").toString("utf8")).roles;
      } catch {
        throw new ConnectError(`Microsoft returned a token ${brand().productName} could not read`);
      }
      if (Array.isArray(roles) && roles.length)
        throw new ConnectError(
          `This app has organization-wide Microsoft Graph permissions (${roles.slice(0, 5).join(", ")}). A bot needs none: remove them from the app registration, or set up a new bot.`,
        );
    },

    /** A Bot Connector call at a conversation's service address, retried on Microsoft's
     *  throttling answers. */
    async connector(
      bot: BotCredential,
      serviceUrl: string,
      request: { method: "GET" | "POST" | "PUT"; path: string; body?: unknown },
    ): Promise<VendorResponse> {
      if (!validServiceUrl(serviceUrl))
        throw new Error(`Not a Teams service address: ${serviceUrl}`);
      for (let attempt = 0; ; attempt++) {
        const res = await authorized(bot, BOT_CONNECTOR_SCOPE, (token) =>
          fetchImpl(join(serviceUrl, request.path), {
            method: request.method,
            redirect: "error",
            headers: {
              authorization: `Bearer ${token}`,
              accept: "application/json",
              ...(request.body === undefined ? {} : { "content-type": "application/json" }),
            },
            body: request.body === undefined ? undefined : JSON.stringify(request.body),
          }),
        );
        if (!RETRY.has(res.status) || attempt >= BACKOFF_MS.length) return parse(res);
        await res.body?.cancel();
        await sleep(BACKOFF_MS[attempt] * (0.5 + Math.random()));
      }
    },

    /** A Graph read with the bot's own token. */
    async graph(bot: BotCredential, path: string): Promise<VendorResponse> {
      const res = await authorized(bot, GRAPH_SCOPE, (token) =>
        fetchImpl(`${GRAPH}${path}`, {
          redirect: "error",
          headers: { authorization: `Bearer ${token}`, accept: "application/json" },
        }),
      );
      return parse(res);
    },
  };
}

export type BotClient = ReturnType<typeof botClient>;
