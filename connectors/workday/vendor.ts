import { createHash } from "node:crypto";
import { vendorFetch } from "../../packages/connector-kit/src/fetch.js";
import { vendorError } from "../../packages/connector-kit/src/oauth.js";
import {
  ConnectError,
  TokenCache,
  type Fetch,
  type Vendor,
  type VendorConnection,
  type VendorRequest,
} from "../../packages/connector-kit/src/types.js";

// A Workday API client for integrations, signing in as its Integration System User with a
// non-expiring refresh token (OAuth 2.0 refresh token grant at /ccx/oauth2/<tenant>/token, the
// client id and secret as HTTP Basic). REST calls under /ccx/api/<service>/<version>/<tenant>,
// custom reports (Reports-as-a-Service) under /ccx/service/customreport2/<tenant>. Service
// paths are Workday's REST directory (developer.workday.com, services_oas3.json, read October
// 27, 2026); its spec files need a Workday sign-in, so answer fields follow CData's Workday
// tables, which are generated from them.

const HOST = /^[a-z0-9-]+(\.[a-z0-9-]+)*\.(myworkday|workday)\.com$/;
const TENANT = /^[A-Za-z0-9_]{1,64}$/;
const REPORT_SEGMENT = /^[A-Za-z0-9_.\-]{1,200}$/;
const PROMPT = /^[A-Za-z0-9_!.\-]{1,100}$/;
const MAX_REPORTS = 30;
// Workday's token answer may carry no lifetime: reuse a token this long, re-mint once on a 401.
const DEFAULT_TOKEN_SECONDS = 600;

export type ReportPrompt = { name: string; type: "date" | "text" };
export type Report = { name: string; owner: string; report: string; prompts: ReportPrompt[] };

// The REST API endpoint shown on Workday's View API Clients page, e.g.
// https://wd5-services1.myworkday.com/ccx/api/v1/acme: host and tenant.
export function parseEndpoint(raw: string): { host: string; tenant: string } {
  let url: URL;
  try {
    url = new URL(raw.trim());
  } catch {
    throw new ConnectError(
      "Enter the Workday REST API endpoint, such as https://wd5-services1.myworkday.com/ccx/api/v1/acme",
    );
  }
  const host = url.hostname.toLowerCase();
  if (url.protocol !== "https:" || url.port || !HOST.test(host))
    throw new ConnectError(
      `${url.host} is not a Workday address (ending myworkday.com or workday.com)`,
    );
  const tenant = /^\/ccx\/api\/(?:[A-Za-z]+\/)?v\d+\/([^/]+)\/?$/.exec(url.pathname)?.[1];
  if (!tenant || !TENANT.test(tenant))
    throw new ConnectError(
      "The endpoint should end with /ccx/api/v1/<tenant>, as shown on the View API Clients page",
    );
  return { host, tenant };
}

// Each report's JSON link from Workday (the report's Actions > Web Service > View URLs > JSON),
// separated by spaces or new lines. Its query names are the report's prompts; one whose example
// value is a date takes dates.
export function parseReports(raw: string, host: string, tenant: string): Report[] {
  const links = raw.split(/\s+/).filter(Boolean);
  if (links.length > MAX_REPORTS) throw new ConnectError(`Enter at most ${MAX_REPORTS} reports`);
  const reports = links.map((link) => {
    let url: URL;
    try {
      url = new URL(link);
    } catch {
      throw new ConnectError(`${link.slice(0, 100)} is not a report link`);
    }
    const parts = url.pathname.split("/").map((p) => decodeURIComponent(p));
    const [, ccx, service, kind, t, owner, report] = parts;
    if (
      url.protocol !== "https:" ||
      url.hostname.toLowerCase() !== host ||
      ccx !== "ccx" ||
      service !== "service" ||
      kind !== "customreport2" ||
      parts.length !== 7 ||
      !REPORT_SEGMENT.test(owner ?? "") ||
      !REPORT_SEGMENT.test(report ?? "")
    )
      throw new ConnectError(
        `${link.slice(0, 100)} is not a custom report link on ${host} (…/ccx/service/customreport2/${tenant}/<owner>/<report>)`,
      );
    if (t !== tenant) throw new ConnectError(`${report} belongs to tenant ${t}, not ${tenant}`);
    const prompts: ReportPrompt[] = [];
    for (const [name, value] of url.searchParams) {
      if (name === "format") continue;
      if (!PROMPT.test(name)) throw new ConnectError(`${report}: ${name} is not a prompt name`);
      prompts.push({ name, type: /^\d{4}-\d{2}-\d{2}/.test(value) ? "date" : "text" });
    }
    return { name: report, owner, report, prompts };
  });
  const names = reports.map((r) => r.name);
  const twice = names.find((n, i) => names.indexOf(n) !== i);
  if (twice) throw new ConnectError(`${twice} is entered twice`);
  return reports;
}

export function connectedReports(config: Record<string, string>): Report[] {
  return config.reports ? (JSON.parse(config.reports) as Report[]) : [];
}

export function workdayVendor(unguarded: Fetch = fetch): Vendor {
  const fetchImpl = vendorFetch("Workday", unguarded);
  const tokens = new TokenCache();
  const key = (c: VendorConnection) =>
    createHash("sha256")
      .update(
        `${c.config.host}\n${c.config.tenant}\n${c.secret.client_id}\n${c.secret.refresh_token}`,
      )
      .digest("hex");

  async function mint(c: VendorConnection) {
    const { host, tenant } = c.config;
    const res = await fetchImpl(`https://${host}/ccx/oauth2/${tenant}/token`, {
      method: "POST",
      headers: {
        authorization: `Basic ${Buffer.from(`${c.secret.client_id}:${c.secret.client_secret}`).toString("base64")}`,
        "content-type": "application/x-www-form-urlencoded",
        accept: "application/json",
      },
      body: new URLSearchParams({
        grant_type: "refresh_token",
        refresh_token: c.secret.refresh_token,
      }),
    });
    const raw = await res.text();
    const body = parse(raw) as { access_token?: unknown; expires_in?: unknown } | string | null;
    if (!res.ok)
      throw new ConnectError(
        `Workday refused the sign-in (${res.status}${vendorError(body) ? `: ${vendorError(body)}` : ""}). Check the client ID, secret and refresh token, and that the refresh token was made for this API client.`,
      );
    if (typeof body !== "object" || !body || typeof body.access_token !== "string")
      throw new ConnectError("Workday answered the sign-in without a token");
    const expiresIn = Number(body.expires_in);
    return {
      token: body.access_token,
      expiresIn: Number.isFinite(expiresIn) && expiresIn > 60 ? expiresIn : DEFAULT_TOKEN_SECONDS,
    };
  }

  async function send(c: VendorConnection, access: string, request: VendorRequest) {
    const res = await fetchImpl(`https://${c.config.host}${request.path}`, {
      method: request.method,
      headers: {
        authorization: `Bearer ${access}`,
        accept: "application/json",
        ...(request.body === undefined ? {} : { "content-type": "application/json" }),
      },
      body: request.body === undefined ? undefined : JSON.stringify(request.body),
    });
    const raw = await res.text();
    return {
      status: res.status,
      data: raw ? parse(raw) : null,
      headers: Object.fromEntries(res.headers),
    };
  }

  const call: Vendor["call"] = async (c, request) => {
    if (!HOST.test(c.config.host ?? "") || !TENANT.test(c.config.tenant ?? ""))
      throw new Error("Workday connection has no host or tenant");
    if (!request.path.startsWith("/ccx/")) throw new Error(`Not a Workday path: ${request.path}`);
    let res = await send(c, await tokens.get(key(c), () => mint(c)), request);
    if (res.status === 401) {
      tokens.forget(key(c));
      res = await send(c, await tokens.get(key(c), () => mint(c)), request);
    }
    return res;
  };

  return {
    async connect(fields) {
      const { host, tenant } = parseEndpoint(fields.endpoint ?? "");
      const clientId = fields.client_id?.trim();
      const clientSecret = fields.client_secret?.trim();
      const refreshToken = fields.refresh_token?.trim();
      if (!clientId || !clientSecret || !refreshToken)
        throw new ConnectError("Enter the client ID, client secret and refresh token");
      const reports = parseReports(fields.reports ?? "", host, tenant);
      const orgColumn = fields.report_org_column?.trim() ?? "";
      if (orgColumn && !PROMPT.test(orgColumn))
        throw new ConnectError(`${orgColumn} is not a report column name`);
      const conn: VendorConnection = {
        config: {
          host,
          tenant,
          account: `${host}/${tenant}`,
          reports: JSON.stringify(reports),
          ...(orgColumn ? { reportOrgColumn: orgColumn } : {}),
        },
        secret: { client_id: clientId, client_secret: clientSecret, refresh_token: refreshToken },
      };
      tokens.forget(key(conn));
      await tokens.get(key(conn), () => mint(conn));
      return { ...conn, accountLabel: `${tenant} (${host})` };
    },

    call,

    checks: {
      // One worker read proves the account reaches staffing data; nothing of it is shown.
      async tenant(connection) {
        const res = await call(connection, {
          method: "GET",
          path: `/ccx/api/staffing/v7/${connection.config.tenant}/workers?limit=1`,
        });
        if (res.status === 401 || res.status === 403)
          throw new ConnectError(
            `Workday refused to list workers (${res.status}). Give the account's security group View on Worker Data: Public Worker Reports, activate the pending security policy changes, then check again.`,
          );
        if (res.status < 200 || res.status >= 300)
          throw new ConnectError(
            `Workday answered ${res.status}: ${vendorError(res.data) || "no reason given"}`,
          );
        return { values: [`${connection.config.tenant} (${connection.config.host})`] };
      },
      async reports(connection) {
        return { values: connectedReports(connection.config).map((r) => r.name) };
      },
    },
  };
}

function parse(raw: string): unknown {
  try {
    return JSON.parse(raw);
  } catch {
    return raw;
  }
}
