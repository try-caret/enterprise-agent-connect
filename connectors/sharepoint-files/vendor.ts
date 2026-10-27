import { brand } from "../../packages/connector-kit/src/branding.js";
import type { MicrosoftClient } from "../microsoft/client.js";
import { ConnectError, type Vendor } from "../../packages/connector-kit/src/types.js";

// SharePoint sites through the Microsoft connection, for SharePoint files and spreadsheets. Under
// Sites.Selected the app reaches only the sites granted to it, one by one; with the "all sites,
// read-only" choice (Sites.Read.All) it reads every site. Either way the admin adds the sites
// agents may use one at a time (config.sites), and each key is kept to its own.
// A OneDrive is a person's site (https://<tenant>-my.sharepoint.com/personal/<name>), reached
// only when it is added and granted like any other.

const SITE_URL =
  /^https:\/\/[a-z0-9-]+\.sharepoint\.com(\/(sites|teams|personal)\/[a-z0-9][a-z0-9._-]*)?$/;
const SITE_KINDS = ["sites", "teams", "personal"];
export const MAX_SITES = 50;

/** A SharePoint site address as admins paste it, reduced to the site's own address. */
export function normalizeSiteUrl(input: string): string {
  const url = URL.parse(input.trim());
  const parts = url?.pathname.toLowerCase().split("/").filter(Boolean) ?? [];
  let name = "";
  try {
    name = decodeURIComponent(parts[1] ?? "");
  } catch {
    // Not a valid escape: the pattern below refuses it.
    name = "";
  }
  // A bare host is the root site; any other path must start with a site kind, so a typo
  // such as /site/Finance is refused rather than widened to the root site.
  const site =
    url &&
    url.protocol === "https:" &&
    !url.port &&
    (!parts.length || SITE_KINDS.includes(parts[0]))
      ? `https://${url.hostname}${parts.length ? `/${parts[0]}/${name}` : ""}`
      : "";
  if (!SITE_URL.test(site))
    throw new ConnectError(
      `${input.trim() || "A site"} is not a SharePoint site address, such as https://your-company.sharepoint.com/sites/Finance`,
    );
  return site;
}

/** The Graph path that reads a site by its address. */
export function sitePath(siteUrl: string): string {
  const { hostname, pathname } = new URL(normalizeSiteUrl(siteUrl));
  return pathname === "/"
    ? `/v1.0/sites/${hostname}`
    : `/v1.0/sites/${hostname}:${pathname.split("/").map(encodeURIComponent).join("/")}`;
}

/** Site addresses, normalized, each once, at most MAX_SITES. */
export function normalizeSites(raw: string[]): string[] {
  const sites = [...new Set(raw.filter((s) => s.trim()).map(normalizeSiteUrl))];
  if (sites.length > MAX_SITES) throw new ConnectError(`Add at most ${MAX_SITES} sites`);
  return sites;
}

export const configSites = (config: Record<string, string>) =>
  (config.sites ?? "").split("\n").filter(Boolean);

type Site = { id?: string; webUrl?: string; displayName?: string };
const writeProbe = () => `${brand().resourcePrefix}-access-check.txt`;

export function sitesVendor(ms: MicrosoftClient): Vendor {
  const { call } = ms;

  // Opens the site by its address, then reads its document library: what agents need.
  async function access(connection: Parameters<typeof call>[0], raw: string, write: boolean) {
    const site = normalizeSiteUrl(raw);
    const refused = (status: number) =>
      status === 401 || status === 403
        ? ({
            status: "not_granted",
            message:
              "Microsoft has not given the app access to this site yet. Run the command below, or send the note to whoever manages SharePoint, then check again.",
          } as const)
        : status === 404 || status === 400
          ? ({
              status: "not_found",
              message: "There is no site at this address. Open it in a browser to check it.",
            } as const)
          : ({
              status: "failed",
              message: `Microsoft answered ${status}. Check again later.`,
            } as const);
    const res = await call(connection, {
      method: "GET",
      path: `${sitePath(site)}?$select=id,webUrl,displayName`,
    });
    const found = res.data as Site | null;
    if (res.status !== 200 || !found?.id) return { site, name: null, ...refused(res.status) };
    const name = found.displayName ?? null;
    const drive = await call(connection, {
      method: "GET",
      path: `/v1.0/sites/${encodeURIComponent(found.id)}/drive/root?$select=id`,
    });
    if (drive.status !== 200) return { site, name, ...refused(drive.status) };
    if (write) {
      // Opening an upload session proves write access. SharePoint may hold the name meanwhile,
      // so the session is cancelled at once; nothing is uploaded.
      const upload = await call(connection, {
        method: "POST",
        path: `/v1.0/sites/${encodeURIComponent(found.id)}/drive/root:/${writeProbe()}:/createUploadSession`,
        body: { item: { "@microsoft.graph.conflictBehavior": "fail" } },
      });
      if (upload.status === 401 || upload.status === 403)
        return {
          site,
          name,
          status: "not_granted" as const,
          message:
            "The app can read this site but not change files there. Run the command below (it gives write access), or send the note to whoever manages SharePoint, then check again.",
        };
      // 409: a file by that name is already there, which still means the app may write.
      if (upload.status !== 200 && upload.status !== 409)
        return {
          site,
          name,
          status: "failed" as const,
          message: `Microsoft answered ${upload.status} when ${brand().productName} checked write access. Check again later.`,
        };
      const uploadUrl = (upload.data as { uploadUrl?: string } | null)?.uploadUrl;
      if (upload.status === 200) {
        if (!uploadUrl) throw new Error("Microsoft opened an upload session without its address");
        const cancelled = await ms.cancelUpload(uploadUrl);
        if (cancelled !== 204)
          console.error(`[agent-connect] write check upload session not cancelled: ${cancelled}`);
      }
    }
    return {
      site,
      name,
      status: "granted" as const,
      message: write
        ? "Agents can open and change this site's files."
        : "Agents can open this site's files.",
    };
  }

  return {
    async connect(fields, provider, previous) {
      if (!provider) throw new Error("Site apps connect through the Microsoft connection");
      // Sites are added after connecting, and kept across a reconnect.
      const sites = configSites(previous ?? {});
      return {
        config: {
          ...(fields.access ? { access: fields.access } : {}),
          ...(sites.length ? { sites: sites.join("\n") } : {}),
        },
        secret: {},
        accountLabel: provider.accountLabel,
      };
    },

    call,

    sites: { normalize: normalizeSites, access },

    checks: {
      permissions: ms.permissions,

      async sites(connection) {
        const values: string[] = [];
        const waiting: string[] = [];
        for (const site of configSites(connection.config)) {
          const res = await call(connection, {
            method: "GET",
            path: `${sitePath(site)}?$select=id`,
          });
          (res.status === 200 ? values : waiting).push(site);
        }
        return {
          values,
          ...(waiting.length
            ? {
                warning: `The app cannot open ${waiting.join(", ")} yet. Give it access, then check again.`,
              }
            : {}),
        };
      },
    },
  };
}
