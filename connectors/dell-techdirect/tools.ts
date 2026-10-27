import { z } from "zod";
import { ENTITLEMENTS_PATH } from "./vendor.js";
import { vendorError } from "../../packages/connector-kit/src/oauth.js";
import {
  audited,
  Denied,
  json,
  VendorFailed,
  type Toolset,
} from "../../packages/connector-kit/src/common.js";

const SERVICE_TAG = /^[A-Za-z0-9]{5,7}$/;

type Asset = {
  serviceTag?: string;
  invalid?: boolean;
  productLineDescription?: string;
  systemDescription?: string;
  shipDate?: string;
  countryCode?: string;
  entitlements?: {
    serviceLevelDescription?: string;
    entitlementType?: string;
    startDate?: string;
    endDate?: string;
  }[];
};

export const dellTools: Toolset = {
  names: ["dell_warranty_lookup"],

  register(server, ctx, allowed) {
    if (allowed.has("dell_warranty_lookup"))
      server.registerTool(
        "dell_warranty_lookup",
        {
          description:
            "Look up Dell warranty coverage for up to 100 service tags: the model, ship date and each support entitlement with its start and end dates.",
          inputSchema: z.object({ service_tags: z.array(z.string()).min(1).max(100) }),
          annotations: { readOnlyHint: true, openWorldHint: true },
        },
        ({ service_tags }) =>
          audited(ctx, "dell_warranty_lookup", service_tags.join(","), async (mark) => {
            const bad = service_tags.find((t) => !SERVICE_TAG.test(t));
            if (bad) throw new Denied(`${bad} is not a Dell service tag`);
            const tags = [...new Set(service_tags.map((t) => t.toUpperCase()))];
            const res = await ctx.call({
              method: "GET",
              path: `${ENTITLEMENTS_PATH}?servicetags=${tags.join(",")}`,
            });
            mark(res.status);
            if (res.status < 200 || res.status >= 300) {
              const detail = vendorError(res.data);
              throw new VendorFailed(
                res.status,
                `Dell answered ${res.status}${detail ? `: ${detail}` : ""}`,
              );
            }
            if (!Array.isArray(res.data))
              throw new VendorFailed(res.status, "Dell answered without a list of assets");
            const now = Date.now();
            return json(
              (res.data as Asset[]).map((a) => ({
                serviceTag: a.serviceTag,
                found: !a.invalid,
                product: a.productLineDescription ?? null,
                system: a.systemDescription ?? null,
                shipDate: a.shipDate ?? null,
                countryCode: a.countryCode ?? null,
                entitlements: (a.entitlements ?? []).map((e) => ({
                  service: e.serviceLevelDescription ?? null,
                  type: e.entitlementType ?? null,
                  start: e.startDate ?? null,
                  end: e.endDate ?? null,
                  active:
                    Date.parse(e.startDate ?? "") <= now && now <= Date.parse(e.endDate ?? ""),
                })),
              })),
            );
          }),
      );
  },
};
