import { z } from "zod";
import { vendorError } from "../../packages/connector-kit/src/oauth.js";
import { connectedAccounts } from "./vendor.js";
import {
  audited,
  Denied,
  json,
  listLimit,
  VendorFailed,
  type ToolContext,
  type Toolset,
} from "../../packages/connector-kit/src/common.js";

// FedEx Track, Rate and Ship (developer.fedex.com OpenAPI files: /track/v1/trackingnumbers,
// /rate/v1/rates/quotes, /ship/v1/shipments, PUT /ship/v1/shipments/cancel).

const TRACKING_NUMBER = /^[A-Za-z0-9]{8,40}$/;

async function ok(
  ctx: ToolContext,
  mark: (s: number) => void,
  req: Parameters<ToolContext["call"]>[0],
) {
  const res = await ctx.call(req);
  mark(res.status);
  if (res.status < 200 || res.status >= 300) {
    const detail = vendorError(res.data);
    throw new VendorFailed(
      res.status,
      `FedEx answered ${res.status}${detail ? `: ${detail}` : ""}`,
    );
  }
  return res.data as { output?: unknown };
}

// The account numbers this key may bill or quote on: the key's limit, and never one the
// connection was not set up with.
function allowedAccounts(ctx: ToolContext): Set<string> {
  const connected = connectedAccounts(ctx.config);
  const limit = listLimit(ctx.limits, "accounts");
  return new Set(limit ? connected.filter((a) => limit.includes(a)) : connected);
}

// Every account number anywhere in the request (the shipper's, a payor's for shipping or
// duties, a third party's, a freight account) must be one this key may use. FedEx names them
// accountNumber or ...AccountNumber, as { value } or a plain string.
export function checkAccounts(ctx: ToolContext, body: unknown) {
  const allowed = allowedAccounts(ctx);
  const walk = (node: unknown) => {
    if (Array.isArray(node)) return node.forEach(walk);
    if (!node || typeof node !== "object") return;
    for (const [key, value] of Object.entries(node)) {
      if (key.toLowerCase().endsWith("accountnumber")) {
        const number =
          typeof value === "string"
            ? value
            : value &&
                typeof value === "object" &&
                typeof (value as { value?: unknown }).value === "string"
              ? (value as { value: string }).value
              : null;
        if (number === null)
          throw new Denied('Give each account number as { "value": "<account number>" }');
        if (!allowed.has(number))
          throw new Denied(`This agent is not allowed to use FedEx account ${number}`);
      }
      walk(value);
    }
  };
  walk(body);
}

const accountNumber = z.string().describe("A FedEx account number this agent may use.");
const shipment = z
  .record(z.string(), z.unknown())
  .describe(
    "FedEx requestedShipment object, as in FedEx's API (shipper, recipient(s), pickupType, serviceType, packagingType, requestedPackageLineItems, and for shipments shippingChargesPayment and labelSpecification).",
  );

export const fedexTools: Toolset = {
  names: [
    "fedex_track_shipments",
    "fedex_get_rates",
    "fedex_create_shipment",
    "fedex_cancel_shipment",
  ],

  register(server, ctx, allowed) {
    const readOnly = { readOnlyHint: true, openWorldHint: true } as const;

    if (allowed.has("fedex_track_shipments"))
      server.registerTool(
        "fedex_track_shipments",
        {
          description: "Track up to 30 FedEx shipments by tracking number.",
          inputSchema: z.object({
            tracking_numbers: z.array(z.string()).min(1).max(30),
            include_detailed_scans: z
              .boolean()
              .default(false)
              .describe("Every scan event, not only the latest status."),
          }),
          annotations: readOnly,
        },
        ({ tracking_numbers, include_detailed_scans }) =>
          audited(ctx, "fedex_track_shipments", tracking_numbers.join(","), async (mark) => {
            const bad = tracking_numbers.find((t) => !TRACKING_NUMBER.test(t));
            if (bad) throw new Denied(`${bad} is not a FedEx tracking number`);
            const data = await ok(ctx, mark, {
              method: "POST",
              path: "/track/v1/trackingnumbers",
              body: {
                includeDetailedScans: include_detailed_scans,
                trackingInfo: tracking_numbers.map((trackingNumber) => ({
                  trackingNumberInfo: { trackingNumber },
                })),
              },
            });
            return json(data);
          }),
      );

    if (allowed.has("fedex_get_rates"))
      server.registerTool(
        "fedex_get_rates",
        {
          description:
            "Get FedEx rates and transit times for a shipment, billed to an account this agent may use. Nothing is shipped.",
          inputSchema: z.object({ account_number: accountNumber, requested_shipment: shipment }),
          annotations: readOnly,
        },
        ({ account_number, requested_shipment }) =>
          audited(ctx, "fedex_get_rates", account_number, async (mark) => {
            const body = {
              accountNumber: { value: account_number },
              requestedShipment: requested_shipment,
            };
            checkAccounts(ctx, body);
            const data = await ok(ctx, mark, {
              method: "POST",
              path: "/rate/v1/rates/quotes",
              body,
            });
            return json(data);
          }),
      );

    if (allowed.has("fedex_create_shipment"))
      server.registerTool(
        "fedex_create_shipment",
        {
          description:
            "Create a FedEx shipment and its label, billed to an account this agent may use. This buys the label: charges apply unless it is cancelled. Returns the tracking number and a label link valid for 12 hours.",
          inputSchema: z.object({ account_number: accountNumber, requested_shipment: shipment }),
          annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
        },
        ({ account_number, requested_shipment }) =>
          audited(ctx, "fedex_create_shipment", account_number, async (mark, retarget) => {
            const body = {
              accountNumber: { value: account_number },
              labelResponseOptions: "URL_ONLY",
              requestedShipment: requested_shipment,
            };
            checkAccounts(ctx, body);
            const data = await ok(ctx, mark, { method: "POST", path: "/ship/v1/shipments", body });
            const shipments = (
              data.output as { transactionShipments?: { masterTrackingNumber?: string }[] }
            )?.transactionShipments;
            const tracking = shipments?.map((s) => s.masterTrackingNumber).filter(Boolean) ?? [];
            retarget(`${account_number}:${tracking.join(",") || "no tracking number"}`);
            return json(data);
          }),
      );

    if (allowed.has("fedex_cancel_shipment"))
      server.registerTool(
        "fedex_cancel_shipment",
        {
          description:
            "Cancel a FedEx shipment (all of its packages) created on an account this agent may use.",
          inputSchema: z.object({ account_number: accountNumber, tracking_number: z.string() }),
          annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
        },
        ({ account_number, tracking_number }) =>
          audited(
            ctx,
            "fedex_cancel_shipment",
            `${account_number}:${tracking_number}`,
            async (mark) => {
              if (!TRACKING_NUMBER.test(tracking_number))
                throw new Denied(`${tracking_number} is not a FedEx tracking number`);
              const body = {
                accountNumber: { value: account_number },
                trackingNumber: tracking_number,
                deletionControl: "DELETE_ALL_PACKAGES",
              };
              checkAccounts(ctx, body);
              const data = await ok(ctx, mark, {
                method: "PUT",
                path: "/ship/v1/shipments/cancel",
                body,
              });
              return json(data);
            },
          ),
      );
  },
};
