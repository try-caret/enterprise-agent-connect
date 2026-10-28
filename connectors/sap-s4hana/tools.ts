import { z } from "zod";
import { ODATA } from "./vendor.js";
import {
  audited,
  Denied,
  json,
  listLimit,
  VendorFailed,
  type ToolContext,
  type Toolset,
} from "../../packages/connector-kit/src/common.js";
import { ConnectError } from "../../packages/connector-kit/src/types.js";

const BP = `${ODATA}/API_BUSINESS_PARTNER`;
const SALES = `${ODATA}/API_SALES_ORDER_SRV`;
const PURCHASING = `${ODATA}/API_PURCHASEORDER_PROCESS_SRV`;
const STOCK = `${ODATA}/API_MATERIAL_STOCK_SRV`;
const PRODUCT = `${ODATA}/API_PRODUCT_SRV`;
const INVOICES = `${ODATA}/API_SUPPLIERINVOICE_PROCESS_SRV`;
const ITEMS = `${ODATA}/API_OPLACCTGDOCITEMCUBE_SRV`;

const ORG = /^[A-Za-z0-9]{1,4}$/;
const DOC = /^[A-Za-z0-9]{1,10}$/;
const MATERIAL = /^[A-Za-z0-9][A-Za-z0-9 ._/-]{0,39}$/;
const DAY = /^\d{4}-\d{2}-\d{2}$/;

type Row = Record<string, any>;
type Mark = (s: number) => void;

async function ok(ctx: ToolContext, mark: Mark, req: Parameters<ToolContext["call"]>[0]) {
  let res;
  try {
    res = await ctx.call(req);
  } catch (error) {
    if (error instanceof ConnectError) throw new VendorFailed(502, error.message);
    throw error;
  }
  mark(res.status);
  if (res.status < 200 || res.status >= 300) {
    const detail = sapError(res.data);
    throw new VendorFailed(res.status, `SAP answered ${res.status}${detail ? `: ${detail}` : ""}`);
  }
  return res.data as any;
}

// OData v2 errors: { error: { code, message: { lang, value } } }. Anything else (an HTML logon
// page, say) is not shown.
export function sapError(data: unknown): string {
  const e = (data as { error?: { code?: unknown; message?: { value?: unknown } } } | null)?.error;
  const text = e?.message?.value;
  return typeof text === "string" ? text.slice(0, 300) : "";
}

// OData v2 JSON: a collection is d.results, an entity d, an expanded list its own results.
const results = (v: any): Row[] => (Array.isArray(v?.results) ? v.results : []);

// OData v2 dates are "/Date(ms)/" in UTC.
export function day(v: unknown): string | null {
  const m = typeof v === "string" ? /^\/Date\((-?\d+)/.exec(v) : null;
  return m ? new Date(Number(m[1])).toISOString().slice(0, 10) : null;
}
const sapDate = (d: string) => `/Date(${Date.parse(`${d}T00:00:00Z`)})/`;
const num = (v: unknown) => (v === null || v === undefined || v === "" ? null : Number(v));
const text = (v: unknown) => (typeof v === "string" && v !== "" ? v : null);

const lit = (v: string) => `'${v.replaceAll("'", "''")}'`;
const anyOf = (field: string, values: string[]) =>
  `(${values.map((v) => `${field} eq ${lit(v)}`).join(" or ")})`;
const since = (field: string, d: string) => `${field} ge datetime'${d}T00:00:00'`;
const until = (field: string, d: string) => `${field} le datetime'${d}T23:59:59'`;

function checkDay(d: string | undefined) {
  if (d && (!DAY.test(d) || Number.isNaN(Date.parse(`${d}T00:00:00Z`))))
    throw new Denied(`${d} is not a date (YYYY-MM-DD)`);
}

// The values a read may cover for one limit: the one asked for (which must be allowed), else
// the agent's whole list, else null (no limit and none asked).
function scope(ctx: ToolContext, limit: string, asked: string | undefined, what: string) {
  const allowed = listLimit(ctx.limits, limit);
  if (asked !== undefined) {
    if (allowed && !allowed.includes(asked))
      throw new Denied(`This agent is not allowed to use ${what} ${asked}`);
    return [asked];
  }
  return allowed;
}

const allowedValue = (ctx: ToolContext, limit: string, value: unknown) => {
  const allowed = listLimit(ctx.limits, limit);
  return !allowed || (typeof value === "string" && allowed.includes(value));
};

function collection(
  path: string,
  o: {
    select: string[];
    filters: string[];
    orderby?: string;
    expand?: string;
    skip: number;
    top: number;
  },
) {
  const q = [
    `$select=${o.select.join(",")}`,
    ...(o.expand ? [`$expand=${o.expand}`] : []),
    ...(o.filters.length ? [`$filter=${encodeURIComponent(o.filters.join(" and "))}`] : []),
    ...(o.orderby ? [`$orderby=${encodeURIComponent(o.orderby)}`] : []),
    `$top=${o.top + 1}`,
    `$skip=${o.skip}`,
  ];
  return `${path}?${q.join("&")}`;
}

const paging = {
  cursor: z
    .string()
    .regex(/^\d{1,6}$/)
    .optional()
    .describe("The `next` of an earlier answer."),
  limit: z.number().int().min(1).max(100).default(20),
};

// Asks one more row than shown, so `next` is only given when there is more.
async function page(
  ctx: ToolContext,
  mark: Mark,
  path: string,
  o: Omit<Parameters<typeof collection>[1], "skip" | "top">,
  cursor: string | undefined,
  limit: number,
) {
  const skip = Number(cursor ?? 0);
  const rows = results(
    (await ok(ctx, mark, { method: "GET", path: collection(path, { ...o, skip, top: limit }) })).d,
  );
  return { rows: rows.slice(0, limit), next: rows.length > limit ? String(skip + limit) : null };
}

const SALES_STATUS: Record<string, string> = { A: "open", B: "in process", C: "completed" };
const DELIVERY_STATUS: Record<string, string> = {
  A: "not delivered",
  B: "partly delivered",
  C: "fully delivered",
};
const PO_STATUS: Record<string, string> = {
  "01": "draft",
  "02": "active",
  "03": "in approval",
  "04": "partly approved",
  "05": "approved",
  "08": "rejected",
};
const INVOICE_STATUS: Record<string, string> = {
  "5": "posted",
  A: "parked",
  B: "parked and completed",
  D: "held",
  E: "parked and held",
};
const coded = (map: Record<string, string>, v: unknown) =>
  typeof v === "string" && v !== "" ? (map[v] ?? v) : null;

const SALES_HEADER = [
  "SalesOrder",
  "SalesOrderType",
  "SalesOrganization",
  "DistributionChannel",
  "OrganizationDivision",
  "SoldToParty",
  "PurchaseOrderByCustomer",
  "CreationDate",
  "RequestedDeliveryDate",
  "TotalNetAmount",
  "TransactionCurrency",
  "OverallSDProcessStatus",
  "OverallTotalDeliveryStatus",
];

const salesOrder = (o: Row) => ({
  sales_order: o.SalesOrder,
  type: o.SalesOrderType,
  sales_organization: o.SalesOrganization,
  distribution_channel: o.DistributionChannel,
  division: o.OrganizationDivision,
  sold_to: o.SoldToParty,
  customer_reference: text(o.PurchaseOrderByCustomer),
  created: day(o.CreationDate),
  requested_delivery: day(o.RequestedDeliveryDate),
  net_amount: num(o.TotalNetAmount),
  currency: o.TransactionCurrency,
  status: coded(SALES_STATUS, o.OverallSDProcessStatus),
  delivery_status: coded(DELIVERY_STATUS, o.OverallTotalDeliveryStatus),
});

const PO_HEADER = [
  "PurchaseOrder",
  "PurchaseOrderType",
  "CompanyCode",
  "PurchasingOrganization",
  "PurchasingGroup",
  "Supplier",
  "PurchaseOrderDate",
  "DocumentCurrency",
  "PurchasingProcessingStatus",
  "PurchasingDocumentDeletionCode",
];

const purchaseOrder = (o: Row) => ({
  purchase_order: o.PurchaseOrder,
  type: o.PurchaseOrderType,
  company_code: o.CompanyCode,
  purchasing_organization: o.PurchasingOrganization,
  purchasing_group: o.PurchasingGroup,
  supplier: o.Supplier,
  date: day(o.PurchaseOrderDate),
  currency: o.DocumentCurrency,
  status: coded(PO_STATUS, o.PurchasingProcessingStatus),
  deleted: Boolean(o.PurchasingDocumentDeletionCode),
});

const inPurchasingLimits = (ctx: ToolContext, o: Row) =>
  allowedValue(ctx, "company_codes", o.CompanyCode) &&
  allowedValue(ctx, "purchasing_organizations", o.PurchasingOrganization);

// Plants of new order lines: with a plants limit each line names one of them.
function checkLinePlants(ctx: ToolContext, plants: (string | undefined)[]) {
  const allowed = listLimit(ctx.limits, "plants");
  if (!allowed) return;
  for (const p of plants) {
    if (!p) throw new Denied("Give a plant on every line: this agent is limited to some plants");
    if (!allowed.includes(p)) throw new Denied(`This agent is not allowed to use plant ${p}`);
  }
}

function checkPlants(ctx: ToolContext, items: Row[], field: string, what: string) {
  if (!items.every((i) => !i[field] || allowedValue(ctx, "plants", i[field])))
    throw new Denied(`${what} has lines at a plant this agent is not allowed to use`);
}

const quantity = z.number().positive().max(1e9);

export const sapTools: Toolset = {
  names: [
    "sap_find_business_partners",
    "sap_list_sales_orders",
    "sap_get_sales_order",
    "sap_create_sales_order",
    "sap_list_purchase_orders",
    "sap_get_purchase_order",
    "sap_create_purchase_order",
    "sap_find_products",
    "sap_get_stock",
    "sap_list_supplier_invoices",
    "sap_list_account_items",
  ],

  register(server, ctx, allowed) {
    const readOnly = { readOnlyHint: true, openWorldHint: true } as const;
    const write = { readOnlyHint: false, destructiveHint: false, openWorldHint: true } as const;
    const org = (what: string) => z.string().regex(ORG).describe(what);
    const doc = (what: string) => z.string().regex(DOC).describe(what);
    const date = (what: string) => z.string().describe(`${what}, YYYY-MM-DD.`);
    const target = (args: unknown) => JSON.stringify(args).slice(0, 300);

    if (allowed.has("sap_find_business_partners"))
      server.registerTool(
        "sap_find_business_partners",
        {
          description:
            "Find business partners (customers and suppliers) by name or number. The name match is a case-sensitive part of the full name.",
          inputSchema: z.object({
            name: z.string().min(2).max(80).optional(),
            partner: doc("Business partner, customer or supplier number.").optional(),
            role: z.enum(["any", "customer", "supplier"]).default("any"),
            ...paging,
          }),
          annotations: readOnly,
        },
        (args) =>
          audited(ctx, "sap_find_business_partners", target(args), async (mark) => {
            const filters = [
              ...(args.partner
                ? [
                    `(${["BusinessPartner", "Customer", "Supplier"].map((f) => `${f} eq ${lit(args.partner!)}`).join(" or ")})`,
                  ]
                : []),
              ...(args.name ? [`substringof(${lit(args.name)},BusinessPartnerFullName)`] : []),
              ...(args.role === "customer" ? ["Customer ne ''"] : []),
              ...(args.role === "supplier" ? ["Supplier ne ''"] : []),
            ];
            const { rows, next } = await page(
              ctx,
              mark,
              `${BP}/A_BusinessPartner`,
              {
                select: [
                  "BusinessPartner",
                  "BusinessPartnerFullName",
                  "BusinessPartnerCategory",
                  "BusinessPartnerIsBlocked",
                  "Customer",
                  "Supplier",
                  "to_BusinessPartnerAddress/CityName",
                  "to_BusinessPartnerAddress/Country",
                ],
                expand: "to_BusinessPartnerAddress",
                filters,
                orderby: "BusinessPartner",
              },
              args.cursor,
              args.limit,
            );
            return json({
              partners: rows.map((p) => {
                const address = results(p.to_BusinessPartnerAddress)[0];
                return {
                  partner: p.BusinessPartner,
                  name: p.BusinessPartnerFullName,
                  kind:
                    (
                      { "1": "person", "2": "organization", "3": "group" } as Record<string, string>
                    )[p.BusinessPartnerCategory] ?? null,
                  customer: text(p.Customer),
                  supplier: text(p.Supplier),
                  blocked: Boolean(p.BusinessPartnerIsBlocked),
                  city: text(address?.CityName),
                  country: text(address?.Country),
                };
              }),
              next,
            });
          }),
      );

    if (allowed.has("sap_list_sales_orders"))
      server.registerTool(
        "sap_list_sales_orders",
        {
          description:
            "List sales orders, newest first, with their status and delivery status. Filter by sales organization, customer (sold-to party), the customer's own reference or creation date.",
          inputSchema: z.object({
            sales_organization: org("Sales organization, e.g. 1010.").optional(),
            sold_to: doc("Customer number of the sold-to party.").optional(),
            customer_reference: z
              .string()
              .min(1)
              .max(35)
              .optional()
              .describe("The customer's purchase order number."),
            created_from: date("Created on or after").optional(),
            created_to: date("Created on or before").optional(),
            ...paging,
          }),
          annotations: readOnly,
        },
        (args) =>
          audited(ctx, "sap_list_sales_orders", target(args), async (mark) => {
            checkDay(args.created_from);
            checkDay(args.created_to);
            const orgs = scope(
              ctx,
              "sales_organizations",
              args.sales_organization,
              "sales organization",
            );
            const filters = [
              ...(orgs ? [anyOf("SalesOrganization", orgs)] : []),
              ...(args.sold_to ? [`SoldToParty eq ${lit(args.sold_to)}`] : []),
              ...(args.customer_reference
                ? [`PurchaseOrderByCustomer eq ${lit(args.customer_reference)}`]
                : []),
              ...(args.created_from ? [since("CreationDate", args.created_from)] : []),
              ...(args.created_to ? [until("CreationDate", args.created_to)] : []),
            ];
            const { rows, next } = await page(
              ctx,
              mark,
              `${SALES}/A_SalesOrder`,
              { select: SALES_HEADER, filters, orderby: "SalesOrder desc" },
              args.cursor,
              args.limit,
            );
            return json({
              sales_orders: rows
                .filter((o) => allowedValue(ctx, "sales_organizations", o.SalesOrganization))
                .map(salesOrder),
              next,
            });
          }),
      );

    if (allowed.has("sap_get_sales_order"))
      server.registerTool(
        "sap_get_sales_order",
        {
          description: "Read one sales order with its lines, status and delivery status.",
          inputSchema: z.object({ sales_order: doc("Sales order number.") }),
          annotations: readOnly,
        },
        ({ sales_order }) =>
          audited(ctx, "sap_get_sales_order", sales_order, async (mark) => {
            const o: Row = (
              await ok(ctx, mark, {
                method: "GET",
                path: `${SALES}/A_SalesOrder(${encodeURIComponent(lit(sales_order))})?$expand=to_Item`,
              })
            ).d;
            if (!allowedValue(ctx, "sales_organizations", o?.SalesOrganization))
              throw new Denied(
                `Sales order ${sales_order} is in a sales organization this agent is not allowed to use`,
              );
            const items = results(o.to_Item);
            checkPlants(ctx, items, "ProductionPlant", `Sales order ${sales_order}`);
            return json({
              ...salesOrder(o),
              lines: items.map((i) => ({
                line: i.SalesOrderItem,
                material: i.Material,
                description: text(i.SalesOrderItemText),
                quantity: num(i.RequestedQuantity),
                unit: i.RequestedQuantityUnit,
                net_amount: num(i.NetAmount),
                plant: text(i.ProductionPlant),
                status: coded(SALES_STATUS, i.SDProcessStatus),
                delivery_status: coded(DELIVERY_STATUS, i.DeliveryStatus),
              })),
            });
          }),
      );

    if (allowed.has("sap_create_sales_order"))
      server.registerTool(
        "sap_create_sales_order",
        {
          description:
            "Create a sales order for a customer from structured lines (material and quantity). SAP fills prices, partners and schedule from the customer and material master data.",
          inputSchema: z.object({
            sales_organization: org("Sales organization."),
            distribution_channel: z
              .string()
              .regex(/^[A-Za-z0-9]{2}$/)
              .describe("Distribution channel, e.g. 10."),
            division: z
              .string()
              .regex(/^[A-Za-z0-9]{2}$/)
              .describe("Division, e.g. 00."),
            sold_to: doc("Customer number of the sold-to party."),
            order_type: z
              .string()
              .regex(/^[A-Za-z0-9]{1,4}$/)
              .default("OR")
              .describe("Sales order type; OR is a standard order."),
            customer_reference: z
              .string()
              .min(1)
              .max(35)
              .optional()
              .describe("The customer's purchase order number."),
            requested_delivery_date: date("Requested delivery date").optional(),
            lines: z
              .array(
                z.object({
                  material: z.string().regex(MATERIAL),
                  quantity,
                  unit: z
                    .string()
                    .regex(/^[A-Za-z0-9]{1,3}$/)
                    .optional()
                    .describe("Sales unit; the material's own if omitted."),
                  plant: org("Delivering plant.").optional(),
                }),
              )
              .min(1)
              .max(50),
          }),
          annotations: write,
        },
        (args) =>
          audited(
            ctx,
            "sap_create_sales_order",
            `${args.sales_organization}/${args.sold_to}`,
            async (mark, retarget) => {
              checkDay(args.requested_delivery_date);
              scope(ctx, "sales_organizations", args.sales_organization, "sales organization");
              checkLinePlants(
                ctx,
                args.lines.map((l) => l.plant),
              );
              const created: Row = (
                await ok(ctx, mark, {
                  method: "POST",
                  path: `${SALES}/A_SalesOrder`,
                  body: {
                    SalesOrderType: args.order_type,
                    SalesOrganization: args.sales_organization,
                    DistributionChannel: args.distribution_channel,
                    OrganizationDivision: args.division,
                    SoldToParty: args.sold_to,
                    ...(args.customer_reference
                      ? { PurchaseOrderByCustomer: args.customer_reference }
                      : {}),
                    ...(args.requested_delivery_date
                      ? { RequestedDeliveryDate: sapDate(args.requested_delivery_date) }
                      : {}),
                    to_Item: args.lines.map((l, i) => ({
                      SalesOrderItem: String((i + 1) * 10),
                      Material: l.material,
                      RequestedQuantity: String(l.quantity),
                      ...(l.unit ? { RequestedQuantityUnit: l.unit } : {}),
                      ...(l.plant ? { ProductionPlant: l.plant } : {}),
                    })),
                  },
                })
              ).d;
              retarget(
                `${args.sales_organization}/${args.sold_to} sales order ${created?.SalesOrder}`,
              );
              return json({
                sales_order: created?.SalesOrder ?? null,
                net_amount: num(created?.TotalNetAmount),
                currency: created?.TransactionCurrency ?? null,
                status: coded(SALES_STATUS, created?.OverallSDProcessStatus),
              });
            },
          ),
      );

    if (allowed.has("sap_list_purchase_orders"))
      server.registerTool(
        "sap_list_purchase_orders",
        {
          description:
            "List purchase orders, newest first, with their approval status. Filter by company code, purchasing organization, supplier or order date.",
          inputSchema: z.object({
            company_code: org("Company code, e.g. 1010.").optional(),
            purchasing_organization: org("Purchasing organization.").optional(),
            supplier: doc("Supplier number.").optional(),
            ordered_from: date("Ordered on or after").optional(),
            ordered_to: date("Ordered on or before").optional(),
            ...paging,
          }),
          annotations: readOnly,
        },
        (args) =>
          audited(ctx, "sap_list_purchase_orders", target(args), async (mark) => {
            checkDay(args.ordered_from);
            checkDay(args.ordered_to);
            const codes = scope(ctx, "company_codes", args.company_code, "company code");
            const porgs = scope(
              ctx,
              "purchasing_organizations",
              args.purchasing_organization,
              "purchasing organization",
            );
            const filters = [
              ...(codes ? [anyOf("CompanyCode", codes)] : []),
              ...(porgs ? [anyOf("PurchasingOrganization", porgs)] : []),
              ...(args.supplier ? [`Supplier eq ${lit(args.supplier)}`] : []),
              ...(args.ordered_from ? [since("PurchaseOrderDate", args.ordered_from)] : []),
              ...(args.ordered_to ? [until("PurchaseOrderDate", args.ordered_to)] : []),
            ];
            const { rows, next } = await page(
              ctx,
              mark,
              `${PURCHASING}/A_PurchaseOrder`,
              { select: PO_HEADER, filters, orderby: "PurchaseOrder desc" },
              args.cursor,
              args.limit,
            );
            return json({
              purchase_orders: rows.filter((o) => inPurchasingLimits(ctx, o)).map(purchaseOrder),
              next,
            });
          }),
      );

    if (allowed.has("sap_get_purchase_order"))
      server.registerTool(
        "sap_get_purchase_order",
        {
          description:
            "Read one purchase order with its lines: price, quantity, plant, and whether each line is delivered and invoiced.",
          inputSchema: z.object({ purchase_order: doc("Purchase order number.") }),
          annotations: readOnly,
        },
        ({ purchase_order }) =>
          audited(ctx, "sap_get_purchase_order", purchase_order, async (mark) => {
            const o: Row = (
              await ok(ctx, mark, {
                method: "GET",
                path: `${PURCHASING}/A_PurchaseOrder(${encodeURIComponent(lit(purchase_order))})?$expand=to_PurchaseOrderItem`,
              })
            ).d;
            if (!o || !inPurchasingLimits(ctx, o))
              throw new Denied(
                `Purchase order ${purchase_order} is in a company code or purchasing organization this agent is not allowed to use`,
              );
            const items = results(o.to_PurchaseOrderItem);
            checkPlants(ctx, items, "Plant", `Purchase order ${purchase_order}`);
            return json({
              ...purchaseOrder(o),
              lines: items.map((i) => ({
                line: i.PurchaseOrderItem,
                material: text(i.Material),
                description: text(i.PurchaseOrderItemText),
                plant: text(i.Plant),
                quantity: num(i.OrderQuantity),
                unit: i.PurchaseOrderQuantityUnit,
                net_price: num(i.NetPriceAmount),
                price_per: num(i.NetPriceQuantity),
                delivered: Boolean(i.IsCompletelyDelivered),
                invoiced: Boolean(i.IsFinallyInvoiced),
                deleted: Boolean(i.PurchasingDocumentDeletionCode),
              })),
            });
          }),
      );

    if (allowed.has("sap_create_purchase_order"))
      server.registerTool(
        "sap_create_purchase_order",
        {
          description:
            "Create a purchase order to a supplier. Each line names a material, or a description and material group for an item without a material. It goes through your purchase order approval like any other; this tool cannot approve it.",
          inputSchema: z.object({
            company_code: org("Company code."),
            purchasing_organization: org("Purchasing organization."),
            purchasing_group: z
              .string()
              .regex(/^[A-Za-z0-9]{1,3}$/)
              .describe("Purchasing group."),
            supplier: doc("Supplier number."),
            order_type: z
              .string()
              .regex(/^[A-Za-z0-9]{1,4}$/)
              .default("NB")
              .describe("Purchase order type; NB is a standard order."),
            currency: z
              .string()
              .regex(/^[A-Z]{3}$/)
              .optional()
              .describe("Order currency; the supplier's if omitted."),
            lines: z
              .array(
                z.object({
                  material: z.string().regex(MATERIAL).optional(),
                  description: z.string().min(1).max(40).optional(),
                  material_group: z
                    .string()
                    .regex(/^[A-Za-z0-9]{1,9}$/)
                    .optional(),
                  plant: org("Receiving plant."),
                  quantity,
                  unit: z
                    .string()
                    .regex(/^[A-Za-z0-9]{1,3}$/)
                    .optional()
                    .describe("Order unit; the material's own if omitted."),
                  net_price: z
                    .number()
                    .nonnegative()
                    .max(1e12)
                    .optional()
                    .describe("Price per unit; from the info record if omitted."),
                }),
              )
              .min(1)
              .max(50),
          }),
          annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
        },
        (args) =>
          audited(
            ctx,
            "sap_create_purchase_order",
            `${args.company_code}/${args.supplier}`,
            async (mark, retarget) => {
              scope(ctx, "company_codes", args.company_code, "company code");
              scope(
                ctx,
                "purchasing_organizations",
                args.purchasing_organization,
                "purchasing organization",
              );
              checkLinePlants(
                ctx,
                args.lines.map((l) => l.plant),
              );
              if (args.lines.some((l) => !l.material && !(l.description && l.material_group)))
                throw new Denied(
                  "Give each line a material, or a description and a material group",
                );
              const created: Row = (
                await ok(ctx, mark, {
                  method: "POST",
                  path: `${PURCHASING}/A_PurchaseOrder`,
                  body: {
                    PurchaseOrderType: args.order_type,
                    CompanyCode: args.company_code,
                    PurchasingOrganization: args.purchasing_organization,
                    PurchasingGroup: args.purchasing_group,
                    Supplier: args.supplier,
                    ...(args.currency ? { DocumentCurrency: args.currency } : {}),
                    to_PurchaseOrderItem: args.lines.map((l, i) => ({
                      PurchaseOrderItem: String((i + 1) * 10),
                      Plant: l.plant,
                      OrderQuantity: String(l.quantity),
                      ...(l.material ? { Material: l.material } : {}),
                      ...(l.description ? { PurchaseOrderItemText: l.description } : {}),
                      ...(l.material_group ? { MaterialGroup: l.material_group } : {}),
                      ...(l.unit ? { PurchaseOrderQuantityUnit: l.unit } : {}),
                      ...(l.net_price !== undefined ? { NetPriceAmount: String(l.net_price) } : {}),
                    })),
                  },
                })
              ).d;
              retarget(
                `${args.company_code}/${args.supplier} purchase order ${created?.PurchaseOrder}`,
              );
              return json({
                purchase_order: created?.PurchaseOrder ?? null,
                status: coded(PO_STATUS, created?.PurchasingProcessingStatus),
              });
            },
          ),
      );

    if (allowed.has("sap_find_products"))
      server.registerTool(
        "sap_find_products",
        {
          description:
            "Find products (materials) by number or by a case-sensitive part of their English description.",
          inputSchema: z.object({
            product: z.string().regex(MATERIAL).optional().describe("Product number."),
            description: z.string().min(2).max(40).optional(),
            ...paging,
          }),
          annotations: readOnly,
        },
        (args) =>
          audited(ctx, "sap_find_products", target(args), async (mark) => {
            const filters = [
              "Language eq 'EN'",
              ...(args.product ? [`Product eq ${lit(args.product)}`] : []),
              ...(args.description
                ? [`substringof(${lit(args.description)},ProductDescription)`]
                : []),
            ];
            const { rows, next } = await page(
              ctx,
              mark,
              `${PRODUCT}/A_ProductDescription`,
              { select: ["Product", "ProductDescription"], filters, orderby: "Product" },
              args.cursor,
              args.limit,
            );
            return json({
              products: rows.map((p) => ({
                product: p.Product,
                description: p.ProductDescription,
              })),
              next,
            });
          }),
      );

    if (allowed.has("sap_get_stock"))
      server.registerTool(
        "sap_get_stock",
        {
          description:
            "Stock of one material by plant, storage location and batch, with the unrestricted (available) quantity per plant.",
          inputSchema: z.object({
            material: z.string().regex(MATERIAL).describe("Material (product) number."),
            plant: org("Plant; every plant this agent may use if omitted.").optional(),
          }),
          annotations: readOnly,
        },
        (args) =>
          audited(ctx, "sap_get_stock", target(args), async (mark) => {
            const plants = scope(ctx, "plants", args.plant, "plant");
            const { rows, next } = await page(
              ctx,
              mark,
              `${STOCK}/A_MatlStkInAcctMod`,
              {
                select: [
                  "Material",
                  "Plant",
                  "StorageLocation",
                  "Batch",
                  "InventoryStockType",
                  "InventorySpecialStockType",
                  "MatlWrhsStkQtyInMatlBaseUnit",
                  "MaterialBaseUnit",
                ],
                filters: [
                  `Material eq ${lit(args.material)}`,
                  ...(plants ? [anyOf("Plant", plants)] : []),
                ],
                orderby: "Plant",
              },
              undefined,
              200,
            );
            const kept = rows.filter((r) => allowedValue(ctx, "plants", r.Plant));
            const unrestricted: Record<string, number> = {};
            for (const r of kept)
              if (r.InventoryStockType === "01" && !r.InventorySpecialStockType)
                unrestricted[r.Plant] =
                  (unrestricted[r.Plant] ?? 0) + Number(r.MatlWrhsStkQtyInMatlBaseUnit);
            return json({
              material: args.material,
              unit: kept[0]?.MaterialBaseUnit ?? null,
              unrestricted_by_plant: unrestricted,
              stock: kept.map((r) => ({
                plant: r.Plant,
                storage_location: text(r.StorageLocation),
                batch: text(r.Batch),
                stock_type: coded(
                  { "01": "unrestricted", "02": "quality inspection" },
                  r.InventoryStockType,
                ),
                special_stock: text(r.InventorySpecialStockType),
                quantity: num(r.MatlWrhsStkQtyInMatlBaseUnit),
              })),
              more: next !== null,
            });
          }),
      );

    if (allowed.has("sap_list_supplier_invoices"))
      server.registerTool(
        "sap_list_supplier_invoices",
        {
          description:
            "List supplier invoices, newest posting first, with amount, status and any payment block. Filter by company code, supplier, the supplier's invoice number or posting date. Use sap_list_account_items to see whether they are paid.",
          inputSchema: z.object({
            company_code: org("Company code.").optional(),
            supplier: doc("Supplier (invoicing party) number.").optional(),
            supplier_reference: z
              .string()
              .min(1)
              .max(16)
              .optional()
              .describe("The supplier's own invoice number."),
            posted_from: date("Posted on or after").optional(),
            posted_to: date("Posted on or before").optional(),
            ...paging,
          }),
          annotations: readOnly,
        },
        (args) =>
          audited(ctx, "sap_list_supplier_invoices", target(args), async (mark) => {
            checkDay(args.posted_from);
            checkDay(args.posted_to);
            const codes = scope(ctx, "company_codes", args.company_code, "company code");
            const filters = [
              ...(codes ? [anyOf("CompanyCode", codes)] : []),
              ...(args.supplier ? [`InvoicingParty eq ${lit(args.supplier)}`] : []),
              ...(args.supplier_reference
                ? [`SupplierInvoiceIDByInvcgParty eq ${lit(args.supplier_reference)}`]
                : []),
              ...(args.posted_from ? [since("PostingDate", args.posted_from)] : []),
              ...(args.posted_to ? [until("PostingDate", args.posted_to)] : []),
            ];
            const { rows, next } = await page(
              ctx,
              mark,
              `${INVOICES}/A_SupplierInvoice`,
              {
                select: [
                  "SupplierInvoice",
                  "FiscalYear",
                  "CompanyCode",
                  "InvoicingParty",
                  "SupplierInvoiceIDByInvcgParty",
                  "DocumentDate",
                  "PostingDate",
                  "InvoiceGrossAmount",
                  "DocumentCurrency",
                  "SupplierInvoiceStatus",
                  "PaymentBlockingReason",
                ],
                filters,
                orderby: "PostingDate desc",
              },
              args.cursor,
              args.limit,
            );
            return json({
              invoices: rows
                .filter((r) => allowedValue(ctx, "company_codes", r.CompanyCode))
                .map((r) => ({
                  invoice: r.SupplierInvoice,
                  fiscal_year: r.FiscalYear,
                  company_code: r.CompanyCode,
                  supplier: r.InvoicingParty,
                  supplier_reference: text(r.SupplierInvoiceIDByInvcgParty),
                  document_date: day(r.DocumentDate),
                  posted: day(r.PostingDate),
                  gross_amount: num(r.InvoiceGrossAmount),
                  currency: r.DocumentCurrency,
                  status: coded(INVOICE_STATUS, r.SupplierInvoiceStatus),
                  payment_block: text(r.PaymentBlockingReason),
                })),
              next,
            });
          }),
      );

    if (allowed.has("sap_list_account_items"))
      server.registerTool(
        "sap_list_account_items",
        {
          description:
            "List open or paid (cleared) customer or supplier items in accounting, such as invoices and payments, with due date and the clearing document that paid them.",
          inputSchema: z.object({
            side: z.enum(["customer", "supplier"]),
            account: doc("Customer or supplier number.").optional(),
            company_code: org("Company code.").optional(),
            status: z.enum(["open", "cleared", "all"]).default("open"),
            due_to: date("Due on or before").optional(),
            ...paging,
          }),
          annotations: readOnly,
        },
        (args) =>
          audited(ctx, "sap_list_account_items", target(args), async (mark) => {
            checkDay(args.due_to);
            const codes = scope(ctx, "company_codes", args.company_code, "company code");
            const party = args.side === "customer" ? "Customer" : "Supplier";
            const filters = [
              `FinancialAccountType eq '${args.side === "customer" ? "D" : "K"}'`,
              ...(codes ? [anyOf("CompanyCode", codes)] : []),
              ...(args.account ? [`${party} eq ${lit(args.account)}`] : []),
              ...(args.status === "open" ? ["IsCleared eq false"] : []),
              ...(args.status === "cleared" ? ["IsCleared eq true"] : []),
              ...(args.due_to ? [until("NetDueDate", args.due_to)] : []),
            ];
            const { rows, next } = await page(
              ctx,
              mark,
              `${ITEMS}/A_OperationalAcctgDocItemCube`,
              {
                select: [
                  "CompanyCode",
                  "FiscalYear",
                  "AccountingDocument",
                  "AccountingDocumentItem",
                  "AccountingDocumentType",
                  party,
                  "DocumentReferenceID",
                  "PostingDate",
                  "NetDueDate",
                  "AmountInTransactionCurrency",
                  "TransactionCurrency",
                  "IsCleared",
                  "ClearingAccountingDocument",
                  "ClearingDate",
                ],
                filters,
                orderby: "NetDueDate",
              },
              args.cursor,
              args.limit,
            );
            return json({
              items: rows
                .filter((r) => allowedValue(ctx, "company_codes", r.CompanyCode))
                .map((r) => ({
                  company_code: r.CompanyCode,
                  fiscal_year: r.FiscalYear,
                  document: r.AccountingDocument,
                  line: r.AccountingDocumentItem,
                  document_type: r.AccountingDocumentType,
                  account: r[party],
                  reference: text(r.DocumentReferenceID),
                  posted: day(r.PostingDate),
                  due: day(r.NetDueDate),
                  amount: num(r.AmountInTransactionCurrency),
                  currency: r.TransactionCurrency,
                  cleared: Boolean(r.IsCleared),
                  clearing_document: text(r.ClearingAccountingDocument),
                  cleared_on: day(r.ClearingDate),
                })),
              next,
            });
          }),
      );
  },
};
