// The README's "Use it in your code" apps, as a project outside this repo writes them.
import { z } from "zod";
import type { ConnectorJson } from "@arbium/enterprise-agent-connect";
import {
  audited,
  ConnectError,
  Denied,
  json,
  listLimit,
  VendorFailed,
  type ConnectorModule,
  type Toolset,
} from "@arbium/enterprise-agent-connect/kit";

const setup = {
  identity: "An API key for the app.",
  connectionKind: "guided",
  setup: [{ title: "Create an API key", body: "Ask the team that runs the app for an API key." }],
  afterConnect: [],
  fields: [{ id: "api_key", label: "API key", secret: true }],
  revoke: "Ask the team that runs the app to delete the API key.",
} satisfies Partial<ConnectorJson>;

// No code: the service runs it from its `http` block.
export const orders = (baseUrl: string): ConnectorJson => ({
  ...setup,
  id: "orders",
  name: "Orders",
  description: "Let agents look up orders.",
  logo: "/brand/registry/orders.svg",
  accountKey: "host",
  capabilities: [
    {
      id: "read",
      label: "Read orders",
      description: "Look up an order.",
      scopes: [],
      permissions: [],
      tools: ["orders_get"],
    },
  ],
  limits: [
    {
      id: "customers",
      label: "Customers",
      help: "Only these customers' orders.",
      kind: "list",
      required: false,
    },
  ],
  verify: { method: "GET", path: "/health", label: "Orders answers", show: "name" },
  http: {
    baseUrl,
    allowInsecureHttp: true,
    auth: { type: "header", header: "X-Api-Key", value: "api_key" },
    tools: [
      {
        name: "orders_get",
        description: "Read one order by its id.",
        method: "GET",
        path: "/orders/{id}",
        params: { id: { type: "string", description: "The order id", in: "path" } },
        answerLimits: [{ field: "customer", limit: "customers" }],
      },
    ],
  },
});

// Code: a ConnectorModule built on the kit, carrying its definition inline.
const tools: Toolset = {
  names: ["inventory_get_stock"],
  register(server, ctx, allowed) {
    if (allowed.has("inventory_get_stock"))
      server.registerTool(
        "inventory_get_stock",
        {
          description: "How many of one item are in stock.",
          inputSchema: z.object({ sku: z.string() }),
          annotations: { readOnlyHint: true },
        },
        ({ sku }) =>
          audited(ctx, "inventory_get_stock", sku, async (mark) => {
            const skus = listLimit(ctx.limits, "skus");
            if (skus && !skus.includes(sku)) throw new Denied(`This agent may not read ${sku}`);
            const res = await ctx.call({ method: "GET", path: `/stock/${encodeURIComponent(sku)}` });
            mark(res.status);
            if (res.status !== 200) throw new VendorFailed(res.status, `Inventory answered ${res.status}`);
            return json(res.data);
          }),
      );
  },
};

export const inventory = (baseUrl: string): ConnectorModule => ({
  id: "inventory",
  connector: {
    ...setup,
    id: "inventory",
    name: "Inventory",
    description: "Let agents read stock levels.",
    logo: "/brand/registry/inventory.svg",
    accountKey: "host",
    capabilities: [
      {
        id: "read",
        label: "Read stock",
        description: "Read how many of an item are in stock.",
        scopes: [],
        permissions: [],
        tools: ["inventory_get_stock"],
      },
    ],
    limits: [
      { id: "skus", label: "Items", help: "Only these items.", kind: "list", required: false },
    ],
    verify: { method: "GET", path: "/health", label: "Inventory answers", show: "name" },
  },
  create: () => ({
    toolset: tools,
    vendor: {
      async connect(fields) {
        if (!fields.api_key) throw new ConnectError("Enter the API key");
        return {
          config: { host: new URL(baseUrl).host },
          secret: { api_key: fields.api_key },
          accountLabel: "Inventory",
        };
      },
      async call(connection, request) {
        const res = await fetch(`${baseUrl}${request.path}`, {
          method: request.method,
          headers: { "x-api-key": connection.secret.api_key },
        });
        return { status: res.status, data: await res.json() };
      },
    },
  }),
});
