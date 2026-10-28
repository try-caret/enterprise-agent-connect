import type { ConnectorModule } from "../../packages/connector-kit/src/connector.js";
import { dynamicsTools } from "./tools.js";
import { dynamicsVendor } from "./vendor.js";

export default {
  id: "dynamics-365",
  create: ({ microsoft }) => ({
    vendor: dynamicsVendor(microsoft),
    toolset: dynamicsTools,
  }),
} satisfies ConnectorModule;
