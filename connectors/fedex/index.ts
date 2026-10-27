import type { ConnectorModule } from "../../packages/connector-kit/src/connector.js";
import { fedexTools } from "./tools.js";
import { fedexVendor } from "./vendor.js";

export default {
  id: "fedex",
  create: () => ({ vendor: fedexVendor(), toolset: fedexTools }),
} satisfies ConnectorModule;
