import type { ConnectorModule } from "../../packages/connector-kit/src/connector.js";
import { servicenowTools } from "./tools.js";
import { servicenowVendor } from "./vendor.js";

export default {
  id: "servicenow",
  create: () => ({ vendor: servicenowVendor(), toolset: servicenowTools }),
} satisfies ConnectorModule;
