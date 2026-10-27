import type { ConnectorModule } from "../../packages/connector-kit/src/connector.js";
import { salesforceTools } from "./tools.js";
import { salesforceVendor } from "./vendor.js";

export default {
  id: "salesforce",
  create: () => ({ vendor: salesforceVendor(), toolset: salesforceTools }),
} satisfies ConnectorModule;
