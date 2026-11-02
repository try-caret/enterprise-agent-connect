import type { ConnectorModule } from "../../packages/connector-kit/src/connector.js";
import { sapTools } from "./tools.js";
import { sapVendor } from "./vendor.js";

export default {
  id: "sap-s4hana",
  create: () => ({ vendor: sapVendor(), toolset: sapTools }),
} satisfies ConnectorModule;
