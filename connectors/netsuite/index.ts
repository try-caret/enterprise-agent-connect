import type { ConnectorModule } from "../../packages/connector-kit/src/connector.js";
import { netsuiteTools } from "./tools.js";
import { netsuiteVendor } from "./vendor.js";

export default {
  id: "netsuite",
  create: () => ({ vendor: netsuiteVendor(), toolset: netsuiteTools }),
} satisfies ConnectorModule;
