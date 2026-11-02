import type { ConnectorModule } from "../../packages/connector-kit/src/connector.js";
import { gongTools } from "./tools.js";
import { gongVendor } from "./vendor.js";

export default {
  id: "gong",
  create: () => ({ vendor: gongVendor(), toolset: gongTools }),
} satisfies ConnectorModule;
