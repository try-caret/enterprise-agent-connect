import type { ConnectorModule } from "../../packages/connector-kit/src/connector.js";
import { dellTools } from "./tools.js";
import { dellVendor } from "./vendor.js";

export default {
  id: "dell-techdirect",
  create: () => ({ vendor: dellVendor(), toolset: dellTools }),
} satisfies ConnectorModule;
