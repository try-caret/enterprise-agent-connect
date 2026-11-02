import type { ConnectorModule } from "../../packages/connector-kit/src/connector.js";
import { lushaTools } from "./tools.js";
import { lushaVendor } from "./vendor.js";

export default {
  id: "lusha",
  create: () => ({ vendor: lushaVendor(), toolset: lushaTools }),
} satisfies ConnectorModule;
