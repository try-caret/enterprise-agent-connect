import type { ConnectorModule } from "../../packages/connector-kit/src/connector.js";
import { oktaTools } from "./tools.js";
import { oktaVendor } from "./vendor.js";

export default {
  id: "okta",
  create: () => ({ vendor: oktaVendor(), toolset: oktaTools }),
} satisfies ConnectorModule;
