import type { ConnectorModule } from "../../packages/connector-kit/src/connector.js";
import { jamfTools } from "./tools.js";
import { jamfVendor } from "./vendor.js";

export default {
  id: "jamf",
  create: () => ({ vendor: jamfVendor(), toolset: jamfTools }),
} satisfies ConnectorModule;
