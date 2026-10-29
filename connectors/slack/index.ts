import type { ConnectorModule } from "../../packages/connector-kit/src/connector.js";
import { slackTools } from "./tools.js";
import { slackVendor } from "./vendor.js";

export default {
  id: "slack",
  create: () => ({ vendor: slackVendor(), toolset: slackTools }),
} satisfies ConnectorModule;
