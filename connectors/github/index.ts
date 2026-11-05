import type { ConnectorModule } from "../../packages/connector-kit/src/connector.js";
import { githubTools } from "./tools.js";
import { githubVendor } from "./vendor.js";

export default {
  id: "github",
  create: () => ({ vendor: githubVendor(), toolset: githubTools }),
} satisfies ConnectorModule;
