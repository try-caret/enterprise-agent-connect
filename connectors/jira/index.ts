import type { ConnectorModule } from "../../packages/connector-kit/src/connector.js";
import { jiraTools } from "./tools.js";
import { jiraVendor } from "./vendor.js";

export default {
  id: "jira",
  create: () => ({ vendor: jiraVendor(), toolset: jiraTools }),
} satisfies ConnectorModule;
