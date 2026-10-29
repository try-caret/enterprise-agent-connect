import type { ConnectorModule } from "../../packages/connector-kit/src/connector.js";
import { workdayTools } from "./tools.js";
import { workdayVendor } from "./vendor.js";

export default {
  id: "workday",
  create: () => ({ vendor: workdayVendor(), toolset: workdayTools }),
} satisfies ConnectorModule;
