import type { BuiltinModule } from "../microsoft/client.js";
import { toolsets } from "../../packages/connector-kit/src/common.js";
import { excelTools } from "./excel.js";
import { graphTools } from "./tools.js";
import { sitesVendor } from "./vendor.js";

export default {
  id: "sharepoint-files",
  create: ({ microsoft }) => ({
    vendor: sitesVendor(microsoft),
    toolset: toolsets(graphTools, excelTools),
  }),
} satisfies BuiltinModule;
