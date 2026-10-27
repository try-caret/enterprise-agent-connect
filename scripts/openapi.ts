import { writeFileSync } from "node:fs";
import { openapi } from "../packages/server/src/openapi.ts";

// Writes openapi.json from the route list (src/api.ts). CI regenerates it and fails on a diff.
writeFileSync(
  new URL("../packages/server/openapi.json", import.meta.url),
  `${JSON.stringify(openapi(), null, 2)}\n`,
);
