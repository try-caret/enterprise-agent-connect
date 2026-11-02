import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname } from "node:path";
import { pathToFileURL } from "node:url";
import { z } from "zod";
import { connectorSchema } from "../packages/connector-kit/src/connector.ts";
import { draftConnector } from "../connectors/http/openapi.ts";

// Drafts a connector.json for the http connector from an OpenAPI 3.x spec (JSON):
//   npx tsx scripts/openapi-to-connector.ts <spec file or URL> --id <app id>
//     (--operations <id,id> | --tags <tag,tag>) [--out <file>]
// The draft loads, but a person finishes it: copy, capabilities, limits, answer limits.

async function main() {
  const args = process.argv.slice(2);
  const flag = (name: string) => {
    const i = args.indexOf(`--${name}`);
    return i < 0 ? undefined : args.splice(i, 2)[1];
  };
  const id = flag("id");
  const operations = flag("operations")?.split(",");
  const tags = flag("tags")?.split(",");
  const source = args[0];
  if (!source || !id || !/^[a-z0-9-]+$/.test(id) || (!operations && !tags)) {
    console.error(
      "Usage: npx tsx scripts/openapi-to-connector.ts <spec file or URL> --id <app-id> (--operations <id,id> | --tags <tag,tag>) [--out <file>]",
    );
    process.exit(2);
  }
  const out = flag("out") ?? `${id}/connector.json`;
  const remote = /^https?:\/\//.test(source);
  const raw = remote ? await (await fetch(source)).text() : readFileSync(source, "utf8");
  let spec: Record<string, unknown>;
  try {
    spec = JSON.parse(raw);
  } catch {
    throw new Error("The spec is not JSON. A YAML spec needs converting to JSON first");
  }
  const { connector, notes } = draftConnector(spec, {
    id,
    operations,
    tags,
    source: remote ? source : undefined,
  });
  mkdirSync(dirname(out), { recursive: true });
  writeFileSync(out, `${JSON.stringify(connector, null, 2)}\n`);
  const check = connectorSchema.safeParse(connector);
  console.log(
    `Wrote ${out}: a DRAFT with ${connector.http.tools.length} tools, not a finished connector. Before using it:`,
  );
  for (const note of notes) console.log(`- ${note}`);
  if (!check.success) console.log(`It does not load yet:\n${z.prettifyError(check.error)}`);
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href)
  main().catch((error) => {
    console.error((error as Error).message);
    process.exit(1);
  });
