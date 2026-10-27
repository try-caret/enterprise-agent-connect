import assert from "node:assert/strict";
import { readdirSync, readFileSync, statSync } from "node:fs";
import { join } from "node:path";
import { test } from "node:test";
import {
  DEFAULT_BRANDING,
  fillBrand,
  fillBrandDeep,
  parseBranding,
  setBranding,
} from "../../connector-kit/src/branding.ts";
import { loadConnectors, loadProviders } from "../../connector-kit/src/connector.ts";

const custom = parseBranding({
  productName: "Acme Agents",
  resourcePrefix: "acme",
  microsoft: { appName: "Acme Agents App", mailboxGroup: "Acme Mailboxes" },
});

test("branding has defaults for everything and refuses what it does not know", () => {
  assert.equal(DEFAULT_BRANDING.productName, "Enterprise Agent Connect");
  assert.equal(custom.microsoft.mailboxGroupAlias, "agent-mailboxes");
  assert.equal(custom.teams.resourceGroup, "agent-connect");
  assert.throws(() => parseBranding({ productName: "X", colour: "red" }), /colour/);
  // Names reach PowerShell scripts, so nothing that could break out of a string.
  for (const productName of ['A "b"', "A $b", "A `b`", "{brand.x}", "O'Reilly"])
    assert.throws(() => parseBranding({ productName }), /productName/);
  assert.throws(() => parseBranding({ resourcePrefix: "Acme" }), /resourcePrefix/);
});

test("placeholders fill from the branding; one naming nothing throws", () => {
  assert.equal(
    fillBrand("{brand.productName} uses {brand.microsoft.appName}", custom),
    "Acme Agents uses Acme Agents App",
  );
  assert.deepEqual(fillBrandDeep({ a: ["{brand.resourcePrefix}-x", 1, null] }, custom), {
    a: ["acme-x", 1, null],
  });
  assert.throws(() => fillBrand("{brand.nope}", custom), /names no branding text/);
  assert.throws(() => fillBrand("{brand.microsoft}", custom), /names no branding text/);
});

test("loaded connectors carry the branding's names and no placeholder", () => {
  setBranding(custom);
  const text = JSON.stringify([...loadConnectors(), ...loadProviders()]);
  assert.doesNotMatch(text, /\{brand\./);
  assert.match(text, /Acme Agents App/);
  assert.match(text, /Acme Mailboxes/);
});

// The service is not Arbium's: only stored values and names kept for compatibility, each
// documented where it is, may name it.
test("the source and connectors name Arbium only where compatibility needs it", () => {
  const allowed = [
    /export const SHARED_APP_AUTH = "arbium_app";/,
    /\? \{ \.\.\.out, madeByArbium: out\.madeByService \}/,
    /"ARBIUM_MICROSOFT_CLIENT_(ID|SECRET)"/,
  ];
  const files = (dir: string): string[] =>
    readdirSync(dir).flatMap((f) => {
      const p = join(dir, f);
      return statSync(p).isDirectory() ? files(p) : [p];
    });
  const root = new URL("../../..", import.meta.url).pathname;
  const sources = ["packages/server/src", "packages/connector-kit/src", "connectors"];
  const hits = sources
    .flatMap((d) => files(join(root, d)))
    .flatMap((f) =>
      readFileSync(f, "utf8")
        .split("\n")
        .filter((line) => /arbium/i.test(line) && !allowed.some((a) => a.test(line)))
        .map((line) => `${f}: ${line.trim()}`),
    );
  assert.deepEqual(hits, []);
});
