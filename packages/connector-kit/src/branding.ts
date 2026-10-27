import { z } from "zod";
import { ConnectError } from "./types.js";

// The names this deployment shows and creates: in copy, in the connectors' steps (as
// {brand.*} placeholders) and in what it makes in a customer's tenant. The names it creates
// (mailbox group, Azure role, resource group, bot and app names) must not change once
// customers have made them, so an install sets them once.

const https = z.url({ protocol: /^https$/ });
// Names go into copy and into PowerShell strings, double- and single-quoted, so none of
// $ " ' ` { } < >.
const nameText = z
  .string()
  .trim()
  .regex(/^[\p{L}\p{N} .,&()-]{1,80}$/u);
const name = (fallback: string) => nameText.default(fallback);

export const brandingSchema = z.strictObject({
  productName: name("Enterprise Agent Connect"),
  logoUrl: https.optional(),
  // Prefix of what it names in a tenant (bot app `<prefix>-bot-<id>`, Azure Bot `<prefix>-<8>`).
  resourcePrefix: z
    .string()
    .regex(/^[a-z][a-z0-9-]{1,20}$/)
    .default("eac"),
  microsoft: z
    .strictObject({
      appName: name("Enterprise Agent Connect"),
      mailboxGroup: name("Agent Mailboxes"),
      mailboxGroupAlias: z
        .string()
        .regex(/^[a-z0-9-]{1,64}$/)
        .default("agent-mailboxes"),
      botRole: name("Agent Connect bots"),
    })
    .default({
      appName: "Enterprise Agent Connect",
      mailboxGroup: "Agent Mailboxes",
      mailboxGroupAlias: "agent-mailboxes",
      botRole: "Agent Connect bots",
    }),
  teams: z
    .strictObject({
      resourceGroup: z
        .string()
        .regex(/^[A-Za-z0-9._()-]{1,90}$/)
        .default("agent-connect"),
      // The Teams app manifest's developer. Unset: no Teams package can be made until it is.
      developer: z
        .strictObject({
          name: nameText.max(32),
          websiteUrl: https,
          privacyUrl: https,
          termsOfUseUrl: https,
        })
        .optional(),
    })
    .default({ resourceGroup: "agent-connect" }),
});

export type Branding = z.infer<typeof brandingSchema>;
export type BrandingInput = z.input<typeof brandingSchema>;

export const DEFAULT_BRANDING: Branding = brandingSchema.parse({});

export function parseBranding(raw: unknown): Branding {
  const parsed = brandingSchema.safeParse(raw ?? {});
  if (!parsed.success) throw new Error(`branding: ${z.prettifyError(parsed.error)}`);
  return parsed.data;
}

// ponytail: one branding per process (set by createAgentConnect before anything reads it);
// thread it through the deps if two differently branded instances must share a process.
let current = DEFAULT_BRANDING;
let set = false;

export const brand = () => current;

export function setBranding(b: Branding) {
  if (set && JSON.stringify(b) !== JSON.stringify(current))
    throw new Error("This process already runs with another branding");
  current = b;
  set = true;
}

/** Fills `{brand.a.b}` placeholders; one naming no text field of the branding throws. */
export function fillBrand(text: string, b: Branding = current): string {
  return text.replace(/\{brand\.([A-Za-z.]+)\}/g, (whole, path: string) => {
    let v: unknown = b;
    for (const part of path.split(".")) v = (v as Record<string, unknown> | undefined)?.[part];
    if (typeof v !== "string") throw new Error(`${whole} names no branding text`);
    return v;
  });
}

/** fillBrand over every string in a parsed JSON value. */
export function fillBrandDeep<T>(value: T, b: Branding = current): T {
  if (typeof value === "string") return fillBrand(value, b) as T;
  if (Array.isArray(value)) return value.map((v) => fillBrandDeep(v, b)) as T;
  if (value && typeof value === "object")
    return Object.fromEntries(Object.entries(value).map(([k, v]) => [k, fillBrandDeep(v, b)])) as T;
  return value;
}

/** The Teams manifest's developer, which a package needs. */
export function teamsDeveloper(b: Branding = current) {
  if (!b.teams.developer)
    throw new ConnectError(
      "Set branding.teams.developer (name and https website, privacy and terms addresses) to make Teams packages",
    );
  return b.teams.developer;
}
