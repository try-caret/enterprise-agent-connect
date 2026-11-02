import { z } from "zod";
import {
  argumentsTarget,
  audited,
  Denied,
  json,
  VendorFailed,
  type ToolContext,
  type Toolset,
} from "../../packages/connector-kit/src/common.js";
import { lushaError } from "./vendor.js";

// One person or company per call, through Lusha's search-and-enrich (a list of one). Lusha
// charges for the search and for each email and phone it reveals; `reveal` lets the agent ask
// for less.

type Obj = Record<string, any>;
const text = z.string().trim().min(1).max(300);

async function first(ctx: ToolContext, mark: (s: number) => void, path: string, body: Obj) {
  const res = await ctx.call({ method: "POST", path, body });
  mark(res.status);
  if (res.status < 200 || res.status >= 300) throw new VendorFailed(res.status, lushaError(res));
  const data = (res.data ?? {}) as Obj;
  const found: Obj | undefined = data.results?.[0];
  return { found, credits: data.billing?.creditsCharged ?? null };
}

const pick = (o: Obj | undefined, keys: string[]) =>
  o ? Object.fromEntries(keys.filter((k) => o[k] != null).map((k) => [k, o[k]])) : null;

export const lushaTools: Toolset = {
  names: ["lusha_find_contact", "lusha_find_company"],

  register(server, ctx, allowed) {
    if (allowed.has("lusha_find_contact"))
      server.registerTool(
        "lusha_find_contact",
        {
          description:
            "Find one person in Lusha. Send email, or linkedin_url, or person_id, or first_name and last_name with company_name or company_domain. Spends credits for the search and for each email and phone revealed.",
          inputSchema: z
            .object({
              email: text.optional(),
              linkedin_url: text.optional(),
              person_id: text.optional(),
              first_name: text.optional(),
              last_name: text.optional(),
              company_name: text.optional(),
              company_domain: text.optional(),
              reveal: z
                .array(z.enum(["emails", "phones"]))
                .default(["emails", "phones"])
                .describe("What to reveal; each one spends credits. Empty reveals neither."),
            })
            .strict(),
          annotations: { readOnlyHint: true, openWorldHint: true },
        },
        (args) =>
          audited(ctx, "lusha_find_contact", argumentsTarget(args), async (mark) => {
            const named =
              args.first_name && args.last_name && (args.company_name || args.company_domain);
            if (!args.email && !args.linkedin_url && !args.person_id && !named)
              throw new Denied(
                "Send email, linkedin_url or person_id, or first_name and last_name with company_name or company_domain",
              );
            const contact = Object.fromEntries(
              Object.entries({
                id: args.person_id,
                email: args.email,
                linkedinUrl: args.linkedin_url,
                firstName: args.first_name,
                lastName: args.last_name,
                companyName: args.company_name,
                companyDomain: args.company_domain,
              }).filter(([, v]) => v !== undefined),
            );
            const { found, credits } = await first(ctx, mark, "/v3/contacts/search-and-enrich", {
              contacts: [contact],
              reveal: args.reveal,
            });
            if (!found || found.error)
              return json({ found: false, error: found?.error ?? null, credits });
            return json({
              found: true,
              credits,
              person_id: found.id ?? null,
              name: found.fullName ?? null,
              job_title: found.jobTitle ?? null,
              company: found.company ?? null,
              location: pick(found.location, ["city", "state", "country"]),
              emails: (found.emails ?? []).map((e: Obj) =>
                pick(e, ["email", "type", "confidence"]),
              ),
              phones: (found.phones ?? []).map((p: Obj) =>
                pick(p, ["number", "type", "doNotCall"]),
              ),
              linkedin: found.socialLinks?.linkedin ?? null,
              previous_jobs: (found.previousEmployment ?? []).slice(0, 5),
            });
          }),
      );

    if (allowed.has("lusha_find_company"))
      server.registerTool(
        "lusha_find_company",
        {
          description:
            "Find one company in Lusha by domain, name or Lusha id: size, revenue, industry, location, description and funding. Spends credits.",
          inputSchema: z
            .object({
              domain: text.optional(),
              name: text.optional(),
              company_id: text.optional(),
            })
            .strict(),
          annotations: { readOnlyHint: true, openWorldHint: true },
        },
        (args) =>
          audited(ctx, "lusha_find_company", argumentsTarget(args), async (mark) => {
            if (!args.domain && !args.name && !args.company_id)
              throw new Denied("Send domain, name or company_id");
            const company = Object.fromEntries(
              Object.entries({ id: args.company_id, name: args.name, domain: args.domain }).filter(
                ([, v]) => v !== undefined,
              ),
            );
            const { found, credits } = await first(ctx, mark, "/v3/companies/search-and-enrich", {
              companies: [company],
            });
            if (!found || found.error)
              return json({ found: false, error: found?.error ?? null, credits });
            return json({
              found: true,
              credits,
              ...pick(found, [
                "id",
                "name",
                "domain",
                "companyType",
                "yearFounded",
                "employeeCount",
                "revenueRange",
                "industry",
                "subIndustry",
                "location",
                "funding",
                "technologies",
              ]),
              description:
                typeof found.description === "string" ? found.description.slice(0, 1000) : null,
              linkedin: found.socialLinks?.linkedin ?? null,
            });
          }),
      );
  },
};
