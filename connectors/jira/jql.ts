// Confines an agent's JQL to the projects its grant allows. The agent's clause is kept
// whole inside parentheses, so it has to be lexically closed on its own: balanced
// parentheses outside strings, no unterminated string. Anything else could close our
// parenthesis early and OR its way out of the project filter, so it is refused.

export class JqlRejected extends Error {}

const PROJECT_KEY = /^[A-Z][A-Z0-9_]{0,254}$/;
// After strings are blanked to "": fields (cf[10010], "Story Points") with optional direction.
const SORT_FIELD = String.raw`(?:""|[A-Za-z_][\w.]*(?:\[\d+\])?)(?:\s+(?:asc|desc))?`;
const ORDER_BY = new RegExp(String.raw`^order\s+by\s+${SORT_FIELD}(?:\s*,\s*${SORT_FIELD})*$`, "i");

export function isProjectKey(key: string): boolean {
  return PROJECT_KEY.test(key);
}

// Splits at a top-level ORDER BY (outside strings and parentheses).
function splitOrderBy(jql: string): { where: string; orderBy: string | null } {
  let depth = 0;
  let quote: string | null = null;
  let orderAt = -1;
  for (let i = 0; i < jql.length; i++) {
    const c = jql[i];
    if (quote) {
      if (c === "\\") i++;
      else if (c === quote) quote = null;
      continue;
    }
    if (c === '"' || c === "'") quote = c;
    else if (c === "(") depth++;
    else if (c === ")") {
      depth--;
      if (depth < 0) throw new JqlRejected("Unbalanced parentheses in the query");
    } else if (
      orderAt < 0 &&
      depth === 0 &&
      /^order\s+by\b/i.test(jql.slice(i)) &&
      (i === 0 || /[\s)]/.test(jql[i - 1]))
    ) {
      orderAt = i;
    }
  }
  if (quote) throw new JqlRejected("Unterminated string in the query");
  if (depth !== 0) throw new JqlRejected("Unbalanced parentheses in the query");
  if (orderAt < 0) return { where: jql.trim(), orderBy: null };
  return { where: jql.slice(0, orderAt).trim(), orderBy: jql.slice(orderAt).trim() };
}

export function scopeJql(jql: string, projects: string[] | null): string {
  const { where, orderBy } = splitOrderBy(jql);
  if (orderBy && !ORDER_BY.test(orderBy.replace(/"(?:\\.|[^"\\])*"|'(?:\\.|[^'\\])*'/g, '""')))
    throw new JqlRejected("ORDER BY must be the last part of the query and list only fields");
  if (projects === null) return jql.trim();
  if (projects.length === 0) throw new JqlRejected("No projects are allowed");
  for (const p of projects) if (!isProjectKey(p)) throw new Error(`Invalid project key ${p}`);
  const filter = `project in (${projects.map((p) => `"${p}"`).join(", ")})`;
  const scoped = where ? `${filter} AND (${where})` : filter;
  return orderBy ? `${scoped} ${orderBy}` : scoped;
}
