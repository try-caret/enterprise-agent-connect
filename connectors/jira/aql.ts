import { JqlRejected, splitOrderBy } from "./jql.js";

// Confines an agent's AQL to its granted object schemas, as scopeJql does for projects: the
// agent's clause goes whole inside parentheses after the schema filter. AQL strings are
// double-quoted. The tools also drop any answer outside the schemas.

const SCHEMA_ID = /^[1-9][0-9]{0,17}$/;
// After strings are blanked to "": attribute names (letters of any script) with a direction.
const SORT_FIELD = String.raw`(?:""|[\p{L}_][\p{L}\p{N}\p{M}_.]*)(?:\s+(?:asc|desc))?`;
const ORDER_BY = new RegExp(
  String.raw`^order\s+by\s+${SORT_FIELD}(?:\s*,\s*${SORT_FIELD})*$`,
  "iu",
);
// A double-quoted string, its backslash escaping any character (line breaks too), as
// splitOrderBy reads it, so both see the same strings.
const STRING = /"(?:\\[\s\S]|[^"\\])*"/g;
// A number with a decimal point where a value starts (after an operator, `(` or `,`, or at the
// start), with nothing of a name right after it. Anything else with a dot is dot notation.
const DECIMAL = /(?<=(?:^|[(,=<>!~])\s*-?)(?:\d*\.\d+|\d+\.)(?![\p{L}\p{N}\p{M}_.$"])/gu;

// Object schema, object type and object ids are all positive numbers.
export function isAssetsId(id: string): boolean {
  return SCHEMA_ID.test(id);
}

// `projectsLimited`: the agent's Jira projects limit is set. connectedTickets( runs JQL inside
// AQL that the projects limit does not hold, so it is refused then, schema limit or not.
export function scopeAql(aql: string, schemas: string[] | null, projectsLimited = false): string {
  const { where, orderBy } = splitOrderBy(aql, '"');
  if (orderBy && !ORDER_BY.test(orderBy.replace(STRING, '""')))
    throw new JqlRejected("ORDER BY must be the last part of the query and name only an attribute");
  if (schemas !== null || projectsLimited) {
    // AQL documents only \" as an escape, so any other backslash is refused, and a single quote
    // might start a string this check does not know: the strings read here are then the
    // strings Atlassian reads, so nothing (connectedTickets included) hides inside one.
    if (/\\(?!")/.test(aql))
      throw new JqlRejected('A backslash may only escape a double quote (\\")');
    if (/'/.test(aql.replace(STRING, '""')))
      throw new JqlRejected('Quote values with double quotes ("), not single quotes');
    if (/connectedtickets\s*\(/i.test(aql.replace(STRING, '""')))
      throw new JqlRejected(
        "Connected tickets cannot be searched by an agent limited to projects or schemas",
      );
  }
  if (schemas === null) {
    if (!where) throw new JqlRejected("Give an AQL condition, e.g. objectType = Laptop");
    return aql.trim();
  }
  // listLimit gives null for an empty list; this guards a direct caller.
  if (schemas.length === 0) throw new JqlRejected("No object schemas are allowed");
  // Looking through references (HAVING, ...References(, Owner.Name, also in ORDER BY) could test
  // objects in other schemas, and a single quote might start a string this check does not know:
  // both refused. Numbers keep their dot.
  // A quoted name may hold a dot too ("Owner.Name"): refused where a name stands, before an
  // operator or anywhere in ORDER BY; a quoted value ("Dell Inc.") keeps its dot.
  for (const m of aql.matchAll(STRING))
    if (
      m[0].includes(".") &&
      (/^\s*(?:[=!<>~]|(?:not|in|is|like|startswith|endswith|having)\b)/i.test(
        aql.slice((m.index ?? 0) + m[0].length),
      ) ||
        (orderBy !== null && (m.index ?? 0) >= aql.lastIndexOf(orderBy)))
    )
      throw new JqlRejected("Attributes of referenced objects (dot notation) cannot be searched");
  const bare = aql.replace(STRING, '""').replace(DECIMAL, "0");
  if (/\bhaving\b|references\s*\(/i.test(bare))
    throw new JqlRejected(
      "References cannot be searched: give conditions on the objects' own attributes",
    );
  if (bare.includes("."))
    throw new JqlRejected(
      "Attributes of referenced objects (dot notation) cannot be searched; " +
        'quote values that hold a dot, e.g. IP = "10.0.0.1"',
    );
  for (const s of schemas) if (!isAssetsId(s)) throw new Error(`Invalid object schema id ${s}`);
  const filter =
    schemas.length === 1
      ? `objectSchemaId = ${schemas[0]}`
      : `objectSchemaId in (${schemas.join(", ")})`;
  const scoped = where ? `${filter} AND (${where})` : filter;
  return orderBy ? `${scoped} ${orderBy}` : scoped;
}
