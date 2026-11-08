import { JqlRejected, ORDER_BY, splitOrderBy } from "./jql.js";

// Confines an agent's AQL to its granted object schemas, as scopeJql does for projects: the
// agent's clause goes whole inside parentheses after the schema filter. AQL strings are
// double-quoted. The tools also drop any answer outside the schemas.

const SCHEMA_ID = /^[1-9][0-9]{0,17}$/;
// A double-quoted string, its backslash escaping any character (line breaks too), as
// splitOrderBy reads it, so both see the same strings.
const STRING = /"(?:\\[\s\S]|[^"\\])*"/g;
// A number with a decimal point where a value starts (after an operator, `(` or `,`, or at the
// start), with nothing of a name right after it. Anything else with a dot is dot notation.
const DECIMAL = /(?<=(?:^|[(,=<>!~])\s*-?)(?:\d*\.\d+|\d+\.)(?![\p{L}\p{N}\p{M}_.$])/gu;

export function isSchemaId(id: string): boolean {
  return SCHEMA_ID.test(id);
}

export function scopeAql(aql: string, schemas: string[] | null): string {
  const { where, orderBy } = splitOrderBy(aql, '"');
  if (orderBy && !ORDER_BY.test(orderBy.replace(STRING, '""')))
    throw new JqlRejected("ORDER BY must be the last part of the query and name only an attribute");
  if (schemas === null) {
    if (!where) throw new JqlRejected("Give an AQL condition, e.g. objectType = Laptop");
    return aql.trim();
  }
  if (schemas.length === 0) throw new JqlRejected("No object schemas are allowed");
  // Looking through references (HAVING, ...References(, Owner.Name, also in ORDER BY) could test
  // objects in other schemas, connectedTickets( runs JQL no project limit holds, and a single
  // quote might start a string this check does not know: all refused. Numbers keep their dot.
  const bare = aql.replace(STRING, '""').replace(DECIMAL, "0");
  if (/'/.test(bare))
    throw new JqlRejected('Quote values with double quotes ("), not single quotes');
  if (/\bhaving\b|references\s*\(|connectedtickets\s*\(/i.test(bare))
    throw new JqlRejected(
      "References and connected tickets cannot be searched: " +
        "give conditions on the objects' own attributes",
    );
  if (bare.includes("."))
    throw new JqlRejected("Attributes of referenced objects (dot notation) cannot be searched");
  for (const s of schemas) if (!isSchemaId(s)) throw new Error(`Invalid object schema id ${s}`);
  const filter =
    schemas.length === 1
      ? `objectSchemaId = ${schemas[0]}`
      : `objectSchemaId in (${schemas.join(", ")})`;
  const scoped = where ? `${filter} AND (${where})` : filter;
  return orderBy ? `${scoped} ${orderBy}` : scoped;
}
