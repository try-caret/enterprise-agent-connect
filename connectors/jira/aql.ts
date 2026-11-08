import { JqlRejected, ORDER_BY, splitOrderBy } from "./jql.js";

// Confines an agent's AQL to its granted object schemas, as scopeJql does for projects: the
// agent's clause goes whole inside parentheses after the schema filter. AQL strings are
// double-quoted. The tools also drop any answer outside the schemas.

const SCHEMA_ID = /^[1-9][0-9]{0,17}$/;

export function isSchemaId(id: string): boolean {
  return SCHEMA_ID.test(id);
}

export function scopeAql(aql: string, schemas: string[] | null): string {
  const { where, orderBy } = splitOrderBy(aql, '"');
  if (orderBy && !ORDER_BY.test(orderBy.replace(/"(?:\\.|[^"\\])*"/g, '""')))
    throw new JqlRejected("ORDER BY must be the last part of the query and name only an attribute");
  if (schemas === null) {
    if (!where) throw new JqlRejected("Give an AQL condition, e.g. objectType = Laptop");
    return aql.trim();
  }
  if (schemas.length === 0) throw new JqlRejected("No object schemas are allowed");
  // Looking through references (HAVING, ...References(, Owner.Name) could test objects in other
  // schemas, and a single quote might start a string this check does not know: both refused.
  const bare = where.replace(/"(?:\\.|[^"\\])*"/g, '""');
  if (/'/.test(bare))
    throw new JqlRejected('Quote values with double quotes ("), not single quotes');
  if (/\bhaving\b|references\s*\(/i.test(bare))
    throw new JqlRejected(
      "References cannot be searched: give conditions on the objects' own attributes",
    );
  if (/\.\s*["A-Za-z_]/.test(bare))
    throw new JqlRejected("Attributes of referenced objects (dot notation) cannot be searched");
  for (const s of schemas) if (!isSchemaId(s)) throw new Error(`Invalid object schema id ${s}`);
  const filter =
    schemas.length === 1
      ? `objectSchemaId = ${schemas[0]}`
      : `objectSchemaId in (${schemas.join(", ")})`;
  const scoped = where ? `${filter} AND (${where})` : filter;
  return orderBy ? `${scoped} ${orderBy}` : scoped;
}
