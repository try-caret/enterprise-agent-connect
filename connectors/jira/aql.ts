import { JqlRejected, ORDER_BY, splitOrderBy } from "./jql.js";

// Confines an agent's AQL (Jira Assets Query Language) to the object schemas its grant allows,
// the way scopeJql confines JQL to projects: the agent's clause is kept whole inside
// parentheses after the schema filter, so it has to be lexically closed on its own. AQL
// strings are double-quoted with backslash escapes ("Use Assets Query Language (AQL)" in
// Atlassian's Jira Service Management Cloud docs); objectSchemaId is one of its keywords.
// The tools also drop any answer outside the schemas, so this is not the only guard.

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
  // The filter limits only the objects answered. AQL that looks through references (HAVING
  // inboundReferences/outboundReferences, or dot notation such as Owner.Name) could test objects
  // in other schemas, and a single quote might start a string this check does not know, so
  // neither is accepted from a limited agent. Double-quoted strings are blanked first.
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
