// Finds the one object a SOQL query reads, so an agent limited to some objects cannot read
// others. Anything that can reach a second object is refused rather than understood:
// subqueries (child relationships, semi-joins), relationship paths and aliases (a dot outside
// strings and numbers), TYPEOF, and more than one object after FROM.

export class SoqlRejected extends Error {}

type Token = { kind: "word" | "string" | "number" | "punct"; text: string };

function tokenize(soql: string): Token[] {
  const tokens: Token[] = [];
  let i = 0;
  while (i < soql.length) {
    const rest = soql.slice(i);
    const space = /^\s+/.exec(rest);
    if (space) {
      i += space[0].length;
      continue;
    }
    if (rest[0] === "'") {
      let j = 1;
      while (j < rest.length && rest[j] !== "'") j += rest[j] === "\\" ? 2 : 1;
      if (j >= rest.length) throw new SoqlRejected("The query has an unterminated string");
      tokens.push({ kind: "string", text: rest.slice(0, j + 1) });
      i += j + 1;
      continue;
    }
    // Numbers, dates and datetimes: 1.5, 2024-01-31, 2024-01-31T00:00:00Z, 2024-01-31T00:00:00+05:00
    const number = /^\d[\d.:TZ+-]*/.exec(rest);
    const word = /^[A-Za-z_][A-Za-z0-9_]*/.exec(rest);
    const [kind, text] = number
      ? (["number", number[0]] as const)
      : word
        ? (["word", word[0]] as const)
        : (["punct", rest[0]] as const);
    tokens.push({ kind, text });
    i += text.length;
  }
  return tokens;
}

const isWord = (t: Token | undefined, w: string) =>
  t?.kind === "word" && t.text.toUpperCase() === w;

export function soqlObject(soql: string): string {
  const tokens = tokenize(soql);
  if (!isWord(tokens[0], "SELECT")) throw new SoqlRejected("The query must start with SELECT");
  const count = (w: string) => tokens.filter((t) => isWord(t, w)).length;
  if (count("SELECT") !== 1 || count("FROM") !== 1)
    throw new SoqlRejected(
      "Subqueries are not allowed for this agent. Query one object at a time.",
    );
  if (count("TYPEOF") > 0) throw new SoqlRejected("TYPEOF is not allowed for this agent");
  if (tokens.some((t) => t.kind === "punct" && t.text === "."))
    throw new SoqlRejected(
      "Relationship fields and aliases (anything with a dot) are not allowed for this agent. Query the related object directly.",
    );
  const from = tokens.findIndex((t) => isWord(t, "FROM"));
  const object = tokens[from + 1];
  if (object?.kind !== "word") throw new SoqlRejected("FROM must name an object");
  if (tokens[from + 2]?.text === ",") throw new SoqlRejected("Query one object at a time");
  return object.text;
}
