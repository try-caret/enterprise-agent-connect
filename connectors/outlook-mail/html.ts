// Agent-written HTML for mail and event bodies, rebuilt from an allowlist: the output holds only
// tags and attributes written here, with every value re-escaped, so nothing in the input reaches
// the recipient's mail client as markup unless it was allowed.

const TAGS = new Set([
  ..."p br div span b strong i em u s strike sub sup small big h1 h2 h3 h4 h5 h6 ul ol li".split(
    " ",
  ),
  ..."blockquote pre code hr a img table thead tbody tfoot tr td th caption colgroup col".split(
    " ",
  ),
]);
const VOID = new Set(["br", "hr", "img", "col"]);
// Dropped with everything inside them.
const DROP = new Set(
  "script style iframe frame frameset object embed applet form textarea select option button template noscript noembed noframes svg math title head xmp plaintext".split(
    " ",
  ),
);
const ATTRS: Record<string, string[]> = {
  "*": ["style", "align", "dir", "title"],
  a: ["href"],
  img: ["src", "alt", "width", "height"],
  table: ["border", "cellpadding", "cellspacing", "width"],
  td: ["colspan", "rowspan", "valign", "width", "height"],
  th: ["colspan", "rowspan", "valign", "width", "height"],
  col: ["span", "width"],
  colgroup: ["span", "width"],
  ol: ["start"],
};
const STYLE_PROPS = new Set([
  ..."color background-color font-weight font-style font-size font-family text-align".split(" "),
  ..."text-decoration vertical-align line-height white-space width height min-width max-width".split(
    " ",
  ),
  ..."padding padding-top padding-right padding-bottom padding-left margin margin-top".split(" "),
  ..."margin-right margin-bottom margin-left border border-top border-right border-bottom".split(
    " ",
  ),
  ..."border-left border-color border-style border-width border-collapse border-spacing".split(" "),
  "border-radius",
  "display",
]);
const STYLE_VALUE = /^[#a-z0-9.,%\s'"+-]+$|^(?:rgba?|hsla?)\([0-9.,%\s]+\)$/i;
const PLAIN_VALUE = /^[a-z0-9 .%#,_-]{0,200}$/i;
const ENTITY = /&(?:#\d{1,7}|#x[0-9a-f]{1,6}|[a-z][a-z0-9]{1,31});/iy;
const NAMED: Record<string, string> = {
  amp: "&",
  lt: "<",
  gt: ">",
  quot: '"',
  apos: "'",
  nbsp: " ",
};

/** Text as the browser reads an attribute value: entities decoded (so `&#106;avascript:` is caught). */
function decode(value: string): string {
  return value.replace(
    /&(#\d{1,7}|#x[0-9a-f]{1,6}|[a-z][a-z0-9]{1,31});?/gi,
    (whole, e: string) => {
      if (e[0] !== "#") return NAMED[e.toLowerCase()] ?? whole;
      const n = e[1] === "x" || e[1] === "X" ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      return n > 0 && n <= 0x10ffff ? String.fromCodePoint(n) : "";
    },
  );
}

export const escapeHtml = (s: string) =>
  s
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");

// Text keeps the entities it was written with; a bare `&`, `<` or `>` is escaped.
function text(s: string): string {
  let out = "";
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (c === "&") {
      ENTITY.lastIndex = i;
      out += ENTITY.test(s) ? "&" : "&amp;";
    } else out += c === "<" ? "&lt;" : c === ">" ? "&gt;" : c;
  }
  return out;
}

function https(value: string): string | null {
  // Browsers drop tabs and newlines inside a URL and trim control characters around it.
  const url = value.replace(/[\u0000- \u007f]/g, "");
  return /^https:\/\/[^/\\]/i.test(url) && URL.canParse(url) ? url : null;
}

function style(value: string): string | null {
  const kept: string[] = [];
  for (const decl of value.split(";")) {
    const at = decl.indexOf(":");
    if (at < 0) continue;
    const prop = decl.slice(0, at).trim().toLowerCase();
    const v = decl.slice(at + 1).trim();
    if (!STYLE_PROPS.has(prop) || v.length > 200 || !STYLE_VALUE.test(v)) continue;
    if (/expression|url|javascript|behavior|binding/i.test(v)) continue;
    kept.push(`${prop}: ${v}`);
  }
  return kept.length ? kept.join("; ") : null;
}

const ATTR = /([^\s"'>/=]+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'=<>`]+)))?/g;

function attributes(tag: string, raw: string): string {
  const allowed = new Set([...ATTRS["*"], ...(ATTRS[tag] ?? [])]);
  const seen = new Set<string>();
  let out = "";
  for (const m of raw.matchAll(ATTR)) {
    const name = m[1].toLowerCase();
    if (!allowed.has(name) || seen.has(name)) continue;
    seen.add(name);
    const value = decode(m[2] ?? m[3] ?? m[4] ?? "");
    const clean =
      name === "href" || name === "src"
        ? https(value)
        : name === "style"
          ? style(value)
          : PLAIN_VALUE.test(value)
            ? value
            : null;
    if (clean !== null) out += ` ${name}="${escapeHtml(clean)}"`;
  }
  return out;
}

const TAG = /<(\/?)([a-z][a-z0-9]*)((?:[^>"']|"[^"]*"|'[^']*')*)>/iy;

/** HTML safe to send as a message or event body. */
export function sanitizeHtml(input: string): string {
  let out = "";
  let i = 0;
  while (i < input.length) {
    const lt = input.indexOf("<", i);
    if (lt < 0) {
      out += text(input.slice(i));
      break;
    }
    out += text(input.slice(i, lt));
    i = lt;
    if (input.startsWith("<!--", i)) {
      const close = input.indexOf("-->", i + 4);
      i = close < 0 ? input.length : close + 3;
      continue;
    }
    if (input[i + 1] === "!" || input[i + 1] === "?") {
      const close = input.indexOf(">", i);
      i = close < 0 ? input.length : close + 1;
      continue;
    }
    TAG.lastIndex = i;
    const m = TAG.exec(input);
    if (!m) {
      out += "&lt;";
      i++;
      continue;
    }
    i = TAG.lastIndex;
    const [, closing, rawName, rawAttrs] = m;
    const tag = rawName.toLowerCase();
    if (DROP.has(tag)) {
      if (closing) continue;
      const end = new RegExp(`</${tag}\\s*>`, "ig");
      end.lastIndex = i;
      i = end.exec(input) ? end.lastIndex : input.length;
      continue;
    }
    if (!TAGS.has(tag)) continue;
    if (closing) {
      if (!VOID.has(tag)) out += `</${tag}>`;
      continue;
    }
    let attrs = attributes(tag, rawAttrs);
    // An image or link whose address was refused is left out, its content kept.
    if (tag === "img" && !attrs.includes(' src="')) continue;
    if (tag === "a" && attrs.includes(' href="')) attrs += ' rel="noopener noreferrer"';
    out += `<${tag}${attrs}>`;
  }
  return out;
}
