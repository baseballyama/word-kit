const STANDARD_ENTITIES: Record<string, string> = {
  amp: "&",
  lt: "<",
  gt: ">",
  quot: '"',
  apos: "'",
};

/**
 * Decode the five predefined XML entities and numeric character references.
 *
 * Unknown named entities are left untouched (with a leading `&`), which is
 * the safest behavior for OOXML (Word does not emit non-standard entities,
 * so seeing one suggests we should round-trip it verbatim rather than
 * silently corrupt it).
 */
export function decodeEntities(input: string): string {
  return input.replace(
    /&([a-zA-Z][a-zA-Z0-9]*|#x[0-9A-Fa-f]+|#[0-9]+);/g,
    (whole, body: string) => {
      if (body[0] === "#") {
        if (body[1] === "x" || body[1] === "X") {
          const code = Number.parseInt(body.slice(2), 16);
          return Number.isFinite(code) ? String.fromCodePoint(code) : whole;
        }
        const code = Number.parseInt(body.slice(1), 10);
        return Number.isFinite(code) ? String.fromCodePoint(code) : whole;
      }
      const lit = STANDARD_ENTITIES[body];
      return lit !== undefined ? lit : whole;
    },
  );
}

/** Escape a string for use inside an XML text node (`<` and `&` only). */
export function encodeText(input: string): string {
  return input.replace(/[&<]/g, (c) => (c === "&" ? "&amp;" : "&lt;"));
}

/**
 * Escape a string for use inside an XML attribute value. The `quote` argument
 * determines which quote character is being used so we only escape the
 * matching variant.
 */
export function encodeAttrValue(input: string, quote: '"' | "'" = '"'): string {
  let out = input.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  if (quote === '"') {
    out = out.replace(/"/g, "&quot;");
  } else {
    out = out.replace(/'/g, "&apos;");
  }
  // Carriage returns and tabs must be normalized in attribute values per
  // XML 1.0 §3.3.3; we keep them readable by escaping rather than dropping.
  out = out.replace(/\r/g, "&#xD;").replace(/\n/g, "&#xA;").replace(/\t/g, "&#x9;");
  return out;
}
