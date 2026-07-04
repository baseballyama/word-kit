/**
 * Tiny XML reader/writer for OPC's well-known fixed-shape files:
 *
 *   - `[Content_Types].xml` — root `<Types>` with `<Default>` and `<Override>` children.
 *   - `*.rels` — root `<Relationships>` with `<Relationship>` children.
 *
 * Scope is intentionally narrow: a flat single-level element list plus an XML
 * declaration. No mixed content, no CDATA, no comments, no entities beyond
 * the five standard ones, no processing instructions other than `<?xml?>`.
 *
 * If a real-world OPC package contains anything outside this shape in those
 * specific files (extremely rare), we fall back to preserving the original
 * bytes verbatim — the package layer decides when to re-serialize.
 */

export interface MiniElement {
  readonly name: string;
  readonly attrs: ReadonlyArray<readonly [string, string]>;
}

export interface MiniDocument {
  /** Root element name. */
  readonly root: string;
  /** Attributes on the root element (typically `xmlns`). */
  readonly rootAttrs: ReadonlyArray<readonly [string, string]>;
  /** Direct child elements of the root. */
  readonly children: ReadonlyArray<MiniElement>;
  /** The `standalone` value from the XML declaration if present. */
  readonly standalone?: "yes" | "no";
}

const XML_DECL_REGEX = /^\s*<\?xml\s+([^?]*)\?>/;
// Matches `<Name attr=...>` or `<Name attr=.../>`. Attribute values may use
// either ' or " quoting. Captures the body of the tag for further parsing.
const TAG_REGEX = /<([A-Za-z_:][\w:.-]*)([^>]*?)(\/?)>/g;
const ATTR_REGEX = /([A-Za-z_:][\w:.-]*)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

const ENTITIES: Record<string, string> = {
  amp: "&",
  lt: "<",
  gt: ">",
  quot: '"',
  apos: "'",
};

function decodeEntities(input: string): string {
  return input.replace(/&([a-z]+|#x?[0-9A-Fa-f]+);/g, (_, body: string) => {
    if (body.startsWith("#x") || body.startsWith("#X")) {
      const code = Number.parseInt(body.slice(2), 16);
      return Number.isFinite(code) ? String.fromCodePoint(code) : `&${body};`;
    }
    if (body.startsWith("#")) {
      const code = Number.parseInt(body.slice(1), 10);
      return Number.isFinite(code) ? String.fromCodePoint(code) : `&${body};`;
    }
    const ent = ENTITIES[body];
    return ent !== undefined ? ent : `&${body};`;
  });
}

function encodeAttrValue(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function parseAttrs(body: string): Array<[string, string]> {
  const result: Array<[string, string]> = [];
  ATTR_REGEX.lastIndex = 0;
  let match: RegExpExecArray | null;
  while ((match = ATTR_REGEX.exec(body)) !== null) {
    const name = match[1] ?? "";
    const value = match[2] !== undefined ? match[2] : (match[3] ?? "");
    result.push([name, decodeEntities(value)]);
  }
  return result;
}

export function parseMiniXml(xml: string): MiniDocument {
  let standalone: "yes" | "no" | undefined;
  let body = xml;
  const declMatch = XML_DECL_REGEX.exec(body);
  if (declMatch) {
    const attrs = parseAttrs(declMatch[1] ?? "");
    for (const [k, v] of attrs) {
      if (k === "standalone") {
        standalone = v === "yes" ? "yes" : v === "no" ? "no" : undefined;
      }
    }
    body = body.slice(declMatch[0].length);
  }

  TAG_REGEX.lastIndex = 0;
  const tags: Array<{
    name: string;
    attrs: Array<[string, string]>;
    selfClosing: boolean;
    closing: boolean;
  }> = [];
  let m: RegExpExecArray | null;
  while ((m = TAG_REGEX.exec(body)) !== null) {
    const name = m[1] ?? "";
    const inner = (m[2] ?? "").trim();
    const closing = name.startsWith("/");
    const realName = closing ? name.slice(1) : name;
    const selfClosing = m[3] === "/";
    const attrs = closing ? [] : parseAttrs(inner);
    tags.push({ name: realName, attrs, selfClosing, closing });
  }

  if (tags.length === 0) {
    throw new Error("mini-xml: no elements found");
  }

  const first = tags[0];
  if (!first || first.closing) {
    throw new Error("mini-xml: first tag is a closing tag");
  }

  const children: MiniElement[] = [];
  const rootName = first.name;
  const rootAttrs = first.attrs;
  if (first.selfClosing) {
    return { root: rootName, rootAttrs, children, ...(standalone ? { standalone } : {}) };
  }

  for (let i = 1; i < tags.length; i++) {
    const t = tags[i];
    if (!t) continue;
    if (t.closing && t.name === rootName) break;
    if (t.closing) continue;
    if (!t.selfClosing) {
      // Look ahead for matching close tag — we treat content as empty since
      // OPC's well-known files never have nested elements at this layer.
      let j = i + 1;
      let depth = 1;
      while (j < tags.length && depth > 0) {
        const next = tags[j];
        if (!next) {
          j++;
          continue;
        }
        if (next.closing && next.name === t.name) depth--;
        else if (!next.selfClosing && !next.closing && next.name === t.name) depth++;
        j++;
      }
      i = j - 1;
    }
    children.push({ name: t.name, attrs: t.attrs });
  }

  return { root: rootName, rootAttrs, children, ...(standalone ? { standalone } : {}) };
}

export function serializeMiniXml(doc: MiniDocument): string {
  const declStandalone = doc.standalone ? ` standalone="${doc.standalone}"` : "";
  const decl = `<?xml version="1.0" encoding="UTF-8"${declStandalone}?>`;
  const rootAttrStr = doc.rootAttrs.map(([k, v]) => ` ${k}="${encodeAttrValue(v)}"`).join("");
  if (doc.children.length === 0) {
    return `${decl}\r\n<${doc.root}${rootAttrStr}/>`;
  }
  const childStrs = doc.children.map((c) => {
    const a = c.attrs.map(([k, v]) => ` ${k}="${encodeAttrValue(v)}"`).join("");
    return `<${c.name}${a}/>`;
  });
  return `${decl}\r\n<${doc.root}${rootAttrStr}>${childStrs.join("")}</${doc.root}>`;
}
