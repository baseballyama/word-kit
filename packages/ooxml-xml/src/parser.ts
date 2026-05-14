import { decodeEntities } from "./entities.js";
import {
  type QName,
  type XmlAttr,
  type XmlCData,
  type XmlComment,
  type XmlDeclaration,
  type XmlDocument,
  type XmlElement,
  type XmlNode,
  type XmlPI,
  XML_NAMESPACE,
} from "./types.js";

export interface ParseOptions {
  /** Treat an empty `<root/>` document as legal. Default true. */
  readonly allowEmptyRoot?: boolean;
}

class Cursor {
  index = 0;
  constructor(public readonly source: string) {}

  eof(): boolean {
    return this.index >= this.source.length;
  }

  peek(offset = 0): string {
    return this.source[this.index + offset] ?? "";
  }

  startsWith(value: string): boolean {
    return this.source.startsWith(value, this.index);
  }

  advance(n = 1): void {
    this.index += n;
  }

  /** Consume characters up to (but not including) the literal `terminator`. */
  takeUntil(terminator: string): string {
    const idx = this.source.indexOf(terminator, this.index);
    if (idx < 0) {
      throw new XmlParseError(
        `Expected ${JSON.stringify(terminator)} but reached end of input`,
        this,
      );
    }
    const value = this.source.slice(this.index, idx);
    this.index = idx;
    return value;
  }

  /** Consume the literal `expected` or throw. */
  expect(expected: string): void {
    if (!this.startsWith(expected)) {
      throw new XmlParseError(`Expected ${JSON.stringify(expected)}`, this);
    }
    this.index += expected.length;
  }

  /** Skip XML whitespace (space, tab, newline, carriage return). */
  skipWhitespace(): void {
    while (this.index < this.source.length) {
      const c = this.source[this.index];
      if (c === " " || c === "\t" || c === "\n" || c === "\r") {
        this.index++;
      } else {
        return;
      }
    }
  }
}

export class XmlParseError extends Error {
  readonly position: number;
  constructor(message: string, cursor: Cursor) {
    super(`${message} at offset ${cursor.index}`);
    this.position = cursor.index;
  }
}

const NAME_START_CHARS = /[:A-Z_a-zÀ-ÖØ-öø-˿Ͱ-ͽͿ-῿‌-‍⁰-↏Ⰰ-⿯、-퟿豈-﷏ﷰ-�]/;
const NAME_CHARS = /[-:A-Z_a-z0-9.·À-ÖØ-öø-˿Ͱ-ͽͿ-῿‌-‍‿-⁀⁰-↏Ⰰ-⿯、-퟿豈-﷏ﷰ-�]/;

function isNameStart(c: string): boolean {
  return NAME_START_CHARS.test(c);
}
function isNameChar(c: string): boolean {
  return NAME_CHARS.test(c);
}

function readName(c: Cursor): string {
  const start = c.index;
  const first = c.peek();
  if (!isNameStart(first)) {
    throw new XmlParseError(`Expected name start character, got ${JSON.stringify(first)}`, c);
  }
  c.advance();
  while (!c.eof() && isNameChar(c.peek())) c.advance();
  return c.source.slice(start, c.index);
}

interface RawAttr {
  rawName: string;
  value: string;
  quote: '"' | "'";
}

function readAttr(c: Cursor): RawAttr {
  const rawName = readName(c);
  c.skipWhitespace();
  c.expect("=");
  c.skipWhitespace();
  const quote = c.peek();
  if (quote !== '"' && quote !== "'") {
    throw new XmlParseError(`Expected attribute quote, got ${JSON.stringify(quote)}`, c);
  }
  c.advance();
  const raw = c.takeUntil(quote);
  c.advance();
  return { rawName, value: decodeEntities(raw), quote };
}

interface NamespaceScope {
  /** prefix → uri map. Empty prefix maps to the default namespace. */
  readonly map: ReadonlyMap<string, string>;
}

function pushNamespaceScope(parent: NamespaceScope, attrs: readonly RawAttr[]): NamespaceScope {
  let next: Map<string, string> | undefined;
  for (const attr of attrs) {
    if (attr.rawName === "xmlns") {
      if (!next) next = new Map(parent.map);
      next.set("", attr.value);
    } else if (attr.rawName.startsWith("xmlns:")) {
      if (!next) next = new Map(parent.map);
      next.set(attr.rawName.slice("xmlns:".length), attr.value);
    }
  }
  return next ? { map: next } : parent;
}

function resolveQName(raw: string, scope: NamespaceScope, isAttr: boolean, c: Cursor): QName {
  const colon = raw.indexOf(":");
  if (colon < 0) {
    if (raw === "xmlns") {
      // unqualified xmlns is a namespace decl with synthetic uri
      return { uri: "http://www.w3.org/2000/xmlns/", local: "xmlns", prefix: "" };
    }
    if (isAttr) {
      // Unprefixed attributes have no namespace per XML Namespaces 1.0 §6.2.
      return { uri: "", local: raw, prefix: "" };
    }
    const def = scope.map.get("") ?? "";
    return { uri: def, local: raw, prefix: "" };
  }
  const prefix = raw.slice(0, colon);
  const local = raw.slice(colon + 1);
  if (prefix === "xml") {
    return { uri: XML_NAMESPACE, local, prefix };
  }
  if (prefix === "xmlns") {
    return { uri: "http://www.w3.org/2000/xmlns/", local, prefix };
  }
  const uri = scope.map.get(prefix);
  if (uri === undefined) {
    throw new XmlParseError(`Unbound namespace prefix: ${prefix}`, c);
  }
  return { uri, local, prefix };
}

function buildAttrs(
  raws: readonly RawAttr[],
  scope: NamespaceScope,
  c: Cursor,
): readonly XmlAttr[] {
  return raws.map((r) => {
    const name = resolveQName(r.rawName, scope, true, c);
    const isNamespaceDecl = r.rawName === "xmlns" || r.rawName.startsWith("xmlns:");
    return { name, value: r.value, isNamespaceDecl };
  });
}

function parseDeclaration(c: Cursor): XmlDeclaration | undefined {
  if (!c.startsWith("<?xml")) return undefined;
  c.advance("<?xml".length);
  const attrs: RawAttr[] = [];
  c.skipWhitespace();
  while (!c.startsWith("?>")) {
    attrs.push(readAttr(c));
    c.skipWhitespace();
  }
  c.expect("?>");

  let version = "1.0";
  let encoding: string | undefined;
  let standalone: "yes" | "no" | undefined;
  for (const a of attrs) {
    if (a.rawName === "version") version = a.value;
    else if (a.rawName === "encoding") encoding = a.value;
    else if (a.rawName === "standalone") {
      if (a.value === "yes" || a.value === "no") standalone = a.value;
    }
  }
  return {
    version,
    ...(encoding !== undefined ? { encoding } : {}),
    ...(standalone !== undefined ? { standalone } : {}),
  };
}

function parseComment(c: Cursor): XmlComment {
  c.expect("<!--");
  const value = c.takeUntil("-->");
  c.expect("-->");
  return { kind: "comment", value };
}

function parsePI(c: Cursor): XmlPI {
  c.expect("<?");
  const target = readName(c);
  if (target.toLowerCase() === "xml") {
    throw new XmlParseError("Processing instruction target may not be 'xml'", c);
  }
  let data = "";
  if (c.peek() === " " || c.peek() === "\t" || c.peek() === "\n" || c.peek() === "\r") {
    c.skipWhitespace();
    data = c.takeUntil("?>");
  }
  c.expect("?>");
  return { kind: "pi", target, data };
}

function parseCData(c: Cursor): XmlCData {
  c.expect("<![CDATA[");
  const value = c.takeUntil("]]>");
  c.expect("]]>");
  return { kind: "cdata", value };
}

function parseElement(
  c: Cursor,
  parentScope: NamespaceScope,
  parentSpace: "default" | "preserve",
): XmlElement {
  c.expect("<");
  const rawName = readName(c);
  const rawAttrs: RawAttr[] = [];
  c.skipWhitespace();
  while (!c.startsWith("/>") && !c.startsWith(">")) {
    rawAttrs.push(readAttr(c));
    c.skipWhitespace();
  }
  const scope = pushNamespaceScope(parentScope, rawAttrs);
  const name = resolveQName(rawName, scope, false, c);
  const attrs = buildAttrs(rawAttrs, scope, c);

  // Determine effective xml:space for this element's content.
  let xmlSpace: "default" | "preserve" = parentSpace;
  for (const a of attrs) {
    if (a.name.uri === XML_NAMESPACE && a.name.local === "space") {
      xmlSpace = a.value === "preserve" ? "preserve" : "default";
    }
  }

  if (c.startsWith("/>")) {
    c.advance(2);
    return { kind: "element", name, attrs, children: [], xmlSpace, selfClosing: true };
  }
  c.expect(">");

  const children: XmlNode[] = [];
  while (!c.eof()) {
    if (c.startsWith("</")) break;
    if (c.startsWith("<![CDATA[")) {
      children.push(parseCData(c));
      continue;
    }
    if (c.startsWith("<!--")) {
      children.push(parseComment(c));
      continue;
    }
    if (c.startsWith("<?")) {
      children.push(parsePI(c));
      continue;
    }
    if (c.peek() === "<") {
      children.push(parseElement(c, scope, xmlSpace));
      continue;
    }
    // Text content up to the next `<`.
    const rawText = c.takeUntil("<");
    if (rawText.length === 0) continue;
    const value = decodeEntities(rawText);
    if (xmlSpace === "preserve" || value.trim().length > 0 || hasNonElementSiblings(children)) {
      children.push({ kind: "text", value });
    }
  }

  c.expect("</");
  const closeName = readName(c);
  if (closeName !== rawName) {
    throw new XmlParseError(`Mismatched close tag: expected </${rawName}>, got </${closeName}>`, c);
  }
  c.skipWhitespace();
  c.expect(">");

  return { kind: "element", name, attrs, children, xmlSpace, selfClosing: false };
}

function hasNonElementSiblings(children: readonly XmlNode[]): boolean {
  return children.some((n) => n.kind === "text" || n.kind === "cdata");
}

/**
 * Parse a full XML document into a structured AST. Preserves attribute
 * order, namespace prefixes, whitespace (when significant), CDATA sections,
 * comments, and processing instructions.
 */
export function parseXml(source: string, _options: ParseOptions = {}): XmlDocument {
  const cursor = new Cursor(source);
  cursor.skipWhitespace();
  const declaration = parseDeclaration(cursor);
  cursor.skipWhitespace();

  const prologue: XmlNode[] = [];
  while (!cursor.eof() && !isElementStart(cursor)) {
    if (cursor.startsWith("<!--")) {
      prologue.push(parseComment(cursor));
    } else if (cursor.startsWith("<?")) {
      prologue.push(parsePI(cursor));
    } else if (cursor.startsWith("<!")) {
      // DOCTYPE not supported in OOXML; skip the declaration safely.
      cursor.takeUntil(">");
      cursor.expect(">");
    } else if (cursor.peek() === "<") {
      break;
    } else {
      cursor.advance();
    }
    cursor.skipWhitespace();
  }

  if (cursor.eof()) {
    throw new XmlParseError("Document is missing a root element", cursor);
  }

  const rootScope: NamespaceScope = { map: new Map() };
  const root = parseElement(cursor, rootScope, "default");
  cursor.skipWhitespace();

  const epilogue: XmlNode[] = [];
  while (!cursor.eof()) {
    if (cursor.startsWith("<!--")) {
      epilogue.push(parseComment(cursor));
    } else if (cursor.startsWith("<?")) {
      epilogue.push(parsePI(cursor));
    } else {
      cursor.advance();
    }
    cursor.skipWhitespace();
  }

  return {
    ...(declaration ? { declaration } : {}),
    prologue,
    root,
    epilogue,
  } satisfies XmlDocument;
}

function isElementStart(c: Cursor): boolean {
  if (c.peek() !== "<") return false;
  const next = c.peek(1);
  return next !== "!" && next !== "?" && next !== "/";
}
