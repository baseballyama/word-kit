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

interface Cursor {
  readonly source: string;
  index: number;
}

function makeCursor(source: string): Cursor {
  return { source, index: 0 };
}

function cursorEof(c: Cursor): boolean {
  return c.index >= c.source.length;
}

function cursorPeek(c: Cursor, offset = 0): string {
  return c.source[c.index + offset] ?? "";
}

function cursorStartsWith(c: Cursor, value: string): boolean {
  return c.source.startsWith(value, c.index);
}

function cursorAdvance(c: Cursor, n = 1): void {
  c.index += n;
}

/** Consume characters up to (but not including) the literal `terminator`. */
function cursorTakeUntil(c: Cursor, terminator: string): string {
  const idx = c.source.indexOf(terminator, c.index);
  if (idx < 0) {
    throw xmlParseError(`Expected ${JSON.stringify(terminator)} but reached end of input`, c);
  }
  const value = c.source.slice(c.index, idx);
  c.index = idx;
  return value;
}

/** Consume the literal `expected` or throw. */
function cursorExpect(c: Cursor, expected: string): void {
  if (!cursorStartsWith(c, expected)) {
    throw xmlParseError(`Expected ${JSON.stringify(expected)}`, c);
  }
  c.index += expected.length;
}

/** Skip XML whitespace (space, tab, newline, carriage return). */
function cursorSkipWhitespace(c: Cursor): void {
  while (c.index < c.source.length) {
    const ch = c.source[c.index];
    if (ch === " " || ch === "\t" || ch === "\n" || ch === "\r") {
      c.index++;
    } else {
      return;
    }
  }
}

/**
 * Tagged `Error` subtype thrown by {@link parseXml}. Carries the source
 * offset of the failure. Identity check uses `err.name === "XmlParseError"`
 * — we ship a `XmlParseError.is(err)` helper that also narrows the type.
 */
export interface XmlParseError extends Error {
  readonly name: "XmlParseError";
  readonly position: number;
}

function xmlParseError(message: string, cursor: Cursor): XmlParseError {
  const err = new Error(`${message} at offset ${cursor.index}`) as XmlParseError & {
    name: string;
    position: number;
  };
  err.name = "XmlParseError";
  err.position = cursor.index;
  return err as XmlParseError;
}

/**
 * Sentinel namespace for the `XmlParseError` type. Use `XmlParseError.is(e)`
 * instead of `e instanceof XmlParseError` (the latter is no longer
 * available because we don't subclass `Error` — that would prevent
 * tree-shaking the error type in callers that don't care about it).
 */
export const XmlParseError = {
  is(err: unknown): err is XmlParseError {
    return err instanceof Error && err.name === "XmlParseError";
  },
};

const NAME_START_CHARS = /[:A-Z_a-zÀ-ÖØ-öø-˿Ͱ-ͽͿ-῿‌-‍⁰-↏Ⰰ-⿯、-퟿豈-﷏ﷰ-�]/;
const NAME_CHARS = /[-:A-Z_a-z0-9.·À-ÖØ-öø-˿Ͱ-ͽͿ-῿‌-‍‿-⁀⁰-↏Ⰰ-⿯、-퟿豈-﷏ﷰ-�]/;

function isNameStart(c: string): boolean {
  return NAME_START_CHARS.test(c);
}
function isNameChar(c: string): boolean {
  return NAME_CHARS.test(c);
}

function readName(c: Cursor): string {
  const start = c.index;
  const first = cursorPeek(c);
  if (!isNameStart(first)) {
    throw xmlParseError(`Expected name start character, got ${JSON.stringify(first)}`, c);
  }
  cursorAdvance(c);
  while (!cursorEof(c) && isNameChar(cursorPeek(c))) cursorAdvance(c);
  return c.source.slice(start, c.index);
}

interface RawAttr {
  rawName: string;
  value: string;
  quote: '"' | "'";
}

function readAttr(c: Cursor): RawAttr {
  const rawName = readName(c);
  cursorSkipWhitespace(c);
  cursorExpect(c, "=");
  cursorSkipWhitespace(c);
  const quote = cursorPeek(c);
  if (quote !== '"' && quote !== "'") {
    throw xmlParseError(`Expected attribute quote, got ${JSON.stringify(quote)}`, c);
  }
  cursorAdvance(c);
  const raw = cursorTakeUntil(c, quote);
  cursorAdvance(c);
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
      return { uri: "http://www.w3.org/2000/xmlns/", local: "xmlns", prefix: "" };
    }
    if (isAttr) {
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
    throw xmlParseError(`Unbound namespace prefix: ${prefix}`, c);
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
  if (!cursorStartsWith(c, "<?xml")) return undefined;
  cursorAdvance(c, "<?xml".length);
  const attrs: RawAttr[] = [];
  cursorSkipWhitespace(c);
  while (!cursorStartsWith(c, "?>")) {
    attrs.push(readAttr(c));
    cursorSkipWhitespace(c);
  }
  cursorExpect(c, "?>");

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
  cursorExpect(c, "<!--");
  const value = cursorTakeUntil(c, "-->");
  cursorExpect(c, "-->");
  return { kind: "comment", value };
}

function parsePI(c: Cursor): XmlPI {
  cursorExpect(c, "<?");
  const target = readName(c);
  if (target.toLowerCase() === "xml") {
    throw xmlParseError("Processing instruction target may not be 'xml'", c);
  }
  let data = "";
  const p = cursorPeek(c);
  if (p === " " || p === "\t" || p === "\n" || p === "\r") {
    cursorSkipWhitespace(c);
    data = cursorTakeUntil(c, "?>");
  }
  cursorExpect(c, "?>");
  return { kind: "pi", target, data };
}

function parseCData(c: Cursor): XmlCData {
  cursorExpect(c, "<![CDATA[");
  const value = cursorTakeUntil(c, "]]>");
  cursorExpect(c, "]]>");
  return { kind: "cdata", value };
}

function parseElement(
  c: Cursor,
  parentScope: NamespaceScope,
  parentSpace: "default" | "preserve",
): XmlElement {
  cursorExpect(c, "<");
  const rawName = readName(c);
  const rawAttrs: RawAttr[] = [];
  cursorSkipWhitespace(c);
  while (!cursorStartsWith(c, "/>") && !cursorStartsWith(c, ">")) {
    rawAttrs.push(readAttr(c));
    cursorSkipWhitespace(c);
  }
  const scope = pushNamespaceScope(parentScope, rawAttrs);
  const name = resolveQName(rawName, scope, false, c);
  const attrs = buildAttrs(rawAttrs, scope, c);

  let xmlSpace: "default" | "preserve" = parentSpace;
  for (const a of attrs) {
    if (a.name.uri === XML_NAMESPACE && a.name.local === "space") {
      xmlSpace = a.value === "preserve" ? "preserve" : "default";
    }
  }

  if (cursorStartsWith(c, "/>")) {
    cursorAdvance(c, 2);
    return { kind: "element", name, attrs, children: [], xmlSpace, selfClosing: true };
  }
  cursorExpect(c, ">");

  const children: XmlNode[] = [];
  while (!cursorEof(c)) {
    if (cursorStartsWith(c, "</")) break;
    if (cursorStartsWith(c, "<![CDATA[")) {
      children.push(parseCData(c));
      continue;
    }
    if (cursorStartsWith(c, "<!--")) {
      children.push(parseComment(c));
      continue;
    }
    if (cursorStartsWith(c, "<?")) {
      children.push(parsePI(c));
      continue;
    }
    if (cursorPeek(c) === "<") {
      children.push(parseElement(c, scope, xmlSpace));
      continue;
    }
    const rawText = cursorTakeUntil(c, "<");
    if (rawText.length === 0) continue;
    const value = decodeEntities(rawText);
    if (xmlSpace === "preserve" || value.trim().length > 0 || hasNonElementSiblings(children)) {
      children.push({ kind: "text", value });
    }
  }

  cursorExpect(c, "</");
  const closeName = readName(c);
  if (closeName !== rawName) {
    throw xmlParseError(`Mismatched close tag: expected </${rawName}>, got </${closeName}>`, c);
  }
  cursorSkipWhitespace(c);
  cursorExpect(c, ">");

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
  const cursor = makeCursor(source);
  cursorSkipWhitespace(cursor);
  const declaration = parseDeclaration(cursor);
  cursorSkipWhitespace(cursor);

  const prologue: XmlNode[] = [];
  while (!cursorEof(cursor) && !isElementStart(cursor)) {
    if (cursorStartsWith(cursor, "<!--")) {
      prologue.push(parseComment(cursor));
    } else if (cursorStartsWith(cursor, "<?")) {
      prologue.push(parsePI(cursor));
    } else if (cursorStartsWith(cursor, "<!")) {
      // DOCTYPE not supported in OOXML; skip the declaration safely.
      cursorTakeUntil(cursor, ">");
      cursorExpect(cursor, ">");
    } else if (cursorPeek(cursor) === "<") {
      break;
    } else {
      cursorAdvance(cursor);
    }
    cursorSkipWhitespace(cursor);
  }

  if (cursorEof(cursor)) {
    throw xmlParseError("Document is missing a root element", cursor);
  }

  const rootScope: NamespaceScope = { map: new Map() };
  const root = parseElement(cursor, rootScope, "default");
  cursorSkipWhitespace(cursor);

  const epilogue: XmlNode[] = [];
  while (!cursorEof(cursor)) {
    if (cursorStartsWith(cursor, "<!--")) {
      epilogue.push(parseComment(cursor));
    } else if (cursorStartsWith(cursor, "<?")) {
      epilogue.push(parsePI(cursor));
    } else {
      cursorAdvance(cursor);
    }
    cursorSkipWhitespace(cursor);
  }

  return {
    ...(declaration ? { declaration } : {}),
    prologue,
    root,
    epilogue,
  } satisfies XmlDocument;
}

function isElementStart(c: Cursor): boolean {
  if (cursorPeek(c) !== "<") return false;
  const next = cursorPeek(c, 1);
  return next !== "!" && next !== "?" && next !== "/";
}
