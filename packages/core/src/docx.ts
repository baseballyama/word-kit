import { parseXml, serializeXml } from "@word-kit/ooxml-xml";
import {
  addPart,
  addRelationship,
  allRelationships,
  buildMinimalDocx,
  CONTENT_TYPES_PART_NAME,
  getPart,
  hasPart,
  listParts,
  normalizePartName,
  type OpcPackage,
  packageRelationships,
  partRelationships,
  type Part,
  readOpcPackage,
  type RelationshipSet,
  relationshipsByType,
  removePart,
  removeRelationship,
  setContentTypeDefault,
  writeOpcPackage,
} from "@word-kit/opc";
import {
  acceptAllRevisions as wmlAcceptAllRevisions,
  addSectPrFooterRef,
  addSectPrHeaderRef,
  APP_PROPERTIES_CONTENT_TYPE,
  APP_PROPERTIES_REL_TYPE,
  CORE_PROPERTIES_CONTENT_TYPE,
  CORE_PROPERTIES_REL_TYPE,
  type DocumentAppProperties,
  type DocumentCoreProperties,
  EMPTY_APP_PROPERTIES_XML,
  EMPTY_CORE_PROPERTIES_XML,
  parseAppProperties,
  parseCoreProperties,
  writeAppProperties,
  writeCoreProperties,
  buildAbstractNum,
  buildComment,
  buildCommentRangeEnd,
  buildCommentRangeStart,
  buildCommentReferenceRun,
  buildFieldRuns,
  buildFootnote,
  buildFootnoteReferenceRun,
  buildFooterXml,
  buildHeaderXml,
  buildHyperlink,
  buildHyperlinkRun,
  buildInlineDrawing,
  buildNum,
  buildPageNumberFooterXml,
  buildPPrWithNumPr,
  type BuildStyleOptions,
  buildStyle,
  buildTextTable,
  type BuildTableOptions,
  bulletAbstractNumLevels,
  decimalAbstractNumLevels,
  documentText,
  mergeAdjacentRuns as wmlMergeAdjacentRuns,
  EMPTY_COMMENTS_XML,
  EMPTY_NUMBERING_XML,
  extensionForImageContentType,
  parseFootnotesPart,
  SEED_ENDNOTES_XML,
  SEED_FOOTNOTES_XML,
  type WmlFootnotesPart,
  writeFootnotesPart,
  findStyle,
  styleName,
  findText as wmlFindText,
  type HeaderFooterType,
  MINIMAL_STYLES_XML,
  numAbstractRef,
  numId as readNumId,
  paragraphToElement,
  parseCommentsPart,
  parseParagraph,
  replaceInParagraph,
  type PageMargins,
  PAGE_SIZE_LETTER,
  type PageSize,
  parseNumberingPart,
  parseStylesPart,
  parseWmlDocument,
  rejectAllRevisions as wmlRejectAllRevisions,
  replaceText as wmlReplaceText,
  setSectPrPageMargins,
  setSectPrPageSize,
  sniffImageContentType,
  type TextMatch,
  WML_CONTENT_TYPES,
  WML_NS,
  WML_RELATIONSHIPS,
  type WmlBlock,
  type WmlCommentsPart,
  type WmlDocument,
  type WmlInline,
  type WmlNumberingPart,
  type WmlParagraph,
  type WmlRun,
  type WmlRunPiece,
  type WmlStylesPart,
  type WmlTable,
  writeCommentsPart,
  writeNumberingPart,
  writeStylesPart,
  writeWmlDocument,
} from "@word-kit/wml";
import type { XmlElement, XmlNode } from "@word-kit/ooxml-xml";
import { type ValidationIssue, validatePackage } from "./validator.js";

const DOCUMENT_PART_FALLBACK = "/word/document.xml";

export interface DocxCreateOptions {
  /**
   * Initial paragraphs to seed the body with. Each string becomes one
   * `<w:p>` containing a single run with the given text.
   *
   * - If omitted, the document starts with a single empty paragraph (the
   *   behavior of File → New in Word).
   * - Pass `[]` to start with no paragraphs at all (rare but supported).
   */
  readonly paragraphs?: readonly string[];
}

export interface AddImageOptions {
  /** Image width in EMU (1 inch = 914400 EMU). */
  readonly widthEmu: number;
  /** Image height in EMU. */
  readonly heightEmu: number;
  /** Override the auto-detected content type (e.g. `"image/png"`). */
  readonly contentType?: string;
  /** docPr name (defaults to `Picture {n}`). */
  readonly name?: string;
  /** Accessibility alt text. */
  readonly altText?: string;
}

export interface AppendParagraphOptions {
  /** Value of `w:pStyle` (e.g. `"Heading1"`). Not validated against `styles.xml`. */
  readonly style?: string;
  /** Apply `<w:b/>` to the run. */
  readonly bold?: boolean;
  /** Apply `<w:i/>` to the run. */
  readonly italic?: boolean;
}

/**
 * High-level wrapper around a `.docx` package. Provides the
 * `open` / `create` / `toUint8Array` lifecycle and convenience operations
 * over the main document part.
 *
 * The lower-level building blocks are still reachable via the {@link opc}
 * and {@link document} accessors when callers need fine-grained control.
 */
const STYLES_PART_NAME = "/word/styles.xml";
const NUMBERING_PART_NAME = "/word/numbering.xml";
const COMMENTS_PART_NAME = "/word/comments.xml";
const FOOTNOTES_PART_NAME = "/word/footnotes.xml";
const ENDNOTES_PART_NAME = "/word/endnotes.xml";
const CORE_PROPERTIES_PART_NAME = "/docProps/core.xml";
const APP_PROPERTIES_PART_NAME = "/docProps/app.xml";

const BULLET_ABSTRACT_NUM_ID = 9000;
const DECIMAL_ABSTRACT_NUM_ID = 9001;

export interface AddCommentOptions {
  readonly author: string;
  readonly text: string;
  readonly initials?: string;
  readonly date?: string;
}

/**
 * Plain data shape for a `.docx` document. Behaviour is provided by the
 * standalone functions exported from this module (`createDocx`, `openDocx`,
 * `appendParagraph`, `addTable`, …) so tree-shaking can drop any operation
 * the caller does not import.
 */
export interface Docx {
  opc: OpcPackage;
  document: WmlDocument;
  partName: string;
  dirty: boolean;
  stylesCache: WmlStylesPart | undefined;
  stylesDirty: boolean;
  numberingCache: WmlNumberingPart | undefined;
  numberingDirty: boolean;
  commentsCache: WmlCommentsPart | undefined;
  commentsDirty: boolean;
  footnotesCache: WmlFootnotesPart | undefined;
  footnotesDirty: boolean;
  endnotesCache: WmlFootnotesPart | undefined;
  endnotesDirty: boolean;
}

function makeDocx(opc: OpcPackage, document: WmlDocument, partName: string): Docx {
  return {
    opc,
    document,
    partName,
    dirty: false,
    stylesCache: undefined,
    stylesDirty: false,
    numberingCache: undefined,
    numberingDirty: false,
    commentsCache: undefined,
    commentsDirty: false,
    footnotesCache: undefined,
    footnotesDirty: false,
    endnotesCache: undefined,
    endnotesDirty: false,
  };
}

/** Parse an existing `.docx` package. */
export function openDocx(bytes: Uint8Array): Docx {
  const pkg = readOpcPackage(bytes);
  const part = findMainDocumentPart(pkg);
  if (!part) {
    throw new Error("Package has no main WordprocessingML document part");
  }
  const xml = new TextDecoder("utf-8").decode(part.data);
  const wml = parseWmlDocument(parseXml(xml));
  return makeDocx(pkg, wml, part.name);
}

/**
 * Convenience for browser-side callers: open a `.docx` from a `Blob` or
 * `File`. Awaits the underlying `ArrayBuffer` and delegates to
 * {@link openDocx}.
 */
export async function fromBlob(blob: Blob): Promise<Docx> {
  const buf = await blob.arrayBuffer();
  return openDocx(new Uint8Array(buf));
}

/**
 * Return one entry per header part referenced by `document.xml`. Each
 * entry includes the part name, the relationship id, and the plain-text
 * content extracted from the part.
 */
export function headers(doc: Docx): Array<{ partName: string; relId: string; text: string }> {
  return collectHeaderFooterParts(doc, WML_RELATIONSHIPS.header);
}

/** Same shape as {@link headers} but for footer parts. */
export function footers(doc: Docx): Array<{ partName: string; relId: string; text: string }> {
  return collectHeaderFooterParts(doc, WML_RELATIONSHIPS.footer);
}

/**
 * Enumerate every named bookmark in the body. Each entry includes the
 * bookmark name, its numeric id, and the paragraph that contains the
 * `<w:bookmarkStart>`.
 */
export function bookmarks(doc: Docx): Array<{ name: string; id: number; paragraph: WmlParagraph }> {
  const out: Array<{ name: string; id: number; paragraph: WmlParagraph }> = [];
  for (const block of doc.document.body.blocks) {
    if (block.kind !== "paragraph") continue;
    for (const child of block.children) {
      if (child.kind !== "raw") continue;
      if (child.node.name.uri !== WML_NS) continue;
      if (child.node.name.local !== "bookmarkStart") continue;
      const idAttr = child.node.attrs.find((a) => a.name.uri === WML_NS && a.name.local === "id");
      const nameAttr = child.node.attrs.find(
        (a) => a.name.uri === WML_NS && a.name.local === "name",
      );
      if (!idAttr || !nameAttr) continue;
      const id = Number.parseInt(idAttr.value, 10);
      if (!Number.isFinite(id)) continue;
      out.push({ id, name: nameAttr.value, paragraph: block });
    }
  }
  return out;
}

/**
 * Remove a bookmark by name. Strips both the `<w:bookmarkStart>` and
 * `<w:bookmarkEnd>` markers (matched by the start's id). Returns true if
 * a bookmark with that name was found and removed.
 */
export function removeBookmark(doc: Docx, name: string): boolean {
  let removedId: number | undefined;
  const checkRemoveStart = (children: WmlInline[]): WmlInline[] => {
    return children.filter((child) => {
      if (
        child.kind === "raw" &&
        child.node.name.uri === WML_NS &&
        child.node.name.local === "bookmarkStart"
      ) {
        const nameAttr = child.node.attrs.find(
          (a) => a.name.uri === WML_NS && a.name.local === "name",
        );
        if (nameAttr?.value === name) {
          const idAttr = child.node.attrs.find(
            (a) => a.name.uri === WML_NS && a.name.local === "id",
          );
          if (idAttr) {
            const n = Number.parseInt(idAttr.value, 10);
            if (Number.isFinite(n)) removedId = n;
          }
          return false;
        }
      }
      return true;
    });
  };
  const checkRemoveEnd = (children: WmlInline[]): WmlInline[] => {
    if (removedId === undefined) return children;
    const targetId = String(removedId);
    return children.filter((child) => {
      if (
        child.kind === "raw" &&
        child.node.name.uri === WML_NS &&
        child.node.name.local === "bookmarkEnd"
      ) {
        const idAttr = child.node.attrs.find((a) => a.name.uri === WML_NS && a.name.local === "id");
        return idAttr?.value !== targetId;
      }
      return true;
    });
  };
  for (const block of doc.document.body.blocks) {
    if (block.kind !== "paragraph") continue;
    block.children = checkRemoveStart(block.children);
  }
  if (removedId === undefined) return false;
  for (const block of doc.document.body.blocks) {
    if (block.kind !== "paragraph") continue;
    block.children = checkRemoveEnd(block.children);
  }
  doc.dirty = true;
  return true;
}

/**
 * Return one entry per media (image) part. Useful for inspecting which
 * images a document already contains and feeding their bytes back into
 * {@link replaceImage}.
 */
export function images(
  doc: Docx,
): Array<{ partName: string; contentType: string; data: Uint8Array }> {
  const out: Array<{ partName: string; contentType: string; data: Uint8Array }> = [];
  for (const part of listParts(doc.opc)) {
    if (part.name.startsWith("/word/media/")) {
      out.push({ partName: part.name, contentType: part.contentType, data: part.data });
    }
  }
  return out;
}

/**
 * Replace the bytes of an existing image part in place. The original
 * relationships are kept, so every place in the document that referenced
 * the image now points at the new bytes.
 *
 * Returns `true` if the part was found and updated, `false` otherwise.
 */
export function replaceImage(doc: Docx, partName: string, newBytes: Uint8Array): boolean {
  const part = getPart(doc.opc, partName);
  if (!part) return false;
  part.data = newBytes;
  return true;
}

function collectHeaderFooterParts(
  doc: Docx,
  relType: string,
): Array<{ partName: string; relId: string; text: string }> {
  const out: Array<{ partName: string; relId: string; text: string }> = [];
  const docRels = partRelationships(doc.opc, doc.partName);
  for (const rel of relationshipsByType(docRels, relType)) {
    if (rel.targetMode === "External") continue;
    const partName = resolvePartTarget(doc, rel.target);
    const part = getPart(doc.opc, partName);
    if (!part) continue;
    const xml = new TextDecoder("utf-8").decode(part.data);
    const xmlDoc = parseXml(xml);
    const text = collectVisibleTextFromElement(xmlDoc.root);
    out.push({ partName, relId: rel.id, text });
  }
  return out;
}

/** Create a fresh `.docx`. */
export function createDocx(options: DocxCreateOptions = {}): Docx {
  const pkg = buildMinimalDocx();
  const docPart = getPart(pkg, DOCUMENT_PART_FALLBACK);
  if (!docPart) {
    throw new Error("Minimal docx is missing /word/document.xml");
  }
  // Seed a minimal but valid styles.xml so the produced docx opens cleanly
  // in both Word and LibreOffice without an "uses an unknown style" prompt.
  addPart(pkg, {
    name: STYLES_PART_NAME,
    contentType: WML_CONTENT_TYPES.styles,
    data: new TextEncoder().encode(MINIMAL_STYLES_XML),
  });
  const docRels = partRelationships(pkg, DOCUMENT_PART_FALLBACK);
  addRelationship(docRels, { type: WML_RELATIONSHIPS.styles, target: "styles.xml" });

  const xml = new TextDecoder("utf-8").decode(docPart.data);
  const wml = parseWmlDocument(parseXml(xml));
  const docx = makeDocx(pkg, wml, docPart.name);
  if (options.paragraphs !== undefined) {
    // Clear the seed empty paragraph; the caller is supplying the body.
    wml.body.blocks.length = 0;
    for (const text of options.paragraphs) {
      appendParagraph(docx, text);
    }
  }
  return docx;
}

/** Paragraph blocks in document order. */
export function paragraphs(doc: Docx): readonly WmlParagraph[] {
  return doc.document.body.blocks.filter(isParagraph);
}

/** Table blocks in document order. */
export function tables(doc: Docx): readonly WmlTable[] {
  return doc.document.body.blocks.filter(isTable);
}

/**
 * Walk every paragraph in the body — top-level and inside table cells —
 * and call {@link mergeAdjacentRuns} on each one. Returns the total
 * number of merges performed across the whole body.
 *
 * Useful as a one-shot cleanup pass on a template that's been
 * fragmented into many tiny same-formatting runs (typically by Word's
 * spell-checker). Header/footer/comment/footnote parts are not touched —
 * call this once per side-part if you want to extend coverage.
 */
export function mergeAdjacentRunsInBody(doc: Docx): number {
  let total = 0;
  const visit = (blocks: readonly WmlBlock[]): void => {
    for (const b of blocks) {
      if (b.kind === "paragraph") {
        total += wmlMergeAdjacentRuns(b);
      } else if (b.kind === "table") {
        for (const row of b.rows) {
          for (const cell of row.cells) {
            visit(cell.paragraphs);
          }
        }
      }
    }
  };
  visit(doc.document.body.blocks);
  if (total > 0) doc.dirty = true;
  return total;
}

/**
 * Parsed `word/styles.xml` AST, or `undefined` if the package has no
 * styles part. Lazily parsed on first access; subsequent mutations are
 * flushed back on `toUint8Array()`.
 */
export function stylesPart(doc: Docx): WmlStylesPart | undefined {
  if (doc.stylesCache) return doc.stylesCache;
  const part = getPart(doc.opc, STYLES_PART_NAME);
  if (!part) return undefined;
  const xml = new TextDecoder("utf-8").decode(part.data);
  doc.stylesCache = parseStylesPart(parseXml(xml));
  return doc.stylesCache;
}

/**
 * Add (or replace) a `<w:style>` entry. Creates `word/styles.xml` and
 * its relationship if they do not already exist.
 */
export function addStyle(doc: Docx, options: BuildStyleOptions): void {
  const part = ensureStylesPart(doc);
  const existing = findStyle(part, options.styleId);
  const built = buildStyle(options);
  if (existing) {
    const idx = part.styles.indexOf(existing);
    if (idx >= 0) part.styles[idx] = built;
  } else {
    part.styles.push(built);
  }
  doc.stylesDirty = true;
}

/**
 * Remove a `<w:style>` entry by its `w:styleId`. Returns true if a style
 * with that id existed and was removed. Note that paragraphs and runs
 * referencing the removed style by `w:pStyle` / `w:rStyle` are NOT
 * scrubbed — they will fall back to the Normal style in Word, which is
 * usually the intended behaviour. Use {@link validate} afterwards to
 * surface dangling references as warnings.
 */
export function removeStyle(doc: Docx, styleId: string): boolean {
  const part = stylesPart(doc);
  if (!part) return false;
  const existing = findStyle(part, styleId);
  if (!existing) return false;
  const idx = part.styles.indexOf(existing);
  if (idx < 0) return false;
  part.styles.splice(idx, 1);
  doc.stylesDirty = true;
  return true;
}

/**
 * Enumerate every `<w:style>` entry in `word/styles.xml`. Each item
 * carries the style id and the `w:type` attribute (paragraph, character,
 * table, or numbering). Returns an empty array if the package has no
 * styles part.
 */
export function listStyles(doc: Docx): Array<{ styleId: string; type: string }> {
  const part = stylesPart(doc);
  if (!part) return [];
  const out: Array<{ styleId: string; type: string }> = [];
  for (const style of part.styles) {
    const idAttr = style.attrs.find((a) => a.name.local === "styleId");
    const typeAttr = style.attrs.find((a) => a.name.local === "type");
    if (!idAttr) continue;
    out.push({ styleId: idAttr.value, type: typeAttr?.value ?? "" });
  }
  return out;
}

/**
 * Resolve a style by its `<w:name w:val="…">` display name, returning the
 * matching `styleId`. Useful when a template uses localised style names
 * (e.g., Word in Japanese names "Heading 1" as "見出し 1" but its
 * styleId stays "Heading1"). The match is case-sensitive on the display
 * name. Returns `undefined` if no style has that name.
 */
export function findStyleIdByName(doc: Docx, name: string): string | undefined {
  const part = stylesPart(doc);
  if (!part) return undefined;
  for (const style of part.styles) {
    if (styleName(style) === name) {
      const idAttr = style.attrs.find((a) => a.name.local === "styleId");
      if (idAttr) return idAttr.value;
    }
  }
  return undefined;
}

/**
 * Parsed `word/numbering.xml` AST, or `undefined` if absent. Lazy and
 * cached just like {@link stylesPart}.
 */
export function numberingPart(doc: Docx): WmlNumberingPart | undefined {
  if (doc.numberingCache) return doc.numberingCache;
  const part = getPart(doc.opc, NUMBERING_PART_NAME);
  if (!part) return undefined;
  const xml = new TextDecoder("utf-8").decode(part.data);
  doc.numberingCache = parseNumberingPart(parseXml(xml));
  return doc.numberingCache;
}

/**
 * Append a bullet list (one paragraph per item) to the body. Sets up
 * `numbering.xml` and its relationship on first use.
 */
export function addBulletList(doc: Docx, items: readonly string[]): WmlParagraph[] {
  const idValue = ensureBulletNumbering(doc);
  return items.map((text) => appendListParagraph(doc, text, idValue));
}

/**
 * Append a numbered list (one paragraph per item) to the body. Sets up
 * `numbering.xml` and its relationship on first use.
 */
export function addNumberedList(doc: Docx, items: readonly string[]): WmlParagraph[] {
  const idValue = ensureDecimalNumbering(doc);
  return items.map((text) => appendListParagraph(doc, text, idValue));
}

/**
 * Apply numbering to an existing paragraph (turn it into a list item).
 * Use `numId` from one of `addBulletList` / `addNumberedList` / a value
 * found via `docx.numberingPart`.
 */
export function applyListToParagraph(
  doc: Docx,
  paragraph: WmlParagraph,
  numId: number,
  ilvl = 0,
): void {
  const pPr = paragraph.pPr;
  if (!pPr) {
    paragraph.pPr = buildPPrWithNumPr(numId, ilvl);
  } else {
    // Replace any existing <w:numPr> with the new one.
    const existing = pPr.children.findIndex(
      (c) => c.kind === "element" && c.name.local === "numPr",
    );
    const numPr = buildPPrWithNumPr(numId, ilvl).children.find(
      (c) => c.kind === "element" && c.name.local === "numPr",
    );
    if (numPr && numPr.kind === "element") {
      const children = pPr.children as XmlElement[];
      if (existing >= 0) children[existing] = numPr;
      else children.push(numPr);
    }
  }
  doc.dirty = true;
}

function appendListParagraph(doc: Docx, text: string, numIdValue: number): WmlParagraph {
  const piece: WmlRunPiece = {
    kind: "text",
    value: text,
    preserveSpace: /^\s|\s$/.test(text),
  };
  const run: WmlRun = { kind: "run", pieces: [piece], extras: [] };
  const paragraph: WmlParagraph = {
    kind: "paragraph",
    pPr: buildPPrWithNumPr(numIdValue, 0),
    children: [run],
    extras: [],
  };
  doc.document.body.blocks.push(paragraph);
  doc.dirty = true;
  return paragraph;
}

function ensureBulletNumbering(doc: Docx): number {
  return ensureNumberingDefinition(doc, BULLET_ABSTRACT_NUM_ID, bulletAbstractNumLevels);
}

function ensureDecimalNumbering(doc: Docx): number {
  return ensureNumberingDefinition(doc, DECIMAL_ABSTRACT_NUM_ID, decimalAbstractNumLevels);
}

function ensureNumberingDefinition(
  doc: Docx,
  abstractNumIdValue: number,
  levelsFactory: () => Array<{
    ilvl: number;
    numFmt: "bullet" | "decimal" | string;
    lvlText: string;
  }>,
): number {
  const part = ensureNumberingPart(doc);
  const hasAbstract = part.abstractNums.some(
    (a) =>
      a.attrs.find((x) => x.name.uri === WML_NS && x.name.local === "abstractNumId")?.value ===
      String(abstractNumIdValue),
  );
  if (!hasAbstract) {
    part.abstractNums.push(
      buildAbstractNum({
        abstractNumId: abstractNumIdValue,
        levels: levelsFactory() as never,
      }),
    );
    doc.numberingDirty = true;
  }
  // Find an existing num pointing to this abstractNumId; otherwise add one.
  let chosenId: number | undefined;
  for (const n of part.nums) {
    if (numAbstractRef(n) === abstractNumIdValue) {
      chosenId = readNumId(n);
      if (chosenId !== undefined) break;
    }
  }
  if (chosenId === undefined) {
    const used = new Set<number>();
    for (const n of part.nums) {
      const id = readNumId(n);
      if (id !== undefined) used.add(id);
    }
    let next = 1;
    while (used.has(next)) next++;
    part.nums.push(buildNum(next, abstractNumIdValue));
    doc.numberingDirty = true;
    chosenId = next;
  }
  return chosenId;
}

function ensureNumberingPart(doc: Docx): WmlNumberingPart {
  const existing = numberingPart(doc);
  if (existing) return existing;
  addPart(doc.opc, {
    name: NUMBERING_PART_NAME,
    contentType: WML_CONTENT_TYPES.numbering,
    data: new TextEncoder().encode(EMPTY_NUMBERING_XML),
  });
  const docRels = partRelationships(doc.opc, doc.partName);
  if (relationshipsByType(docRels, WML_RELATIONSHIPS.numbering).length === 0) {
    addRelationship(docRels, { type: WML_RELATIONSHIPS.numbering, target: "numbering.xml" });
  }
  const xml = new TextDecoder("utf-8").decode(
    getPart(doc.opc, NUMBERING_PART_NAME)?.data ?? new Uint8Array(),
  );
  doc.numberingCache = parseNumberingPart(parseXml(xml));
  doc.numberingDirty = true;
  return doc.numberingCache;
}

function ensureStylesPart(doc: Docx): WmlStylesPart {
  const existing = stylesPart(doc);
  if (existing) return existing;
  addPart(doc.opc, {
    name: STYLES_PART_NAME,
    contentType: WML_CONTENT_TYPES.styles,
    data: new TextEncoder().encode(MINIMAL_STYLES_XML),
  });
  const docRels = partRelationships(doc.opc, doc.partName);
  const hasStylesRel = relationshipsByType(docRels, WML_RELATIONSHIPS.styles).length > 0;
  if (!hasStylesRel) {
    addRelationship(docRels, { type: WML_RELATIONSHIPS.styles, target: "styles.xml" });
  }
  const xml = new TextDecoder("utf-8").decode(
    getPart(doc.opc, STYLES_PART_NAME)?.data ?? new Uint8Array(),
  );
  doc.stylesCache = parseStylesPart(parseXml(xml));
  doc.stylesDirty = true;
  return doc.stylesCache;
}

/**
 * Append a table to the body. `rows` is a row-major matrix of strings;
 * each cell becomes a single paragraph with a single run containing the
 * supplied text. Empty cells are allowed.
 */
export function addTable(
  doc: Docx,
  rows: ReadonlyArray<ReadonlyArray<string>>,
  options: BuildTableOptions = {},
): WmlTable {
  const table = buildTextTable(rows, options);
  doc.document.body.blocks.push(table);
  doc.dirty = true;
  return table;
}

/**
 * Search the document body for all matches of `query`.
 *
 * **Match semantics:** matches every occurrence regardless of the
 * regex `g` flag — i.e. the result is what you'd get from "Find All"
 * in Word's UI rather than `String.prototype.match` without `/g`.
 * String queries are treated as plain substrings (no regex
 * interpretation, case-sensitive).
 */
export function findText(doc: Docx, query: string | RegExp): TextMatch[] {
  return wmlFindText(doc.document, query);
}

/**
 * Find every occurrence of `query` across the body and all
 * header/footer/comment/footnote/endnote parts. Returns an array of
 * `{ partName, matches }` entries (the body uses `"/word/document.xml"`).
 *
 * Useful for inspecting a template before calling
 * {@link replaceTextEverywhere}.
 */
export function findTextEverywhere(
  doc: Docx,
  query: string | RegExp,
): Array<{
  partName: string;
  matches: TextMatch[];
}> {
  const out: Array<{ partName: string; matches: TextMatch[] }> = [];
  const bodyMatches = wmlFindText(doc.document, query);
  if (bodyMatches.length > 0) {
    out.push({ partName: doc.partName, matches: bodyMatches });
  }
  // Flush dirty side-parts so the XML walk sees their current state.
  if (doc.commentsDirty && doc.commentsCache) {
    flushComments(doc, doc.commentsCache);
    doc.commentsDirty = false;
    doc.commentsCache = undefined;
  }
  if (doc.footnotesDirty && doc.footnotesCache) {
    flushNotes(doc, doc.footnotesCache, FOOTNOTES_PART_NAME, "footnotes");
    doc.footnotesDirty = false;
    doc.footnotesCache = undefined;
  }
  if (doc.endnotesDirty && doc.endnotesCache) {
    flushNotes(doc, doc.endnotesCache, ENDNOTES_PART_NAME, "endnotes");
    doc.endnotesDirty = false;
    doc.endnotesCache = undefined;
  }
  const docRels = partRelationships(doc.opc, doc.partName);
  for (const rel of allRelationships(docRels)) {
    if (rel.targetMode === "External") continue;
    if (
      rel.type !== WML_RELATIONSHIPS.header &&
      rel.type !== WML_RELATIONSHIPS.footer &&
      rel.type !== WML_RELATIONSHIPS.comments &&
      rel.type !== WML_RELATIONSHIPS.footnotes &&
      rel.type !== WML_RELATIONSHIPS.endnotes
    )
      continue;
    const partName = resolvePartTarget(doc, rel.target);
    const part = getPart(doc.opc, partName);
    if (!part) continue;
    const xml = new TextDecoder("utf-8").decode(part.data);
    const xmlDoc = parseXml(xml);
    const matches: TextMatch[] = [];
    const visit = (el: XmlElement): void => {
      for (const child of el.children) {
        if (!child || child.kind !== "element") continue;
        if (child.name.uri === WML_NS && child.name.local === "p") {
          const para = parseParagraph(child);
          matches.push(
            ...wmlFindText(
              { rootAttrs: [], body: { blocks: [para], extras: [] }, extras: [] },
              query,
            ),
          );
        } else {
          visit(child);
        }
      }
    };
    visit(xmlDoc.root);
    if (matches.length > 0) out.push({ partName, matches });
  }
  return out;
}

/**
 * Replace every occurrence of `query`. Returns the number of replacements
 * performed. Marks the document as dirty so the next `toUint8Array()`
 * re-serializes the body.
 *
 * **Match semantics:** matches every occurrence regardless of the
 * regex `g` flag — equivalent to Word's "Replace All". String queries
 * are matched literally (no regex interpretation). The replacement is
 * applied in a single pass, so a `replacement` callback that returns
 * a string containing the original pattern does NOT loop.
 */
export function replaceText(
  doc: Docx,
  query: string | RegExp,
  replacement: string | ((match: TextMatch) => string),
): number {
  const count = wmlReplaceText(doc.document, query, replacement);
  if (count > 0) doc.dirty = true;
  return count;
}

/**
 * Replace text everywhere it can appear inside the package: body, all
 * headers, footers, footnotes, endnotes and comment bodies. Returns the
 * total count of replacements across every part.
 *
 * Run-spanning matches are handled within each `<w:p>` individually
 * (matches that span across paragraphs are not supported).
 */
export function replaceTextEverywhere(
  doc: Docx,
  query: string | RegExp,
  replacement: string | ((match: TextMatch) => string),
): number {
  // Flush any dirty in-memory parts to package bytes so the XML walk sees
  // the latest state. We invalidate the caches afterwards so the next
  // accessor reloads fresh.
  if (doc.commentsDirty && doc.commentsCache) {
    flushComments(doc, doc.commentsCache);
    doc.commentsDirty = false;
    doc.commentsCache = undefined;
  }
  if (doc.footnotesDirty && doc.footnotesCache) {
    flushNotes(doc, doc.footnotesCache, FOOTNOTES_PART_NAME, "footnotes");
    doc.footnotesDirty = false;
    doc.footnotesCache = undefined;
  }
  if (doc.endnotesDirty && doc.endnotesCache) {
    flushNotes(doc, doc.endnotesCache, ENDNOTES_PART_NAME, "endnotes");
    doc.endnotesDirty = false;
    doc.endnotesCache = undefined;
  }

  let total = replaceText(doc, query, replacement);

  const partsToVisit = new Set<string>();
  const docRels = partRelationships(doc.opc, doc.partName);
  for (const rel of allRelationships(docRels)) {
    if (rel.targetMode === "External") continue;
    if (
      rel.type === WML_RELATIONSHIPS.header ||
      rel.type === WML_RELATIONSHIPS.footer ||
      rel.type === WML_RELATIONSHIPS.comments ||
      rel.type === WML_RELATIONSHIPS.footnotes ||
      rel.type === WML_RELATIONSHIPS.endnotes
    ) {
      partsToVisit.add(resolvePartTarget(doc, rel.target));
    }
  }

  for (const partName of partsToVisit) {
    total += replaceTextInPartXml(doc, partName, query, replacement);
  }
  return total;
}

function resolvePartTarget(doc: Docx, relTarget: string): string {
  if (relTarget.startsWith("/")) return relTarget;
  // Relative to /word/document.xml's folder (/word/).
  return `/word/${relTarget}`;
}

function replaceTextInPartXml(
  doc: Docx,
  partName: string,
  query: string | RegExp,
  replacement: string | ((match: TextMatch) => string),
): number {
  const part = getPart(doc.opc, partName);
  if (!part) return 0;
  const xmlText = new TextDecoder("utf-8").decode(part.data);
  const xmlDoc = parseXml(xmlText);
  let count = 0;
  const visit = (el: XmlElement): void => {
    const children = el.children as XmlNode[];
    for (let i = 0; i < children.length; i++) {
      const child = children[i];
      if (!child || child.kind !== "element") continue;
      if (child.name.uri === WML_NS && child.name.local === "p") {
        const para = parseParagraph(child);
        const n = replaceInParagraph(para, query, replacement);
        if (n > 0) {
          children[i] = paragraphToElement(para);
          count += n;
        }
      } else {
        visit(child);
      }
    }
  };
  visit(xmlDoc.root);
  if (count > 0) {
    part.data = new TextEncoder().encode(serializeXml(xmlDoc));
  }
  return count;
}

/** Visible text of the document. Paragraphs are joined with `\n`. */
export function text(doc: Docx): string {
  return documentText(doc.document);
}

/**
 * Append a section break paragraph. The paragraph terminates the current
 * section; subsequent paragraphs belong to the new section. Page size /
 * margins on the section break override the body trailing sectPr for
 * everything from this point on.
 */
export function appendSectionBreak(
  doc: Docx,
  type: "continuous" | "nextPage" | "evenPage" | "oddPage" | "nextColumn" = "nextPage",
  options: { pageSize?: PageSize; pageMargins?: PageMargins } = {},
): WmlParagraph {
  const sectPrChildren: XmlElement[] = [
    {
      kind: "element",
      name: { uri: WML_NS, local: "type", prefix: "w" },
      attrs: [
        { name: { uri: WML_NS, local: "val", prefix: "w" }, value: type, isNamespaceDecl: false },
      ],
      children: [],
      xmlSpace: "default",
      selfClosing: true,
    },
  ];
  if (options.pageSize) {
    sectPrChildren.push(buildPgSzElement(options.pageSize));
  }
  if (options.pageMargins) {
    sectPrChildren.push(buildPgMarElement(options.pageMargins));
  }
  const sectPr: XmlElement = {
    kind: "element",
    name: { uri: WML_NS, local: "sectPr", prefix: "w" },
    attrs: [],
    children: sectPrChildren,
    xmlSpace: "default",
    selfClosing: false,
  };
  const pPr: XmlElement = {
    kind: "element",
    name: { uri: WML_NS, local: "pPr", prefix: "w" },
    attrs: [],
    children: [sectPr],
    xmlSpace: "default",
    selfClosing: false,
  };
  const paragraph: WmlParagraph = {
    kind: "paragraph",
    pPr,
    children: [],
    extras: [],
  };
  doc.document.body.blocks.push(paragraph);
  doc.dirty = true;
  return paragraph;
}

/**
 * Insert a new text paragraph at the given paragraph-relative index.
 * `index` counts only paragraphs (so `0` means "before the first
 * paragraph"); other body block kinds are skipped past. If `index` is
 * past the end of the existing paragraphs, the new paragraph is
 * appended.
 */
export function insertParagraphAt(
  doc: Docx,
  index: number,
  text: string,
  options: AppendParagraphOptions = {},
): WmlParagraph {
  let count = 0;
  let blockInsertAt = doc.document.body.blocks.length;
  for (let i = 0; i < doc.document.body.blocks.length; i++) {
    const b = doc.document.body.blocks[i];
    if (b?.kind === "paragraph") {
      if (count === index) {
        blockInsertAt = i;
        break;
      }
      count++;
    }
  }
  // Build the paragraph using the same logic as appendParagraph by
  // calling it temporarily, then moving the newly-pushed paragraph into
  // place.
  const para = appendParagraph(doc, text, options);
  const lastIdx = doc.document.body.blocks.length - 1;
  if (lastIdx !== blockInsertAt) {
    doc.document.body.blocks.splice(lastIdx, 1);
    doc.document.body.blocks.splice(blockInsertAt, 0, para);
  }
  return para;
}

/** Remove the paragraph at `index` from the body. Returns true on success. */
export function removeParagraph(doc: Docx, index: number): boolean {
  let count = 0;
  for (let i = 0; i < doc.document.body.blocks.length; i++) {
    const b = doc.document.body.blocks[i];
    if (b?.kind === "paragraph") {
      if (count === index) {
        doc.document.body.blocks.splice(i, 1);
        doc.dirty = true;
        return true;
      }
      count++;
    }
  }
  return false;
}

/**
 * Remove the `index`-th `<w:tbl>` block from the body (0-based, counted
 * across only tables — paragraphs in between are not counted). Returns
 * true on success.
 */
export function removeTable(doc: Docx, index: number): boolean {
  let count = 0;
  for (let i = 0; i < doc.document.body.blocks.length; i++) {
    const b = doc.document.body.blocks[i];
    if (b?.kind === "table") {
      if (count === index) {
        doc.document.body.blocks.splice(i, 1);
        doc.dirty = true;
        return true;
      }
      count++;
    }
  }
  return false;
}

/**
 * Remove every `<w:tbl>` block from the body. Returns the number of
 * tables removed. Paragraphs between or around the tables are kept
 * untouched.
 */
export function removeAllTables(doc: Docx): number {
  let removed = 0;
  doc.document.body.blocks = doc.document.body.blocks.filter((b) => {
    if (b.kind === "table") {
      removed++;
      return false;
    }
    return true;
  });
  if (removed > 0) doc.dirty = true;
  return removed;
}

/**
 * Replace the `index`-th `<w:tbl>` in the body with the paragraphs that
 * appear inside its cells (row-major: row 0 first, then row 1, …; within
 * each row left-to-right). The original cell paragraph objects are
 * spliced in directly — table properties (`<w:tblPr>`), row formatting,
 * and cell formatting are dropped. Useful when an imported template
 * uses a table only for visual layout and you want to flatten it back to
 * plain prose.
 *
 * Returns the paragraphs that ended up in the body, or `undefined` if no
 * table exists at the given index.
 */
export function unwrapTable(doc: Docx, index: number): WmlParagraph[] | undefined {
  let count = 0;
  for (let i = 0; i < doc.document.body.blocks.length; i++) {
    const b = doc.document.body.blocks[i];
    if (b?.kind !== "table") continue;
    if (count !== index) {
      count++;
      continue;
    }
    const paragraphs: WmlParagraph[] = [];
    for (const row of b.rows) {
      for (const cell of row.cells) {
        for (const p of cell.paragraphs) paragraphs.push(p);
      }
    }
    doc.document.body.blocks.splice(i, 1, ...paragraphs);
    doc.dirty = true;
    return paragraphs;
  }
  return undefined;
}

/** Remove all body blocks (paragraphs and tables). The body sectPr is kept. */
export function clearBody(doc: Docx): void {
  doc.document.body.blocks.length = 0;
  doc.dirty = true;
}

/**
 * Remove every image part (`/word/media/*`), drop matching image
 * relationships, and strip `<w:drawing>` runs from the body. Returns
 * the number of media parts removed.
 */
export function removeAllImages(doc: Docx): number {
  const docRels = partRelationships(doc.opc, doc.partName);
  let removed = 0;
  for (const rel of relationshipsByType(docRels, WML_RELATIONSHIPS.image)) {
    const partName = resolvePartTarget(doc, rel.target);
    if (rel.targetMode !== "External" && removePart(doc.opc, partName)) removed++;
    removeRelationship(docRels, rel.id);
  }
  const hasDrawing = (el: XmlElement): boolean => {
    if (el.name.uri === WML_NS && el.name.local === "drawing") return true;
    for (const c of el.children) {
      if (c.kind === "element" && hasDrawing(c)) return true;
    }
    return false;
  };
  const stripDrawings = (p: WmlParagraph): void => {
    p.children = p.children.filter((child) => {
      if (child.kind === "raw") {
        if (child.node.name.uri === WML_NS && child.node.name.local === "drawing") return false;
        if (child.node.name.local === "r" && hasDrawing(child.node)) return false;
      }
      if (child.kind === "run") {
        child.pieces = child.pieces.filter((piece) => piece.kind !== "drawing");
      }
      return true;
    });
  };
  const walk = (blocks: WmlBlock[]): void => {
    for (const b of blocks) {
      if (b.kind === "paragraph") stripDrawings(b);
      else if (b.kind === "table") {
        for (const row of b.rows) {
          for (const cell of row.cells) {
            for (const p of cell.paragraphs) stripDrawings(p);
          }
        }
      }
    }
  };
  walk(doc.document.body.blocks);
  if (removed > 0) doc.dirty = true;
  return removed;
}

/**
 * Strip every `<w:hyperlink>` element from the body, keeping the inner
 * runs in place (so the linked text remains as plain text). External
 * relationships (`http://...` targets) referenced by the dropped
 * hyperlinks are also removed from `document.xml.rels`. Internal anchor
 * hyperlinks have no rel, so only the wrapper is removed.
 *
 * Returns the number of `<w:hyperlink>` wrappers that were unwrapped.
 */
export function removeAllHyperlinks(doc: Docx): number {
  const docRels = partRelationships(doc.opc, doc.partName);
  const usedRelIds = new Set<string>();
  let unwrapped = 0;

  for (const block of doc.document.body.blocks) {
    if (block.kind !== "paragraph") continue;
    const out: typeof block.children = [];
    for (const child of block.children) {
      if (
        child.kind === "raw" &&
        child.node.name.uri === WML_NS &&
        child.node.name.local === "hyperlink"
      ) {
        // Record the r:id (if any) so we can drop the matching relationship.
        for (const a of child.node.attrs) {
          if (a.name.local === "id" && a.name.prefix === "r") {
            usedRelIds.add(a.value);
          }
        }
        // Splice the wrapper's element children in as raw inlines. They are
        // typically `<w:r>` runs but can also be tracked-change wrappers.
        for (const inner of child.node.children) {
          if (inner.kind === "element") {
            out.push({ kind: "raw", node: inner });
          }
        }
        unwrapped++;
        continue;
      }
      out.push(child);
    }
    block.children = out;
  }

  // Drop hyperlink-type rels whose ids we just orphaned.
  for (const rel of relationshipsByType(docRels, WML_RELATIONSHIPS.hyperlink)) {
    if (usedRelIds.has(rel.id)) {
      removeRelationship(docRels, rel.id);
    }
  }

  if (unwrapped > 0) doc.dirty = true;
  return unwrapped;
}

/**
 * Walk every external hyperlink relationship on `document.xml.rels` and
 * rewrite the `Target` whenever `predicate(target)` returns a new URL.
 * Pass a string for `predicate` to match an exact target.
 *
 * Returns the number of rels whose target was rewritten. Internal anchor
 * links (`w:anchor=...`) have no rel and are not touched.
 *
 * Useful for swapping the hostname on every template link in one pass,
 * eg. `setHyperlinkUrl(doc, "https://stage.example.com", "https://example.com")`.
 */
export function setHyperlinkUrl(
  doc: Docx,
  matcher: string | ((target: string) => string | null | undefined),
  replacement?: string,
): number {
  const docRels = partRelationships(doc.opc, doc.partName);
  let changed = 0;
  for (const rel of relationshipsByType(docRels, WML_RELATIONSHIPS.hyperlink)) {
    if (rel.targetMode !== "External") continue;
    let next: string | undefined;
    if (typeof matcher === "string") {
      if (rel.target === matcher && replacement !== undefined) next = replacement;
    } else {
      const candidate = matcher(rel.target);
      if (typeof candidate === "string" && candidate !== rel.target) next = candidate;
    }
    if (next !== undefined) {
      rel.target = next;
      changed++;
    }
  }
  if (changed > 0) doc.dirty = true;
  return changed;
}

/**
 * List every external hyperlink relationship currently on
 * `document.xml.rels`. Each entry includes the rel id and the target URL.
 * Use this to audit which URLs a template references before deciding
 * what to rewrite with {@link setHyperlinkUrl}.
 */
export function externalHyperlinks(doc: Docx): Array<{ relId: string; target: string }> {
  const docRels = partRelationships(doc.opc, doc.partName);
  const out: Array<{ relId: string; target: string }> = [];
  for (const rel of relationshipsByType(docRels, WML_RELATIONSHIPS.hyperlink)) {
    if (rel.targetMode !== "External") continue;
    out.push({ relId: rel.id, target: rel.target });
  }
  return out;
}

/** Remove every bookmark from the body. Returns the count removed. */
export function removeAllBookmarks(doc: Docx): number {
  const names = bookmarks(doc).map((b) => b.name);
  let removed = 0;
  for (const name of names) {
    if (removeBookmark(doc, name)) removed++;
  }
  return removed;
}

/**
 * Remove every header part and its references from the body's sectPr.
 * Returns the number of header parts removed.
 */
export function removeAllHeaders(doc: Docx): number {
  return removeAllHeaderFooterParts(doc, WML_RELATIONSHIPS.header, "headerReference");
}

/** Same shape as {@link removeAllHeaders} but for footer parts. */
export function removeAllFooters(doc: Docx): number {
  return removeAllHeaderFooterParts(doc, WML_RELATIONSHIPS.footer, "footerReference");
}

/**
 * Remove every user footnote AND every `<w:footnoteReference>` run from
 * the body. Keeps the standard separator / continuationSeparator entries
 * intact. Returns the number of user footnotes removed.
 */
export function removeAllFootnotes(doc: Docx): number {
  return removeAllNotes(doc, "footnote");
}

/** Same shape as {@link removeAllFootnotes} but for endnotes. */
export function removeAllEndnotes(doc: Docx): number {
  return removeAllNotes(doc, "endnote");
}

function removeAllNotes(doc: Docx, kind: "footnote" | "endnote"): number {
  const part = kind === "footnote" ? footnotesPart(doc) : endnotesPart(doc);
  if (!part) return 0;
  let removed = 0;
  const survivors: typeof part.footnotes = [];
  for (const note of part.footnotes) {
    const typeAttr = note.attrs.find((a) => a.name.uri === WML_NS && a.name.local === "type");
    if (
      typeAttr &&
      (typeAttr.value === "separator" || typeAttr.value === "continuationSeparator")
    ) {
      survivors.push(note);
    } else {
      removed++;
    }
  }
  part.footnotes.length = 0;
  for (const s of survivors) part.footnotes.push(s);
  if (kind === "footnote") doc.footnotesDirty = true;
  else doc.endnotesDirty = true;
  // Strip <w:footnoteReference> / <w:endnoteReference> runs from body
  // paragraphs. Match both shapes:
  //   - raw {kind:"raw", node:<w:r><w:footnoteReference/></w:r>} (the
  //     output addFootnote produces).
  //   - parsed WmlRun whose only piece is a raw <w:footnoteReference/>
  //     (the shape parseRun produces after a save+open round-trip).
  const refLocal = kind === "footnote" ? "footnoteReference" : "endnoteReference";
  const stripFromParagraph = (p: WmlParagraph): void => {
    p.children = p.children.filter((child) => {
      if (child.kind === "raw") {
        if (child.node.name.uri !== WML_NS) return true;
        if (child.node.name.local !== "r") return true;
        return !child.node.children.some(
          (c) => c.kind === "element" && c.name.uri === WML_NS && c.name.local === refLocal,
        );
      }
      if (child.kind === "run") {
        const piecesOnlyRef =
          child.pieces.length > 0 &&
          child.pieces.every(
            (piece) =>
              piece.kind === "raw" &&
              piece.node.name.uri === WML_NS &&
              piece.node.name.local === refLocal,
          );
        if (piecesOnlyRef) return false;
      }
      return true;
    });
  };
  const walk = (blocks: WmlBlock[]): void => {
    for (const b of blocks) {
      if (b.kind === "paragraph") stripFromParagraph(b);
      else if (b.kind === "table") {
        for (const row of b.rows) {
          for (const cell of row.cells) {
            for (const p of cell.paragraphs) stripFromParagraph(p);
          }
        }
      }
    }
  };
  walk(doc.document.body.blocks);
  if (removed > 0) doc.dirty = true;
  return removed;
}

function removeAllHeaderFooterParts(doc: Docx, relType: string, refLocal: string): number {
  let removed = 0;
  const docRels = partRelationships(doc.opc, doc.partName);
  for (const rel of relationshipsByType(docRels, relType)) {
    const partName = resolvePartTarget(doc, rel.target);
    if (removePart(doc.opc, partName)) removed++;
    removeRelationship(docRels, rel.id);
  }
  // Strip *Reference children from the body's sectPr in place.
  const sectPr = doc.document.body.sectPr;
  if (sectPr) {
    const arr = sectPr.children as XmlElement[];
    let writeIdx = 0;
    for (let i = 0; i < arr.length; i++) {
      const c = arr[i];
      if (!c) continue;
      if (c.kind === "element" && c.name.uri === WML_NS && c.name.local === refLocal) continue;
      arr[writeIdx++] = c;
    }
    arr.length = writeIdx;
  }
  if (removed > 0) doc.dirty = true;
  return removed;
}

/**
 * Remove every comment from `comments.xml` AND every comment range
 * marker and reference run from `document.xml`. Returns the number of
 * comments removed.
 */
export function removeAllComments(doc: Docx): number {
  const part = commentsPart(doc);
  let removed = 0;
  if (part) {
    removed = part.comments.length;
    part.comments.length = 0;
    doc.commentsDirty = true;
  }
  // Strip commentRangeStart / commentRangeEnd / commentReference from the
  // body. The reference may live as a raw inline (immediately after
  // addComment, before any round-trip) OR as a WmlRun whose only piece is a
  // raw <w:commentReference/> child (the shape produced by parseRun on
  // re-open). Match both shapes.
  const stripFromParagraph = (p: WmlParagraph): void => {
    p.children = p.children.filter((child) => {
      if (child.kind === "raw") {
        const local = child.node.name.local;
        if (child.node.name.uri !== WML_NS) return true;
        if (local === "commentRangeStart" || local === "commentRangeEnd") return false;
        if (
          local === "r" &&
          child.node.children.some(
            (c) =>
              c.kind === "element" && c.name.uri === WML_NS && c.name.local === "commentReference",
          )
        ) {
          return false;
        }
        return true;
      }
      if (child.kind === "run") {
        const piecesOnlyCommentRef =
          child.pieces.length > 0 &&
          child.pieces.every(
            (piece) =>
              piece.kind === "raw" &&
              piece.node.name.uri === WML_NS &&
              piece.node.name.local === "commentReference",
          );
        if (piecesOnlyCommentRef) return false;
      }
      return true;
    });
  };
  const walk = (blocks: WmlBlock[]): void => {
    for (const b of blocks) {
      if (b.kind === "paragraph") stripFromParagraph(b);
      else if (b.kind === "table") {
        for (const row of b.rows) {
          for (const cell of row.cells) {
            for (const p of cell.paragraphs) stripFromParagraph(p);
          }
        }
      }
    }
  };
  walk(doc.document.body.blocks);
  if (removed > 0) doc.dirty = true;
  return removed;
}

/**
 * Append a paragraph whose only content is a page break. Equivalent to
 * Ctrl+Enter in Word.
 */
export function appendPageBreak(doc: Docx): WmlParagraph {
  const piece: WmlRunPiece = { kind: "break", breakType: "page" };
  const run: WmlRun = { kind: "run", pieces: [piece], extras: [] };
  const paragraph: WmlParagraph = { kind: "paragraph", children: [run], extras: [] };
  doc.document.body.blocks.push(paragraph);
  doc.dirty = true;
  return paragraph;
}

/**
 * Append a `<w:br>` to an existing paragraph as a soft break.
 *
 * `kind`:
 * - `"line"` (default) — soft line break (Shift+Enter in Word).
 * - `"page"` — page break inside the paragraph; Word splits the page
 *   here without inserting a new paragraph above.
 * - `"column"` — column break (only meaningful in multi-column sections).
 */
export function appendLineBreak(
  doc: Docx,
  paragraph: WmlParagraph,
  kind: "line" | "page" | "column" = "line",
): WmlRun {
  const piece: WmlRunPiece =
    kind === "line"
      ? { kind: "break" }
      : { kind: "break", breakType: kind === "page" ? "page" : "column" };
  const run: WmlRun = { kind: "run", pieces: [piece], extras: [] };
  paragraph.children.push(run);
  doc.dirty = true;
  return run;
}

/**
 * Add a named bookmark covering a paragraph. Returns the assigned numeric
 * bookmark id; the same name must not be reused without removing the
 * existing bookmark first (Word will deduplicate silently otherwise).
 */
export function addBookmark(doc: Docx, name: string, paragraph: WmlParagraph): number {
  const id = allocateBookmarkId(doc);
  const start: XmlElement = {
    kind: "element",
    name: { uri: WML_NS, local: "bookmarkStart", prefix: "w" },
    attrs: [
      {
        name: { uri: WML_NS, local: "id", prefix: "w" },
        value: String(id),
        isNamespaceDecl: false,
      },
      { name: { uri: WML_NS, local: "name", prefix: "w" }, value: name, isNamespaceDecl: false },
    ],
    children: [],
    xmlSpace: "default",
    selfClosing: true,
  };
  const end: XmlElement = {
    kind: "element",
    name: { uri: WML_NS, local: "bookmarkEnd", prefix: "w" },
    attrs: [
      {
        name: { uri: WML_NS, local: "id", prefix: "w" },
        value: String(id),
        isNamespaceDecl: false,
      },
    ],
    children: [],
    xmlSpace: "default",
    selfClosing: true,
  };
  paragraph.children = [
    { kind: "raw", node: start },
    ...paragraph.children,
    { kind: "raw", node: end },
  ];
  doc.dirty = true;
  return id;
}

/**
 * Append a paragraph with a hyperlink that points at an internal bookmark.
 */
export function addInternalHyperlink(
  doc: Docx,
  bookmarkName: string,
  text: string,
  options: { tooltip?: string } = {},
): WmlParagraph {
  const hyperlinkEl = buildHyperlink({
    anchor: bookmarkName,
    runs: [buildHyperlinkRun(text)],
    ...(options.tooltip !== undefined ? { tooltip: options.tooltip } : {}),
  });
  const paragraph: WmlParagraph = {
    kind: "paragraph",
    children: [{ kind: "raw", node: hyperlinkEl }],
    extras: [],
  };
  doc.document.body.blocks.push(paragraph);
  doc.dirty = true;
  return paragraph;
}

function allocateBookmarkId(doc: Docx): number {
  // Bookmarks are numbered uniquely per document; we scan existing IDs to
  // avoid colliding with anything in the document.
  let max = -1;
  const scan = (el: XmlElement): void => {
    if (
      el.name.uri === WML_NS &&
      (el.name.local === "bookmarkStart" || el.name.local === "bookmarkEnd")
    ) {
      const idAttr = el.attrs.find((a) => a.name.uri === WML_NS && a.name.local === "id");
      if (idAttr) {
        const n = Number.parseInt(idAttr.value, 10);
        if (Number.isFinite(n) && n > max) max = n;
      }
    }
    for (const c of el.children) {
      if (c.kind === "element") scan(c);
    }
  };
  // Walk all inline nodes in body to find existing bookmark ids.
  for (const block of doc.document.body.blocks) {
    if (block.kind === "paragraph") {
      for (const child of block.children) {
        if (child.kind === "raw") scan(child.node);
      }
    } else if (block.kind === "raw") {
      scan(block.node);
    }
  }
  return max + 1;
}

/**
 * Append a paragraph containing a single external hyperlink. Creates an
 * external relationship (TargetMode=External) for the URL and wraps a
 * styled run inside `<w:hyperlink>`.
 */
export function addHyperlink(
  doc: Docx,
  url: string,
  text: string,
  options: { tooltip?: string } = {},
): WmlParagraph {
  const docRels = partRelationships(doc.opc, doc.partName);
  const rel = addRelationship(docRels, {
    type: WML_RELATIONSHIPS.hyperlink,
    target: url,
    targetMode: "External",
  });
  const hyperlinkEl = buildHyperlink({
    relId: rel.id,
    runs: [buildHyperlinkRun(text)],
    ...(options.tooltip !== undefined ? { tooltip: options.tooltip } : {}),
  });
  const paragraph: WmlParagraph = {
    kind: "paragraph",
    children: [{ kind: "raw", node: hyperlinkEl }],
    extras: [],
  };
  doc.document.body.blocks.push(paragraph);
  doc.dirty = true;
  return paragraph;
}

/**
 * Accept every `<w:ins>` and `<w:del>` revision in the document.
 * Insertions are kept and unwrapped; deletions are dropped. Returns the
 * number of revisions resolved.
 */
export function acceptAllRevisions(doc: Docx): number {
  const n = wmlAcceptAllRevisions(doc.document);
  if (n > 0) doc.dirty = true;
  return n;
}

/**
 * Reject every `<w:ins>` and `<w:del>` revision in the document.
 * Insertions are dropped; deletions are kept and unwrapped. Returns the
 * number of revisions resolved.
 */
export function rejectAllRevisions(doc: Docx): number {
  const n = wmlRejectAllRevisions(doc.document);
  if (n > 0) doc.dirty = true;
  return n;
}

/**
 * Shortcut for `appendParagraph(text, { style: "Heading{level}" })`. Note
 * that the referenced `HeadingN` style must be present in `styles.xml`
 * for Word to render it as a heading — call {@link ensureHeadingStyles}
 * once during setup if you want defaults to be defined automatically.
 */
export function appendHeading(doc: Docx, text: string, level = 1): WmlParagraph {
  if (level < 1 || level > 9 || !Number.isInteger(level)) {
    throw new Error(`Heading level must be an integer in 1..9, got ${level}`);
  }
  return appendParagraph(doc, text, { style: `Heading${level}` });
}

/**
 * Ensure default `Heading1`..`Heading{maxLevel}` styles exist in
 * `styles.xml`. Each missing style is added with sensible defaults
 * (smaller / lighter as the level increases). Existing same-id styles
 * are left untouched.
 */
export function ensureHeadingStyles(doc: Docx, maxLevel = 3): void {
  const sizes = [40, 32, 28, 24, 22, 22, 22, 22, 22];
  const colors = [
    "1F497D",
    "1F497D",
    "4F81BD",
    "4F81BD",
    "8DB3E2",
    "8DB3E2",
    "8DB3E2",
    "8DB3E2",
    "8DB3E2",
  ];
  const part = stylesPart(doc);
  for (let lvl = 1; lvl <= Math.min(maxLevel, 9); lvl++) {
    const id = `Heading${lvl}`;
    if (part?.styles.some((s) => s.attrs.some((a) => a.name.local === "styleId" && a.value === id)))
      continue;
    addStyle(doc, {
      type: "paragraph",
      styleId: id,
      name: `heading ${lvl}`,
      basedOn: "Normal",
      next: "Normal",
      qFormat: true,
      uiPriority: 9,
      bold: lvl <= 2,
      fontSizeHalfPoints: sizes[lvl - 1] ?? 24,
      color: colors[lvl - 1] ?? "1F497D",
    });
  }
}

/**
 * Extract a document outline — every paragraph whose pStyle matches
 * `Heading{level}`. Each entry includes the paragraph reference, the
 * heading level (1-9), and the visible text.
 */
export function outline(
  doc: Docx,
): Array<{ paragraph: WmlParagraph; level: number; text: string }> {
  const out: Array<{ paragraph: WmlParagraph; level: number; text: string }> = [];
  for (const block of doc.document.body.blocks) {
    if (block.kind !== "paragraph") continue;
    const style = block.pPr?.children.find(
      (c) => c.kind === "element" && c.name.uri === WML_NS && c.name.local === "pStyle",
    );
    if (style?.kind !== "element") continue;
    const styleId = style.attrs.find((a) => a.name.uri === WML_NS && a.name.local === "val")?.value;
    if (!styleId) continue;
    const m = /^Heading(\d)$/.exec(styleId);
    if (!m) continue;
    const level = Number.parseInt(m[1] ?? "0", 10);
    if (!Number.isFinite(level) || level < 1 || level > 9) continue;
    out.push({ paragraph: block, level, text: documentTextOfParagraph(block) });
  }
  return out;
}

/** Append a paragraph containing a single text run to the body. */
export function appendParagraph(
  doc: Docx,
  text: string,
  options: AppendParagraphOptions = {},
): WmlParagraph {
  const piece: WmlRunPiece = {
    kind: "text",
    value: text,
    preserveSpace: /^\s|\s$/.test(text),
  };
  const run: WmlRun =
    options.bold || options.italic
      ? {
          kind: "run",
          rPr: {
            kind: "element",
            name: { uri: WML_NS, local: "rPr", prefix: "w" },
            attrs: [],
            children: [
              ...(options.bold ? [wmlEmptyEl("b")] : []),
              ...(options.italic ? [wmlEmptyEl("i")] : []),
            ],
            xmlSpace: "default",
            selfClosing: false,
          },
          pieces: [piece],
          extras: [],
        }
      : {
          kind: "run",
          pieces: [piece],
          extras: [],
        };
  const paragraph: WmlParagraph = options.style
    ? {
        kind: "paragraph",
        pPr: {
          kind: "element",
          name: { uri: WML_NS, local: "pPr", prefix: "w" },
          attrs: [],
          children: [
            {
              kind: "element",
              name: { uri: WML_NS, local: "pStyle", prefix: "w" },
              attrs: [
                {
                  name: { uri: WML_NS, local: "val", prefix: "w" },
                  value: options.style,
                  isNamespaceDecl: false,
                },
              ],
              children: [],
              xmlSpace: "default",
              selfClosing: true,
            },
          ],
          xmlSpace: "default",
          selfClosing: false,
        },
        children: [run],
        extras: [],
      }
    : {
        kind: "paragraph",
        children: [run],
        extras: [],
      };
  doc.document.body.blocks.push(paragraph);
  doc.dirty = true;
  return paragraph;
}

/**
 * Insert an image into the document. The image bytes become a new media
 * part under `/word/media/`, a relationship is added from the main
 * document, and a `<w:drawing>`-bearing run is appended to (or returned
 * for placement in) the body.
 *
 * Use `addImageRun` if you want the constructed run without it being
 * appended automatically.
 */
export function addImage(doc: Docx, bytes: Uint8Array, options: AddImageOptions): WmlParagraph {
  const run = addImageRun(doc, bytes, options);
  const paragraph: WmlParagraph = {
    kind: "paragraph",
    children: [run],
    extras: [],
  };
  doc.document.body.blocks.push(paragraph);
  doc.dirty = true;
  return paragraph;
}

/**
 * Insert an image into an existing paragraph (as the final inline). The
 * image part is added the same way {@link addImage} does, but the
 * paragraph already exists — useful when mixing text and images on one
 * line.
 */
export function insertImageInto(
  doc: Docx,
  paragraph: WmlParagraph,
  bytes: Uint8Array,
  options: AddImageOptions,
): WmlRun {
  const run = addImageRun(doc, bytes, options);
  paragraph.children.push(run);
  doc.dirty = true;
  return run;
}

/**
 * Like {@link addImage} but returns the constructed run without inserting
 * it. Callers can place the run wherever they want (e.g. inside an
 * existing paragraph).
 */
export function addImageRun(doc: Docx, bytes: Uint8Array, options: AddImageOptions): WmlRun {
  const contentType = options.contentType ?? sniffImageContentType(bytes);
  if (!contentType) {
    throw new Error(
      "addImage: could not detect image content type from bytes; pass options.contentType explicitly",
    );
  }
  const ext = extensionForImageContentType(contentType);
  const partName = allocateImagePartName(doc, ext);
  addPart(doc.opc, { name: partName, contentType, data: bytes });
  setContentTypeDefault(doc.opc.contentTypes, ext, contentType);

  const docRels = documentRelationships(doc);
  const rel = addRelationship(docRels, {
    type: WML_RELATIONSHIPS.image,
    target: relativeMediaTarget(doc, partName),
  });

  const docPrId = allocateDocPrId(doc);
  const drawing = buildInlineDrawing({
    relId: rel.id,
    widthEmu: options.widthEmu,
    heightEmu: options.heightEmu,
    docPrId,
    ...(options.name !== undefined ? { name: options.name } : {}),
    ...(options.altText !== undefined ? { altText: options.altText } : {}),
  });
  const piece: WmlRunPiece = { kind: "drawing", node: drawing };
  doc.dirty = true;
  return { kind: "run", pieces: [piece], extras: [] };
}

function documentRelationships(doc: Docx): RelationshipSet {
  const set = partRelationships(doc.opc, doc.partName);
  return set;
}

function allocateImagePartName(doc: Docx, extension: string): string {
  let n = 1;
  // Find lowest unused index.
  while (hasPart(doc.opc, `/word/media/image${n}.${extension}`)) {
    n++;
  }
  return `/word/media/image${n}.${extension}`;
}

function relativeMediaTarget(doc: Docx, partName: string): string {
  // /word/document.xml relates targets relative to its own folder.
  // /word/media/imageN.ext  →  media/imageN.ext
  if (partName.startsWith("/word/")) return partName.slice("/word/".length);
  return partName.startsWith("/") ? partName.slice(1) : partName;
}

function allocateDocPrId(doc: Docx): number {
  // docPr ids must be unique across the document. We bump a counter that
  // starts from the current max we can see in pass-through drawing nodes
  // and from the relationships set length as a coarse upper bound.
  return Math.max(1, allRelationships(documentRelationships(doc)).length + 1);
}

/**
 * Convenience access to the document title (`dc:title`). Reading returns
 * the current value; writing replaces it via {@link setCoreProperties}.
 */
export function title(doc: Docx): string | undefined {
  return coreProperties(doc).title;
}

export function setTitle(doc: Docx, value: string | undefined): void {
  if (value !== undefined) setCoreProperties(doc, { title: value });
}

/**
 * Convenience access to the document author / creator (`dc:creator`).
 */
export function author(doc: Docx): string | undefined {
  return coreProperties(doc).creator;
}

export function setAuthor(doc: Docx, value: string | undefined): void {
  if (value !== undefined) setCoreProperties(doc, { creator: value });
}

/**
 * Enumerate every `<w:instrText>` field instruction found in the body.
 * Useful for inspecting what fields a template carries before deciding
 * how to fill them.
 */
export function fields(doc: Docx): Array<{ paragraph: WmlParagraph; instruction: string }> {
  const out: Array<{ paragraph: WmlParagraph; instruction: string }> = [];
  const visit = (p: WmlParagraph): void => {
    for (const child of p.children) {
      if (child.kind === "run") {
        for (const piece of child.pieces) {
          if (piece.kind === "instrText") {
            out.push({ paragraph: p, instruction: piece.value });
          }
        }
      } else if (child.kind === "raw") {
        const walk = (el: XmlElement): void => {
          if (el.name.uri === WML_NS && el.name.local === "instrText") {
            let acc = "";
            for (const c of el.children) {
              if (c.kind === "text") acc += c.value;
              else if (c.kind === "cdata") acc += c.value;
            }
            out.push({ paragraph: p, instruction: acc });
            return;
          }
          for (const c of el.children) if (c.kind === "element") walk(c);
        };
        walk(child.node);
      }
    }
  };
  for (const block of doc.document.body.blocks) {
    if (block.kind === "paragraph") visit(block);
    else if (block.kind === "table") {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          for (const p of cell.paragraphs) visit(p);
        }
      }
    }
  }
  return out;
}

/**
 * Read the package's core document properties (title, creator, subject,
 * etc.), parsed from `docProps/core.xml`. Returns an empty record if the
 * part is absent.
 */
export function coreProperties(doc: Docx): DocumentCoreProperties {
  const part = getPart(doc.opc, CORE_PROPERTIES_PART_NAME);
  if (!part) return {};
  const xml = new TextDecoder("utf-8").decode(part.data);
  return parseCoreProperties(parseXml(xml));
}

/**
 * Read the package's application-level document properties (Application,
 * AppVersion, Pages, Words, Company, Manager, …) from `docProps/app.xml`.
 * Returns an empty record if the part is absent.
 */
export function appProperties(doc: Docx): DocumentAppProperties {
  const part = getPart(doc.opc, APP_PROPERTIES_PART_NAME);
  if (!part) return {};
  const xml = new TextDecoder("utf-8").decode(part.data);
  return parseAppProperties(parseXml(xml));
}

/**
 * Write application-level document properties. Creates `docProps/app.xml`
 * and the package-level relationship on first use; subsequent calls
 * merge into existing properties.
 */
export function setAppProperties(doc: Docx, props: DocumentAppProperties): void {
  if (!hasPart(doc.opc, APP_PROPERTIES_PART_NAME)) {
    addPart(doc.opc, {
      name: APP_PROPERTIES_PART_NAME,
      contentType: APP_PROPERTIES_CONTENT_TYPE,
      data: new TextEncoder().encode(EMPTY_APP_PROPERTIES_XML),
    });
    const pkgRels = packageRelationships(doc.opc);
    if (relationshipsByType(pkgRels, APP_PROPERTIES_REL_TYPE).length === 0) {
      addRelationship(pkgRels, { type: APP_PROPERTIES_REL_TYPE, target: "docProps/app.xml" });
    }
  }
  const existing = appProperties(doc);
  const merged: DocumentAppProperties = { ...existing, ...props };
  const xml = serializeXml(writeAppProperties(merged));
  const part = getPart(doc.opc, APP_PROPERTIES_PART_NAME);
  if (part) part.data = new TextEncoder().encode(xml);
}

/**
 * Write the package's core document properties. Creates
 * `docProps/core.xml` and the package-level relationship on first use.
 * Subsequent calls merge into existing properties.
 */
export function setCoreProperties(doc: Docx, props: DocumentCoreProperties): void {
  if (!hasPart(doc.opc, CORE_PROPERTIES_PART_NAME)) {
    addPart(doc.opc, {
      name: CORE_PROPERTIES_PART_NAME,
      contentType: CORE_PROPERTIES_CONTENT_TYPE,
      data: new TextEncoder().encode(EMPTY_CORE_PROPERTIES_XML),
    });
    const pkgRels = packageRelationships(doc.opc);
    if (relationshipsByType(pkgRels, CORE_PROPERTIES_REL_TYPE).length === 0) {
      addRelationship(pkgRels, { type: CORE_PROPERTIES_REL_TYPE, target: "docProps/core.xml" });
    }
  }
  const existing = coreProperties(doc);
  const merged: DocumentCoreProperties = { ...existing, ...props };
  const xml = serializeXml(writeCoreProperties(merged));
  const part = getPart(doc.opc, CORE_PROPERTIES_PART_NAME);
  if (part) part.data = new TextEncoder().encode(xml);
}

/** Parsed `word/footnotes.xml` AST, or `undefined` if absent. */
export function footnotesPart(doc: Docx): WmlFootnotesPart | undefined {
  if (doc.footnotesCache) return doc.footnotesCache;
  const part = getPart(doc.opc, FOOTNOTES_PART_NAME);
  if (!part) return undefined;
  const xml = new TextDecoder("utf-8").decode(part.data);
  doc.footnotesCache = parseFootnotesPart(parseXml(xml));
  return doc.footnotesCache;
}

/** Parsed `word/endnotes.xml` AST, or `undefined` if absent. */
export function endnotesPart(doc: Docx): WmlFootnotesPart | undefined {
  if (doc.endnotesCache) return doc.endnotesCache;
  const part = getPart(doc.opc, ENDNOTES_PART_NAME);
  if (!part) return undefined;
  const xml = new TextDecoder("utf-8").decode(part.data);
  doc.endnotesCache = parseFootnotesPart(parseXml(xml));
  return doc.endnotesCache;
}

/**
 * Append a footnote with the given text. Creates `footnotes.xml` (with
 * the standard separator/continuationSeparator entries) on first use,
 * registers the relationship from `word/document.xml`, and inserts a
 * `<w:footnoteReference>` run at the end of the target paragraph.
 *
 * Returns the assigned footnote id (>=1).
 */
export function addFootnote(doc: Docx, paragraph: WmlParagraph, text: string): number {
  return addNoteImpl(doc, paragraph, text, "footnote");
}

/** Same shape as {@link addFootnote} but for endnotes. */
export function addEndnote(doc: Docx, paragraph: WmlParagraph, text: string): number {
  return addNoteImpl(doc, paragraph, text, "endnote");
}

function addNoteImpl(
  doc: Docx,
  paragraph: WmlParagraph,
  text: string,
  kind: "footnote" | "endnote",
): number {
  const part = kind === "footnote" ? ensureFootnotesPart(doc) : ensureEndnotesPart(doc);
  const usedIds = new Set<number>();
  for (const f of part.footnotes) {
    const id = f.attrs.find((a) => a.name.uri === WML_NS && a.name.local === "id")?.value;
    if (id !== undefined) {
      const n = Number.parseInt(id, 10);
      if (Number.isFinite(n)) usedIds.add(n);
    }
  }
  let nextId = 1;
  while (usedIds.has(nextId)) nextId++;
  const note = buildFootnote({ id: nextId, text }, kind);
  part.footnotes.push(note);
  if (kind === "footnote") doc.footnotesDirty = true;
  else doc.endnotesDirty = true;

  // Insert a reference run at the end of the paragraph.
  paragraph.children.push({
    kind: "raw",
    node: buildFootnoteReferenceRun(
      nextId,
      kind === "footnote" ? "footnoteReference" : "endnoteReference",
    ),
  });
  doc.dirty = true;
  return nextId;
}

function ensureFootnotesPart(doc: Docx): WmlFootnotesPart {
  const existing = footnotesPart(doc);
  if (existing) return existing;
  addPart(doc.opc, {
    name: FOOTNOTES_PART_NAME,
    contentType: WML_CONTENT_TYPES.footnotes,
    data: new TextEncoder().encode(SEED_FOOTNOTES_XML),
  });
  const docRels = partRelationships(doc.opc, doc.partName);
  if (relationshipsByType(docRels, WML_RELATIONSHIPS.footnotes).length === 0) {
    addRelationship(docRels, { type: WML_RELATIONSHIPS.footnotes, target: "footnotes.xml" });
  }
  const xml = new TextDecoder("utf-8").decode(
    getPart(doc.opc, FOOTNOTES_PART_NAME)?.data ?? new Uint8Array(),
  );
  doc.footnotesCache = parseFootnotesPart(parseXml(xml));
  doc.footnotesDirty = true;
  return doc.footnotesCache;
}

function ensureEndnotesPart(doc: Docx): WmlFootnotesPart {
  const existing = endnotesPart(doc);
  if (existing) return existing;
  addPart(doc.opc, {
    name: ENDNOTES_PART_NAME,
    contentType: WML_CONTENT_TYPES.endnotes,
    data: new TextEncoder().encode(SEED_ENDNOTES_XML),
  });
  const docRels = partRelationships(doc.opc, doc.partName);
  if (relationshipsByType(docRels, WML_RELATIONSHIPS.endnotes).length === 0) {
    addRelationship(docRels, { type: WML_RELATIONSHIPS.endnotes, target: "endnotes.xml" });
  }
  const xml = new TextDecoder("utf-8").decode(
    getPart(doc.opc, ENDNOTES_PART_NAME)?.data ?? new Uint8Array(),
  );
  doc.endnotesCache = parseFootnotesPart(parseXml(xml));
  doc.endnotesDirty = true;
  return doc.endnotesCache;
}

/** Parsed `word/comments.xml` AST, or `undefined` if absent. */
export function commentsPart(doc: Docx): WmlCommentsPart | undefined {
  if (doc.commentsCache) return doc.commentsCache;
  const part = getPart(doc.opc, COMMENTS_PART_NAME);
  if (!part) return undefined;
  const xml = new TextDecoder("utf-8").decode(part.data);
  doc.commentsCache = parseCommentsPart(parseXml(xml));
  return doc.commentsCache;
}

/**
 * Attach a comment to the given paragraph. The comment range covers the
 * entire paragraph: `<w:commentRangeStart>` is inserted at the start,
 * `<w:commentRangeEnd>` plus a `<w:commentReference>` icon run at the
 * end. The comment body is stored in `word/comments.xml`.
 *
 * Returns the assigned comment id.
 */
export function addComment(doc: Docx, paragraph: WmlParagraph, options: AddCommentOptions): number {
  const part = ensureCommentsPart(doc);
  const usedIds = new Set<number>();
  for (const c of part.comments) {
    const id = c.attrs.find((a) => a.name.uri === WML_NS && a.name.local === "id")?.value;
    if (id !== undefined) {
      const n = Number.parseInt(id, 10);
      if (Number.isFinite(n)) usedIds.add(n);
    }
  }
  let nextId = 0;
  while (usedIds.has(nextId)) nextId++;
  const comment = buildComment({
    id: nextId,
    author: options.author,
    text: options.text,
    ...(options.initials !== undefined ? { initials: options.initials } : {}),
    ...(options.date !== undefined ? { date: options.date } : {}),
  });
  part.comments.push(comment);
  doc.commentsDirty = true;

  // Wrap the paragraph's inline children with rangeStart/rangeEnd+ref.
  paragraph.children = [
    { kind: "raw", node: buildCommentRangeStart(nextId) },
    ...paragraph.children,
    { kind: "raw", node: buildCommentRangeEnd(nextId) },
    { kind: "raw", node: buildCommentReferenceRun(nextId) },
  ];
  doc.dirty = true;
  return nextId;
}

function ensureCommentsPart(doc: Docx): WmlCommentsPart {
  const existing = commentsPart(doc);
  if (existing) return existing;
  addPart(doc.opc, {
    name: COMMENTS_PART_NAME,
    contentType: WML_CONTENT_TYPES.comments,
    data: new TextEncoder().encode(EMPTY_COMMENTS_XML),
  });
  const docRels = partRelationships(doc.opc, doc.partName);
  if (relationshipsByType(docRels, WML_RELATIONSHIPS.comments).length === 0) {
    addRelationship(docRels, { type: WML_RELATIONSHIPS.comments, target: "comments.xml" });
  }
  const xml = new TextDecoder("utf-8").decode(
    getPart(doc.opc, COMMENTS_PART_NAME)?.data ?? new Uint8Array(),
  );
  doc.commentsCache = parseCommentsPart(parseXml(xml));
  doc.commentsDirty = true;
  return doc.commentsCache;
}

/**
 * Add a header part with the given text and wire it to the body's
 * trailing `<w:sectPr>`. Creates the sectPr if missing.
 *
 * Returns the relationship id assigned to the header part.
 */
export function addHeader(doc: Docx, text: string, type: HeaderFooterType = "default"): string {
  const partName = allocateHeaderPartName(doc);
  const xml = buildHeaderXml(text);
  addPart(doc.opc, {
    name: partName,
    contentType: WML_CONTENT_TYPES.header,
    data: new TextEncoder().encode(xml),
  });
  const docRels = partRelationships(doc.opc, doc.partName);
  const rel = addRelationship(docRels, {
    type: WML_RELATIONSHIPS.header,
    target: targetRelativeToDoc(doc, partName),
  });
  const sectPr = ensureBodySectPr(doc);
  addSectPrHeaderRef(sectPr, type, rel.id);
  doc.dirty = true;
  return rel.id;
}

/**
 * Add a footer that displays a centered PAGE field with optional
 * surrounding text. Word renders the live page number when the document
 * is opened.
 */
export function addPageNumberFooter(
  doc: Docx,
  prefix = "",
  suffix = "",
  type: HeaderFooterType = "default",
): string {
  const partName = allocateFooterPartName(doc);
  const xml = buildPageNumberFooterXml(prefix, suffix);
  addPart(doc.opc, {
    name: partName,
    contentType: WML_CONTENT_TYPES.footer,
    data: new TextEncoder().encode(xml),
  });
  const docRels = partRelationships(doc.opc, doc.partName);
  const rel = addRelationship(docRels, {
    type: WML_RELATIONSHIPS.footer,
    target: targetRelativeToDoc(doc, partName),
  });
  const sectPr = ensureBodySectPr(doc);
  addSectPrFooterRef(sectPr, type, rel.id);
  doc.dirty = true;
  return rel.id;
}

/**
 * Append a complex field (PAGE, NUMPAGES, DATE, MERGEFIELD …) to the
 * end of the given paragraph. The output is the canonical
 * `fldChar begin/instrText/separate/display/fldChar end` run sequence.
 */
export function appendField(
  doc: Docx,
  paragraph: WmlParagraph,
  instruction: string,
  displayText = "",
): void {
  for (const run of buildFieldRuns(instruction, displayText)) {
    paragraph.children.push({ kind: "raw", node: run });
  }
  doc.dirty = true;
}

export interface AddTableOfContentsOptions {
  /** Inclusive heading-level range. Defaults to 1-3. */
  readonly headingLevels?: { from: number; to: number };
  /** Placeholder text Word displays before the field is refreshed. */
  readonly placeholderText?: string;
}

/**
 * Append a Table of Contents field to the document body. The TOC is
 * inserted as a single paragraph holding a complex field whose
 * instruction follows Word's canonical form
 * `TOC \o "1-3" \h \z \u` (heading levels 1-3, hyperlinks on, hide
 * tab leader/page number on web layout, use outline level).
 *
 * Word recomputes the actual entries on first open — until then the
 * placeholder text is shown.
 */
export function addTableOfContents(
  doc: Docx,
  options: AddTableOfContentsOptions = {},
): WmlParagraph {
  const from = options.headingLevels?.from ?? 1;
  const to = options.headingLevels?.to ?? 3;
  if (!Number.isInteger(from) || !Number.isInteger(to) || from < 1 || to < from || to > 9) {
    throw new Error(
      `Invalid TOC heading-level range: ${from}-${to}. Expected 1 <= from <= to <= 9.`,
    );
  }
  const instruction = `TOC \\o "${from}-${to}" \\h \\z \\u`;
  const placeholder = options.placeholderText ?? "Right-click and choose Update Field.";
  const paragraph: WmlParagraph = {
    kind: "paragraph",
    children: [],
    extras: [],
  };
  for (const run of buildFieldRuns(instruction, placeholder)) {
    paragraph.children.push({ kind: "raw", node: run });
  }
  doc.document.body.blocks.push(paragraph);
  doc.dirty = true;
  return paragraph;
}

/**
 * Append a MERGEFIELD field referencing `fieldName` to `paragraph`. The
 * placeholder text Word shows is `«fieldName»` by default, matching
 * Word's own UI; pass `displayText` to override.
 */
export function appendMergeField(
  doc: Docx,
  paragraph: WmlParagraph,
  fieldName: string,
  displayText?: string,
): void {
  if (!/^[A-Za-z_][\w]*$/.test(fieldName)) {
    throw new Error(
      `MERGEFIELD name must be an alphanumeric identifier (letters/digits/underscore, starting with a letter), got ${JSON.stringify(fieldName)}.`,
    );
  }
  const instruction = `MERGEFIELD ${fieldName} \\* MERGEFORMAT`;
  const display = displayText ?? `«${fieldName}»`;
  for (const run of buildFieldRuns(instruction, display)) {
    paragraph.children.push({ kind: "raw", node: run });
  }
  doc.dirty = true;
}

/** Like {@link addHeader} but for a footer. */
export function addFooter(doc: Docx, text: string, type: HeaderFooterType = "default"): string {
  const partName = allocateFooterPartName(doc);
  const xml = buildFooterXml(text);
  addPart(doc.opc, {
    name: partName,
    contentType: WML_CONTENT_TYPES.footer,
    data: new TextEncoder().encode(xml),
  });
  const docRels = partRelationships(doc.opc, doc.partName);
  const rel = addRelationship(docRels, {
    type: WML_RELATIONSHIPS.footer,
    target: targetRelativeToDoc(doc, partName),
  });
  const sectPr = ensureBodySectPr(doc);
  addSectPrFooterRef(sectPr, type, rel.id);
  doc.dirty = true;
  return rel.id;
}

/** Set the page size on the body's trailing `<w:sectPr>`. */
export function setPageSize(doc: Docx, size: PageSize): void {
  const sectPr = ensureBodySectPr(doc);
  setSectPrPageSize(sectPr, size);
  doc.dirty = true;
}

/** Set the page margins on the body's trailing `<w:sectPr>`. */
export function setPageMargins(doc: Docx, margins: PageMargins): void {
  const sectPr = ensureBodySectPr(doc);
  setSectPrPageMargins(sectPr, margins);
  doc.dirty = true;
}

/** Switch the page orientation (portrait/landscape) on the body sectPr. */
export function setPageOrientation(doc: Docx, orientation: "portrait" | "landscape"): void {
  const sectPr = ensureBodySectPr(doc);
  const pgSz = sectPr.children.find(
    (c): c is XmlElement => c.kind === "element" && c.name.local === "pgSz",
  );
  if (pgSz) {
    // Swap width/height when toggling orientation if needed.
    const wAttr = pgSz.attrs.find((a) => a.name.local === "w");
    const hAttr = pgSz.attrs.find((a) => a.name.local === "h");
    const widthTwips = wAttr ? Number.parseInt(wAttr.value, 10) : PAGE_SIZE_LETTER.widthTwips;
    const heightTwips = hAttr ? Number.parseInt(hAttr.value, 10) : PAGE_SIZE_LETTER.heightTwips;
    // For landscape, width > height; for portrait, height > width.
    const isLandscape = orientation === "landscape";
    const newWidth = isLandscape
      ? Math.max(widthTwips, heightTwips)
      : Math.min(widthTwips, heightTwips);
    const newHeight = isLandscape
      ? Math.min(widthTwips, heightTwips)
      : Math.max(widthTwips, heightTwips);
    setSectPrPageSize(sectPr, { widthTwips: newWidth, heightTwips: newHeight, orientation });
  } else {
    setSectPrPageSize(sectPr, { ...PAGE_SIZE_LETTER, orientation });
  }
  doc.dirty = true;
}

function ensureBodySectPr(doc: Docx): XmlElement {
  if (doc.document.body.sectPr) return doc.document.body.sectPr;
  const sectPr: XmlElement = {
    kind: "element",
    name: { uri: WML_NS, local: "sectPr", prefix: "w" },
    attrs: [],
    children: [
      {
        kind: "element",
        name: { uri: WML_NS, local: "pgSz", prefix: "w" },
        attrs: [
          {
            name: { uri: WML_NS, local: "w", prefix: "w" },
            value: "12240",
            isNamespaceDecl: false,
          },
          {
            name: { uri: WML_NS, local: "h", prefix: "w" },
            value: "15840",
            isNamespaceDecl: false,
          },
        ],
        children: [],
        xmlSpace: "default",
        selfClosing: true,
      },
    ],
    xmlSpace: "default",
    selfClosing: false,
  };
  doc.document.body.sectPr = sectPr;
  return sectPr;
}

function allocateHeaderPartName(doc: Docx): string {
  let n = 1;
  while (hasPart(doc.opc, `/word/header${n}.xml`)) n++;
  return `/word/header${n}.xml`;
}

function allocateFooterPartName(doc: Docx): string {
  let n = 1;
  while (hasPart(doc.opc, `/word/footer${n}.xml`)) n++;
  return `/word/footer${n}.xml`;
}

function targetRelativeToDoc(doc: Docx, partName: string): string {
  if (partName.startsWith("/word/")) return partName.slice("/word/".length);
  return partName.startsWith("/") ? partName.slice(1) : partName;
}

/**
 * Serialize to a `Blob` with the `.docx` MIME type. Works in any
 * environment where `Blob` is a global (modern browsers and Node 18+).
 */
export function toBlob(doc: Docx): Blob {
  const bytes = toUint8Array(doc);
  // Slice into a fresh ArrayBuffer view to satisfy BlobPart's invariance
  // over ArrayBuffer (the TS lib treats SharedArrayBuffer-backed views as
  // incompatible).
  const buf = bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  ) as ArrayBuffer;
  return new Blob([new Uint8Array(buf)], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}

/**
 * Run structural validation on the underlying package. Catches the most
 * common "Word needs to repair the file" causes (missing rel targets,
 * unmatched comment / footnote / endnote ids, broken bookmark pairs,
 * missing media parts). Returns the issues; never throws.
 */
export function validate(doc: Docx): ValidationIssue[] {
  if (doc.dirty) {
    flushDocument(doc);
    doc.dirty = false;
  }
  if (doc.stylesDirty && doc.stylesCache) {
    flushStyles(doc, doc.stylesCache);
    doc.stylesDirty = false;
  }
  if (doc.numberingDirty && doc.numberingCache) {
    flushNumbering(doc, doc.numberingCache);
    doc.numberingDirty = false;
  }
  if (doc.commentsDirty && doc.commentsCache) {
    flushComments(doc, doc.commentsCache);
    doc.commentsDirty = false;
  }
  if (doc.footnotesDirty && doc.footnotesCache) {
    flushNotes(doc, doc.footnotesCache, FOOTNOTES_PART_NAME, "footnotes");
    doc.footnotesDirty = false;
  }
  if (doc.endnotesDirty && doc.endnotesCache) {
    flushNotes(doc, doc.endnotesCache, ENDNOTES_PART_NAME, "endnotes");
    doc.endnotesDirty = false;
  }
  return validatePackage(doc.opc);
}

/**
 * Coarse content statistics: paragraph / table / image / comment /
 * footnote / endnote counts, plus rough word and character counts
 * based on the body's visible text.
 */
export function statistics(doc: Docx): {
  paragraphs: number;
  tables: number;
  images: number;
  comments: number;
  footnotes: number;
  endnotes: number;
  headings: number;
  words: number;
  characters: number;
  charactersNoSpaces: number;
} {
  const paragraphsCount = paragraphs(doc).length;
  const tablesCount = tables(doc).length;
  const imagesCount = images(doc).length;
  const commentsCount = commentsPart(doc)?.comments.length ?? 0;
  const footnotesCount = Math.max(0, (footnotesPart(doc)?.footnotes.length ?? 0) - 2);
  const endnotesCount = Math.max(0, (endnotesPart(doc)?.footnotes.length ?? 0) - 2);
  const headingsCount = outline(doc).length;
  const docText = text(doc);
  const wordsArr = docText.split(/\s+/).filter((w) => w.length > 0);
  return {
    paragraphs: paragraphsCount,
    tables: tablesCount,
    images: imagesCount,
    comments: commentsCount,
    footnotes: footnotesCount,
    endnotes: endnotesCount,
    headings: headingsCount,
    words: wordsArr.length,
    characters: docText.length,
    charactersNoSpaces: docText.replace(/\s/g, "").length,
  };
}

/**
 * Produce an independent copy of this document. The two documents share
 * no mutable state — edits to one do not affect the other.
 *
 * Internally this serializes and re-reads, which incurs a ZIP round-trip
 * but is the cleanest way to guarantee a deep copy across both the OPC
 * package and the WML AST.
 */
export function clone(doc: Docx): Docx {
  return openDocx(toUint8Array(doc));
}

/** Serialize the package back to `.docx` bytes. */
export function toUint8Array(doc: Docx): Uint8Array {
  if (doc.dirty) {
    flushDocument(doc);
    doc.dirty = false;
  }
  if (doc.stylesDirty && doc.stylesCache) {
    flushStyles(doc, doc.stylesCache);
    doc.stylesDirty = false;
  }
  if (doc.numberingDirty && doc.numberingCache) {
    flushNumbering(doc, doc.numberingCache);
    doc.numberingDirty = false;
  }
  if (doc.commentsDirty && doc.commentsCache) {
    flushComments(doc, doc.commentsCache);
    doc.commentsDirty = false;
  }
  if (doc.footnotesDirty && doc.footnotesCache) {
    flushNotes(doc, doc.footnotesCache, FOOTNOTES_PART_NAME, "footnotes");
    doc.footnotesDirty = false;
  }
  if (doc.endnotesDirty && doc.endnotesCache) {
    flushNotes(doc, doc.endnotesCache, ENDNOTES_PART_NAME, "endnotes");
    doc.endnotesDirty = false;
  }
  return writeOpcPackage(doc.opc);
}

function flushNotes(
  doc: Docx,
  part: WmlFootnotesPart,
  partName: string,
  rootLocal: "footnotes" | "endnotes",
): void {
  const xmlDoc = writeFootnotesPart(part, rootLocal);
  const xmlText = serializeXml(xmlDoc);
  const xmlPart = getPart(doc.opc, partName);
  if (!xmlPart) return;
  xmlPart.data = new TextEncoder().encode(xmlText);
}

function flushComments(doc: Docx, part: WmlCommentsPart): void {
  const xmlDoc = writeCommentsPart(part);
  const xmlText = serializeXml(xmlDoc);
  const commentsPart = getPart(doc.opc, COMMENTS_PART_NAME);
  if (!commentsPart) return;
  commentsPart.data = new TextEncoder().encode(xmlText);
}

function flushNumbering(doc: Docx, part: WmlNumberingPart): void {
  const xmlDoc = writeNumberingPart(part);
  const xmlText = serializeXml(xmlDoc);
  const numberingPart = getPart(doc.opc, NUMBERING_PART_NAME);
  if (!numberingPart) return;
  numberingPart.data = new TextEncoder().encode(xmlText);
}

function flushDocument(doc: Docx): void {
  const xmlDoc = writeWmlDocument(doc.document);
  const xmlText = serializeXml(xmlDoc);
  const part = getPart(doc.opc, doc.partName);
  if (!part) {
    throw new Error(`Lost document part: ${doc.partName}`);
  }
  part.data = new TextEncoder().encode(xmlText);
}

function flushStyles(doc: Docx, stylesPart: WmlStylesPart): void {
  const xmlDoc = writeStylesPart(stylesPart);
  const xmlText = serializeXml(xmlDoc);
  const part = getPart(doc.opc, STYLES_PART_NAME);
  if (!part) return;
  part.data = new TextEncoder().encode(xmlText);
}

function wmlEmptyEl(local: string) {
  return {
    kind: "element" as const,
    name: { uri: WML_NS, local, prefix: "w" },
    attrs: [] as never[],
    children: [] as never[],
    xmlSpace: "default" as const,
    selfClosing: true,
  };
}

function findMainDocumentPart(pkg: OpcPackage): Part | undefined {
  for (const rel of relationshipsByType(
    packageRelationships(pkg),
    WML_RELATIONSHIPS.officeDocument,
  )) {
    if (rel.targetMode !== "Internal") continue;
    const partName = normalizePartName(rel.target.startsWith("/") ? rel.target : `/${rel.target}`);
    if (partName === CONTENT_TYPES_PART_NAME) continue;
    const part = getPart(pkg, partName);
    if (part) return part;
  }
  return getPart(pkg, DOCUMENT_PART_FALLBACK);
}

function documentTextOfParagraph(p: WmlParagraph): string {
  let acc = "";
  for (const child of p.children) {
    if (child.kind === "run") {
      for (const piece of child.pieces) {
        if (piece.kind === "text" || piece.kind === "delText") acc += piece.value;
      }
    } else if (child.kind === "raw") {
      const visit = (el: XmlElement): void => {
        if (el.name.uri === WML_NS && (el.name.local === "t" || el.name.local === "delText")) {
          for (const c of el.children) {
            if (c.kind === "text") acc += c.value;
            else if (c.kind === "cdata") acc += c.value;
          }
          return;
        }
        for (const c of el.children) if (c.kind === "element") visit(c);
      };
      visit(child.node);
    }
  }
  return acc;
}

function buildPgSzElement(size: PageSize): XmlElement {
  const attrs = [
    {
      name: { uri: WML_NS, local: "w", prefix: "w" },
      value: String(size.widthTwips),
      isNamespaceDecl: false,
    },
    {
      name: { uri: WML_NS, local: "h", prefix: "w" },
      value: String(size.heightTwips),
      isNamespaceDecl: false,
    },
  ];
  if (size.orientation === "landscape") {
    attrs.push({
      name: { uri: WML_NS, local: "orient", prefix: "w" },
      value: "landscape",
      isNamespaceDecl: false,
    });
  }
  return {
    kind: "element",
    name: { uri: WML_NS, local: "pgSz", prefix: "w" },
    attrs,
    children: [],
    xmlSpace: "default",
    selfClosing: true,
  };
}

function buildPgMarElement(m: PageMargins): XmlElement {
  return {
    kind: "element",
    name: { uri: WML_NS, local: "pgMar", prefix: "w" },
    attrs: [
      {
        name: { uri: WML_NS, local: "top", prefix: "w" },
        value: String(m.top),
        isNamespaceDecl: false,
      },
      {
        name: { uri: WML_NS, local: "right", prefix: "w" },
        value: String(m.right),
        isNamespaceDecl: false,
      },
      {
        name: { uri: WML_NS, local: "bottom", prefix: "w" },
        value: String(m.bottom),
        isNamespaceDecl: false,
      },
      {
        name: { uri: WML_NS, local: "left", prefix: "w" },
        value: String(m.left),
        isNamespaceDecl: false,
      },
      {
        name: { uri: WML_NS, local: "header", prefix: "w" },
        value: String(m.header),
        isNamespaceDecl: false,
      },
      {
        name: { uri: WML_NS, local: "footer", prefix: "w" },
        value: String(m.footer),
        isNamespaceDecl: false,
      },
      {
        name: { uri: WML_NS, local: "gutter", prefix: "w" },
        value: String(m.gutter),
        isNamespaceDecl: false,
      },
    ],
    children: [],
    xmlSpace: "default",
    selfClosing: true,
  };
}

function isParagraph(block: WmlBlock): block is WmlParagraph {
  return block.kind === "paragraph";
}

function isTable(block: WmlBlock): block is WmlTable {
  return block.kind === "table";
}

function collectVisibleTextFromElement(el: XmlElement): string {
  let acc = "";
  let paragraphBreak = false;
  const visit = (e: XmlElement): void => {
    if (e.name.uri === WML_NS && e.name.local === "p" && paragraphBreak) {
      acc += "\n";
    }
    paragraphBreak = e.name.uri === WML_NS && e.name.local === "p";
    if (e.name.uri === WML_NS && (e.name.local === "t" || e.name.local === "delText")) {
      for (const c of e.children) {
        if (c.kind === "text") acc += c.value;
        else if (c.kind === "cdata") acc += c.value;
      }
      return;
    }
    for (const c of e.children) {
      if (c.kind === "element") visit(c);
    }
  };
  visit(el);
  return acc;
}
