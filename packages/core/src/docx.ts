import { parseXml, serializeXml } from "@word-kit/ooxml-xml";
import {
  buildMinimalDocx,
  CONTENT_TYPES_PART_NAME,
  normalizePartName,
  OpcPackage,
  type Part,
  type RelationshipSet,
} from "@word-kit/opc";
import {
  acceptAllRevisions,
  addSectPrFooterRef,
  addSectPrHeaderRef,
  CORE_PROPERTIES_CONTENT_TYPE,
  CORE_PROPERTIES_REL_TYPE,
  type DocumentCoreProperties,
  EMPTY_CORE_PROPERTIES_XML,
  parseCoreProperties,
  writeCoreProperties,
  buildAbstractNum,
  buildComment,
  buildCommentRangeEnd,
  buildCommentRangeStart,
  buildCommentReferenceRun,
  buildFootnote,
  buildFootnoteReferenceRun,
  buildFooterXml,
  buildHeaderXml,
  buildHyperlink,
  buildHyperlinkRun,
  buildInlineDrawing,
  buildNum,
  buildPPrWithNumPr,
  type BuildStyleOptions,
  buildStyle,
  buildTextTable,
  type BuildTableOptions,
  bulletAbstractNumLevels,
  decimalAbstractNumLevels,
  documentText,
  EMPTY_COMMENTS_XML,
  EMPTY_NUMBERING_XML,
  extensionForImageContentType,
  parseFootnotesPart,
  SEED_ENDNOTES_XML,
  SEED_FOOTNOTES_XML,
  type WmlFootnotesPart,
  writeFootnotesPart,
  findStyle,
  findText,
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
  rejectAllRevisions,
  replaceText,
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

const BULLET_ABSTRACT_NUM_ID = 9000;
const DECIMAL_ABSTRACT_NUM_ID = 9001;

export interface AddCommentOptions {
  readonly author: string;
  readonly text: string;
  readonly initials?: string;
  readonly date?: string;
}

export class Docx {
  private readonly pkg: OpcPackage;
  private readonly docModel: WmlDocument;
  private readonly partName: string;
  private dirty = false;
  private stylesCache: WmlStylesPart | undefined;
  private stylesDirty = false;
  private numberingCache: WmlNumberingPart | undefined;
  private numberingDirty = false;
  private commentsCache: WmlCommentsPart | undefined;
  private commentsDirty = false;
  private footnotesCache: WmlFootnotesPart | undefined;
  private footnotesDirty = false;
  private endnotesCache: WmlFootnotesPart | undefined;
  private endnotesDirty = false;

  private constructor(pkg: OpcPackage, doc: WmlDocument, partName: string) {
    this.pkg = pkg;
    this.docModel = doc;
    this.partName = partName;
  }

  /** Parse an existing `.docx` package. */
  static open(bytes: Uint8Array): Docx {
    const pkg = OpcPackage.read(bytes);
    const part = findMainDocumentPart(pkg);
    if (!part) {
      throw new Error("Package has no main WordprocessingML document part");
    }
    const xml = new TextDecoder("utf-8").decode(part.data);
    const wml = parseWmlDocument(parseXml(xml));
    return new Docx(pkg, wml, part.name);
  }

  /**
   * Convenience for browser-side callers: open a `.docx` from a `Blob` or
   * `File`. Awaits the underlying `ArrayBuffer` and delegates to
   * {@link Docx.open}.
   */
  static async fromBlob(blob: Blob): Promise<Docx> {
    const buf = await blob.arrayBuffer();
    return Docx.open(new Uint8Array(buf));
  }

  /**
   * Return one entry per header part referenced by `document.xml`. Each
   * entry includes the part name, the relationship id, and the plain-text
   * content extracted from the part.
   */
  get headers(): Array<{ partName: string; relId: string; text: string }> {
    return this.collectHeaderFooterParts(WML_RELATIONSHIPS.header);
  }

  /** Same shape as {@link headers} but for footer parts. */
  get footers(): Array<{ partName: string; relId: string; text: string }> {
    return this.collectHeaderFooterParts(WML_RELATIONSHIPS.footer);
  }

  /**
   * Return one entry per media (image) part. Useful for inspecting which
   * images a document already contains and feeding their bytes back into
   * {@link replaceImage}.
   */
  get images(): Array<{ partName: string; contentType: string; data: Uint8Array }> {
    const out: Array<{ partName: string; contentType: string; data: Uint8Array }> = [];
    for (const part of this.pkg.listParts()) {
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
  replaceImage(partName: string, newBytes: Uint8Array): boolean {
    const part = this.pkg.getPart(partName);
    if (!part) return false;
    part.data = newBytes;
    return true;
  }

  private collectHeaderFooterParts(
    relType: string,
  ): Array<{ partName: string; relId: string; text: string }> {
    const out: Array<{ partName: string; relId: string; text: string }> = [];
    const docRels = this.pkg.partRelationships(this.partName);
    for (const rel of docRels.byType(relType)) {
      if (rel.targetMode === "External") continue;
      const partName = this.resolvePartTarget(rel.target);
      const part = this.pkg.getPart(partName);
      if (!part) continue;
      const xml = new TextDecoder("utf-8").decode(part.data);
      const xmlDoc = parseXml(xml);
      const text = collectVisibleTextFromElement(xmlDoc.root);
      out.push({ partName, relId: rel.id, text });
    }
    return out;
  }

  /** Create a fresh `.docx`. */
  static create(options: DocxCreateOptions = {}): Docx {
    const pkg = buildMinimalDocx();
    const docPart = pkg.getPart(DOCUMENT_PART_FALLBACK);
    if (!docPart) {
      throw new Error("Minimal docx is missing /word/document.xml");
    }
    // Seed a minimal but valid styles.xml so the produced docx opens cleanly
    // in both Word and LibreOffice without an "uses an unknown style" prompt.
    pkg.addPart({
      name: STYLES_PART_NAME,
      contentType: WML_CONTENT_TYPES.styles,
      data: new TextEncoder().encode(MINIMAL_STYLES_XML),
    });
    const docRels = pkg.partRelationships(DOCUMENT_PART_FALLBACK);
    docRels.add({ type: WML_RELATIONSHIPS.styles, target: "styles.xml" });

    const xml = new TextDecoder("utf-8").decode(docPart.data);
    const wml = parseWmlDocument(parseXml(xml));
    const docx = new Docx(pkg, wml, docPart.name);
    if (options.paragraphs !== undefined) {
      // Clear the seed empty paragraph; the caller is supplying the body.
      wml.body.blocks.length = 0;
      for (const text of options.paragraphs) {
        docx.appendParagraph(text);
      }
    }
    return docx;
  }

  /** The underlying OPC package. Mutations here bypass document-dirty tracking. */
  get opc(): OpcPackage {
    return this.pkg;
  }

  /** The parsed WordprocessingML AST of `word/document.xml`. */
  get document(): WmlDocument {
    return this.docModel;
  }

  /** Paragraph blocks in document order. */
  get paragraphs(): readonly WmlParagraph[] {
    return this.docModel.body.blocks.filter(isParagraph);
  }

  /** Table blocks in document order. */
  get tables(): readonly WmlTable[] {
    return this.docModel.body.blocks.filter(isTable);
  }

  /**
   * Parsed `word/styles.xml` AST, or `undefined` if the package has no
   * styles part. Lazily parsed on first access; subsequent mutations are
   * flushed back on `toUint8Array()`.
   */
  get stylesPart(): WmlStylesPart | undefined {
    if (this.stylesCache) return this.stylesCache;
    const part = this.pkg.getPart(STYLES_PART_NAME);
    if (!part) return undefined;
    const xml = new TextDecoder("utf-8").decode(part.data);
    this.stylesCache = parseStylesPart(parseXml(xml));
    return this.stylesCache;
  }

  /**
   * Add (or replace) a `<w:style>` entry. Creates `word/styles.xml` and
   * its relationship if they do not already exist.
   */
  addStyle(options: BuildStyleOptions): void {
    const part = this.ensureStylesPart();
    const existing = findStyle(part, options.styleId);
    const built = buildStyle(options);
    if (existing) {
      const idx = part.styles.indexOf(existing);
      if (idx >= 0) part.styles[idx] = built;
    } else {
      part.styles.push(built);
    }
    this.stylesDirty = true;
  }

  /**
   * Parsed `word/numbering.xml` AST, or `undefined` if absent. Lazy and
   * cached just like {@link stylesPart}.
   */
  get numberingPart(): WmlNumberingPart | undefined {
    if (this.numberingCache) return this.numberingCache;
    const part = this.pkg.getPart(NUMBERING_PART_NAME);
    if (!part) return undefined;
    const xml = new TextDecoder("utf-8").decode(part.data);
    this.numberingCache = parseNumberingPart(parseXml(xml));
    return this.numberingCache;
  }

  /**
   * Append a bullet list (one paragraph per item) to the body. Sets up
   * `numbering.xml` and its relationship on first use.
   */
  addBulletList(items: readonly string[]): WmlParagraph[] {
    const idValue = this.ensureBulletNumbering();
    return items.map((text) => this.appendListParagraph(text, idValue));
  }

  /**
   * Append a numbered list (one paragraph per item) to the body. Sets up
   * `numbering.xml` and its relationship on first use.
   */
  addNumberedList(items: readonly string[]): WmlParagraph[] {
    const idValue = this.ensureDecimalNumbering();
    return items.map((text) => this.appendListParagraph(text, idValue));
  }

  private appendListParagraph(text: string, numIdValue: number): WmlParagraph {
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
    this.docModel.body.blocks.push(paragraph);
    this.dirty = true;
    return paragraph;
  }

  private ensureBulletNumbering(): number {
    return this.ensureNumberingDefinition(BULLET_ABSTRACT_NUM_ID, bulletAbstractNumLevels);
  }

  private ensureDecimalNumbering(): number {
    return this.ensureNumberingDefinition(DECIMAL_ABSTRACT_NUM_ID, decimalAbstractNumLevels);
  }

  private ensureNumberingDefinition(
    abstractNumIdValue: number,
    levelsFactory: () => Array<{
      ilvl: number;
      numFmt: "bullet" | "decimal" | string;
      lvlText: string;
    }>,
  ): number {
    const part = this.ensureNumberingPart();
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
      this.numberingDirty = true;
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
      this.numberingDirty = true;
      chosenId = next;
    }
    return chosenId;
  }

  private ensureNumberingPart(): WmlNumberingPart {
    const existing = this.numberingPart;
    if (existing) return existing;
    this.pkg.addPart({
      name: NUMBERING_PART_NAME,
      contentType: WML_CONTENT_TYPES.numbering,
      data: new TextEncoder().encode(EMPTY_NUMBERING_XML),
    });
    const docRels = this.pkg.partRelationships(this.partName);
    if (docRels.byType(WML_RELATIONSHIPS.numbering).length === 0) {
      docRels.add({ type: WML_RELATIONSHIPS.numbering, target: "numbering.xml" });
    }
    const xml = new TextDecoder("utf-8").decode(
      this.pkg.getPart(NUMBERING_PART_NAME)?.data ?? new Uint8Array(),
    );
    this.numberingCache = parseNumberingPart(parseXml(xml));
    this.numberingDirty = true;
    return this.numberingCache;
  }

  private ensureStylesPart(): WmlStylesPart {
    const existing = this.stylesPart;
    if (existing) return existing;
    this.pkg.addPart({
      name: STYLES_PART_NAME,
      contentType: WML_CONTENT_TYPES.styles,
      data: new TextEncoder().encode(MINIMAL_STYLES_XML),
    });
    const docRels = this.pkg.partRelationships(this.partName);
    const hasStylesRel = docRels.byType(WML_RELATIONSHIPS.styles).length > 0;
    if (!hasStylesRel) {
      docRels.add({ type: WML_RELATIONSHIPS.styles, target: "styles.xml" });
    }
    const xml = new TextDecoder("utf-8").decode(
      this.pkg.getPart(STYLES_PART_NAME)?.data ?? new Uint8Array(),
    );
    this.stylesCache = parseStylesPart(parseXml(xml));
    this.stylesDirty = true;
    return this.stylesCache;
  }

  /**
   * Append a table to the body. `rows` is a row-major matrix of strings;
   * each cell becomes a single paragraph with a single run containing the
   * supplied text. Empty cells are allowed.
   */
  addTable(rows: ReadonlyArray<ReadonlyArray<string>>, options: BuildTableOptions = {}): WmlTable {
    const table = buildTextTable(rows, options);
    this.docModel.body.blocks.push(table);
    this.dirty = true;
    return table;
  }

  /** Search the document body for all matches of `query`. */
  findText(query: string | RegExp): TextMatch[] {
    return findText(this.docModel, query);
  }

  /**
   * Find every occurrence of `query` across the body and all
   * header/footer/comment/footnote/endnote parts. Returns an array of
   * `{ partName, matches }` entries (the body uses `"/word/document.xml"`).
   *
   * Useful for inspecting a template before calling
   * {@link replaceTextEverywhere}.
   */
  findTextEverywhere(query: string | RegExp): Array<{
    partName: string;
    matches: TextMatch[];
  }> {
    const out: Array<{ partName: string; matches: TextMatch[] }> = [];
    const bodyMatches = findText(this.docModel, query);
    if (bodyMatches.length > 0) {
      out.push({ partName: this.partName, matches: bodyMatches });
    }
    // Flush dirty side-parts so the XML walk sees their current state.
    if (this.commentsDirty && this.commentsCache) {
      this.flushComments(this.commentsCache);
      this.commentsDirty = false;
      this.commentsCache = undefined;
    }
    if (this.footnotesDirty && this.footnotesCache) {
      this.flushNotes(this.footnotesCache, FOOTNOTES_PART_NAME, "footnotes");
      this.footnotesDirty = false;
      this.footnotesCache = undefined;
    }
    if (this.endnotesDirty && this.endnotesCache) {
      this.flushNotes(this.endnotesCache, ENDNOTES_PART_NAME, "endnotes");
      this.endnotesDirty = false;
      this.endnotesCache = undefined;
    }
    const docRels = this.pkg.partRelationships(this.partName);
    for (const rel of docRels.all) {
      if (rel.targetMode === "External") continue;
      if (
        rel.type !== WML_RELATIONSHIPS.header &&
        rel.type !== WML_RELATIONSHIPS.footer &&
        rel.type !== WML_RELATIONSHIPS.comments &&
        rel.type !== WML_RELATIONSHIPS.footnotes &&
        rel.type !== WML_RELATIONSHIPS.endnotes
      )
        continue;
      const partName = this.resolvePartTarget(rel.target);
      const part = this.pkg.getPart(partName);
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
              ...findText(
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
   */
  replaceText(
    query: string | RegExp,
    replacement: string | ((match: TextMatch) => string),
  ): number {
    const count = replaceText(this.docModel, query, replacement);
    if (count > 0) this.dirty = true;
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
  replaceTextEverywhere(
    query: string | RegExp,
    replacement: string | ((match: TextMatch) => string),
  ): number {
    // Flush any dirty in-memory parts to package bytes so the XML walk sees
    // the latest state. We invalidate the caches afterwards so the next
    // accessor reloads fresh.
    if (this.commentsDirty && this.commentsCache) {
      this.flushComments(this.commentsCache);
      this.commentsDirty = false;
      this.commentsCache = undefined;
    }
    if (this.footnotesDirty && this.footnotesCache) {
      this.flushNotes(this.footnotesCache, FOOTNOTES_PART_NAME, "footnotes");
      this.footnotesDirty = false;
      this.footnotesCache = undefined;
    }
    if (this.endnotesDirty && this.endnotesCache) {
      this.flushNotes(this.endnotesCache, ENDNOTES_PART_NAME, "endnotes");
      this.endnotesDirty = false;
      this.endnotesCache = undefined;
    }

    let total = this.replaceText(query, replacement);

    const partsToVisit = new Set<string>();
    const docRels = this.pkg.partRelationships(this.partName);
    for (const rel of docRels.all) {
      if (rel.targetMode === "External") continue;
      if (
        rel.type === WML_RELATIONSHIPS.header ||
        rel.type === WML_RELATIONSHIPS.footer ||
        rel.type === WML_RELATIONSHIPS.comments ||
        rel.type === WML_RELATIONSHIPS.footnotes ||
        rel.type === WML_RELATIONSHIPS.endnotes
      ) {
        partsToVisit.add(this.resolvePartTarget(rel.target));
      }
    }

    for (const partName of partsToVisit) {
      total += this.replaceTextInPartXml(partName, query, replacement);
    }
    return total;
  }

  private resolvePartTarget(relTarget: string): string {
    if (relTarget.startsWith("/")) return relTarget;
    // Relative to /word/document.xml's folder (/word/).
    return `/word/${relTarget}`;
  }

  private replaceTextInPartXml(
    partName: string,
    query: string | RegExp,
    replacement: string | ((match: TextMatch) => string),
  ): number {
    const part = this.pkg.getPart(partName);
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
  get text(): string {
    return documentText(this.docModel);
  }

  /**
   * Append a paragraph whose only content is a page break. Equivalent to
   * Ctrl+Enter in Word.
   */
  appendPageBreak(): WmlParagraph {
    const piece: WmlRunPiece = { kind: "break", breakType: "page" };
    const run: WmlRun = { kind: "run", pieces: [piece], extras: [] };
    const paragraph: WmlParagraph = { kind: "paragraph", children: [run], extras: [] };
    this.docModel.body.blocks.push(paragraph);
    this.dirty = true;
    return paragraph;
  }

  /**
   * Add a named bookmark covering a paragraph. Returns the assigned numeric
   * bookmark id; the same name must not be reused without removing the
   * existing bookmark first (Word will deduplicate silently otherwise).
   */
  addBookmark(name: string, paragraph: WmlParagraph): number {
    const id = this.allocateBookmarkId();
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
    this.dirty = true;
    return id;
  }

  /**
   * Append a paragraph with a hyperlink that points at an internal bookmark.
   */
  addInternalHyperlink(
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
    this.docModel.body.blocks.push(paragraph);
    this.dirty = true;
    return paragraph;
  }

  private allocateBookmarkId(): number {
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
    for (const block of this.docModel.body.blocks) {
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
  addHyperlink(url: string, text: string, options: { tooltip?: string } = {}): WmlParagraph {
    const docRels = this.pkg.partRelationships(this.partName);
    const rel = docRels.add({
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
    this.docModel.body.blocks.push(paragraph);
    this.dirty = true;
    return paragraph;
  }

  /**
   * Accept every `<w:ins>` and `<w:del>` revision in the document.
   * Insertions are kept and unwrapped; deletions are dropped. Returns the
   * number of revisions resolved.
   */
  acceptAllRevisions(): number {
    const n = acceptAllRevisions(this.docModel);
    if (n > 0) this.dirty = true;
    return n;
  }

  /**
   * Reject every `<w:ins>` and `<w:del>` revision in the document.
   * Insertions are dropped; deletions are kept and unwrapped. Returns the
   * number of revisions resolved.
   */
  rejectAllRevisions(): number {
    const n = rejectAllRevisions(this.docModel);
    if (n > 0) this.dirty = true;
    return n;
  }

  /** Append a paragraph containing a single text run to the body. */
  appendParagraph(text: string, options: AppendParagraphOptions = {}): WmlParagraph {
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
    this.docModel.body.blocks.push(paragraph);
    this.dirty = true;
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
  addImage(bytes: Uint8Array, options: AddImageOptions): WmlParagraph {
    const run = this.addImageRun(bytes, options);
    const paragraph: WmlParagraph = {
      kind: "paragraph",
      children: [run],
      extras: [],
    };
    this.docModel.body.blocks.push(paragraph);
    this.dirty = true;
    return paragraph;
  }

  /**
   * Like {@link addImage} but returns the constructed run without inserting
   * it. Callers can place the run wherever they want (e.g. inside an
   * existing paragraph).
   */
  addImageRun(bytes: Uint8Array, options: AddImageOptions): WmlRun {
    const contentType = options.contentType ?? sniffImageContentType(bytes);
    if (!contentType) {
      throw new Error(
        "addImage: could not detect image content type from bytes; pass options.contentType explicitly",
      );
    }
    const ext = extensionForImageContentType(contentType);
    const partName = this.allocateImagePartName(ext);
    this.pkg.addPart({ name: partName, contentType, data: bytes });
    this.pkg.contentTypes.setDefault(ext, contentType);

    const docRels = this.documentRelationships();
    const rel = docRels.add({
      type: WML_RELATIONSHIPS.image,
      target: this.relativeMediaTarget(partName),
    });

    const docPrId = this.allocateDocPrId();
    const drawing = buildInlineDrawing({
      relId: rel.id,
      widthEmu: options.widthEmu,
      heightEmu: options.heightEmu,
      docPrId,
      ...(options.name !== undefined ? { name: options.name } : {}),
      ...(options.altText !== undefined ? { altText: options.altText } : {}),
    });
    const piece: WmlRunPiece = { kind: "drawing", node: drawing };
    this.dirty = true;
    return { kind: "run", pieces: [piece], extras: [] };
  }

  private documentRelationships(): RelationshipSet {
    const set = this.pkg.partRelationships(this.partName);
    return set;
  }

  private allocateImagePartName(extension: string): string {
    let n = 1;
    // Find lowest unused index.
    while (this.pkg.hasPart(`/word/media/image${n}.${extension}`)) {
      n++;
    }
    return `/word/media/image${n}.${extension}`;
  }

  private relativeMediaTarget(partName: string): string {
    // /word/document.xml relates targets relative to its own folder.
    // /word/media/imageN.ext  →  media/imageN.ext
    if (partName.startsWith("/word/")) return partName.slice("/word/".length);
    return partName.startsWith("/") ? partName.slice(1) : partName;
  }

  private allocateDocPrId(): number {
    // docPr ids must be unique across the document. We bump a counter that
    // starts from the current max we can see in pass-through drawing nodes
    // and from the relationships set length as a coarse upper bound.
    return Math.max(1, this.documentRelationships().all.length + 1);
  }

  /**
   * Read the package's core document properties (title, creator, subject,
   * etc.), parsed from `docProps/core.xml`. Returns an empty record if the
   * part is absent.
   */
  get coreProperties(): DocumentCoreProperties {
    const part = this.pkg.getPart(CORE_PROPERTIES_PART_NAME);
    if (!part) return {};
    const xml = new TextDecoder("utf-8").decode(part.data);
    return parseCoreProperties(parseXml(xml));
  }

  /**
   * Write the package's core document properties. Creates
   * `docProps/core.xml` and the package-level relationship on first use.
   * Subsequent calls merge into existing properties.
   */
  setCoreProperties(props: DocumentCoreProperties): void {
    if (!this.pkg.hasPart(CORE_PROPERTIES_PART_NAME)) {
      this.pkg.addPart({
        name: CORE_PROPERTIES_PART_NAME,
        contentType: CORE_PROPERTIES_CONTENT_TYPE,
        data: new TextEncoder().encode(EMPTY_CORE_PROPERTIES_XML),
      });
      const pkgRels = this.pkg.packageRelationships;
      if (pkgRels.byType(CORE_PROPERTIES_REL_TYPE).length === 0) {
        pkgRels.add({ type: CORE_PROPERTIES_REL_TYPE, target: "docProps/core.xml" });
      }
    }
    const existing = this.coreProperties;
    const merged: DocumentCoreProperties = { ...existing, ...props };
    const xml = serializeXml(writeCoreProperties(merged));
    const part = this.pkg.getPart(CORE_PROPERTIES_PART_NAME);
    if (part) part.data = new TextEncoder().encode(xml);
  }

  /** Parsed `word/footnotes.xml` AST, or `undefined` if absent. */
  get footnotesPart(): WmlFootnotesPart | undefined {
    if (this.footnotesCache) return this.footnotesCache;
    const part = this.pkg.getPart(FOOTNOTES_PART_NAME);
    if (!part) return undefined;
    const xml = new TextDecoder("utf-8").decode(part.data);
    this.footnotesCache = parseFootnotesPart(parseXml(xml));
    return this.footnotesCache;
  }

  /** Parsed `word/endnotes.xml` AST, or `undefined` if absent. */
  get endnotesPart(): WmlFootnotesPart | undefined {
    if (this.endnotesCache) return this.endnotesCache;
    const part = this.pkg.getPart(ENDNOTES_PART_NAME);
    if (!part) return undefined;
    const xml = new TextDecoder("utf-8").decode(part.data);
    this.endnotesCache = parseFootnotesPart(parseXml(xml));
    return this.endnotesCache;
  }

  /**
   * Append a footnote with the given text. Creates `footnotes.xml` (with
   * the standard separator/continuationSeparator entries) on first use,
   * registers the relationship from `word/document.xml`, and inserts a
   * `<w:footnoteReference>` run at the end of the target paragraph.
   *
   * Returns the assigned footnote id (>=1).
   */
  addFootnote(paragraph: WmlParagraph, text: string): number {
    return this.addNoteImpl(paragraph, text, "footnote");
  }

  /** Same shape as {@link addFootnote} but for endnotes. */
  addEndnote(paragraph: WmlParagraph, text: string): number {
    return this.addNoteImpl(paragraph, text, "endnote");
  }

  private addNoteImpl(paragraph: WmlParagraph, text: string, kind: "footnote" | "endnote"): number {
    const part = kind === "footnote" ? this.ensureFootnotesPart() : this.ensureEndnotesPart();
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
    if (kind === "footnote") this.footnotesDirty = true;
    else this.endnotesDirty = true;

    // Insert a reference run at the end of the paragraph.
    paragraph.children.push({
      kind: "raw",
      node: buildFootnoteReferenceRun(
        nextId,
        kind === "footnote" ? "footnoteReference" : "endnoteReference",
      ),
    });
    this.dirty = true;
    return nextId;
  }

  private ensureFootnotesPart(): WmlFootnotesPart {
    const existing = this.footnotesPart;
    if (existing) return existing;
    this.pkg.addPart({
      name: FOOTNOTES_PART_NAME,
      contentType: WML_CONTENT_TYPES.footnotes,
      data: new TextEncoder().encode(SEED_FOOTNOTES_XML),
    });
    const docRels = this.pkg.partRelationships(this.partName);
    if (docRels.byType(WML_RELATIONSHIPS.footnotes).length === 0) {
      docRels.add({ type: WML_RELATIONSHIPS.footnotes, target: "footnotes.xml" });
    }
    const xml = new TextDecoder("utf-8").decode(
      this.pkg.getPart(FOOTNOTES_PART_NAME)?.data ?? new Uint8Array(),
    );
    this.footnotesCache = parseFootnotesPart(parseXml(xml));
    this.footnotesDirty = true;
    return this.footnotesCache;
  }

  private ensureEndnotesPart(): WmlFootnotesPart {
    const existing = this.endnotesPart;
    if (existing) return existing;
    this.pkg.addPart({
      name: ENDNOTES_PART_NAME,
      contentType: WML_CONTENT_TYPES.endnotes,
      data: new TextEncoder().encode(SEED_ENDNOTES_XML),
    });
    const docRels = this.pkg.partRelationships(this.partName);
    if (docRels.byType(WML_RELATIONSHIPS.endnotes).length === 0) {
      docRels.add({ type: WML_RELATIONSHIPS.endnotes, target: "endnotes.xml" });
    }
    const xml = new TextDecoder("utf-8").decode(
      this.pkg.getPart(ENDNOTES_PART_NAME)?.data ?? new Uint8Array(),
    );
    this.endnotesCache = parseFootnotesPart(parseXml(xml));
    this.endnotesDirty = true;
    return this.endnotesCache;
  }

  /** Parsed `word/comments.xml` AST, or `undefined` if absent. */
  get commentsPart(): WmlCommentsPart | undefined {
    if (this.commentsCache) return this.commentsCache;
    const part = this.pkg.getPart(COMMENTS_PART_NAME);
    if (!part) return undefined;
    const xml = new TextDecoder("utf-8").decode(part.data);
    this.commentsCache = parseCommentsPart(parseXml(xml));
    return this.commentsCache;
  }

  /**
   * Attach a comment to the given paragraph. The comment range covers the
   * entire paragraph: `<w:commentRangeStart>` is inserted at the start,
   * `<w:commentRangeEnd>` plus a `<w:commentReference>` icon run at the
   * end. The comment body is stored in `word/comments.xml`.
   *
   * Returns the assigned comment id.
   */
  addComment(paragraph: WmlParagraph, options: AddCommentOptions): number {
    const part = this.ensureCommentsPart();
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
    this.commentsDirty = true;

    // Wrap the paragraph's inline children with rangeStart/rangeEnd+ref.
    paragraph.children = [
      { kind: "raw", node: buildCommentRangeStart(nextId) },
      ...paragraph.children,
      { kind: "raw", node: buildCommentRangeEnd(nextId) },
      { kind: "raw", node: buildCommentReferenceRun(nextId) },
    ];
    this.dirty = true;
    return nextId;
  }

  private ensureCommentsPart(): WmlCommentsPart {
    const existing = this.commentsPart;
    if (existing) return existing;
    this.pkg.addPart({
      name: COMMENTS_PART_NAME,
      contentType: WML_CONTENT_TYPES.comments,
      data: new TextEncoder().encode(EMPTY_COMMENTS_XML),
    });
    const docRels = this.pkg.partRelationships(this.partName);
    if (docRels.byType(WML_RELATIONSHIPS.comments).length === 0) {
      docRels.add({ type: WML_RELATIONSHIPS.comments, target: "comments.xml" });
    }
    const xml = new TextDecoder("utf-8").decode(
      this.pkg.getPart(COMMENTS_PART_NAME)?.data ?? new Uint8Array(),
    );
    this.commentsCache = parseCommentsPart(parseXml(xml));
    this.commentsDirty = true;
    return this.commentsCache;
  }

  /**
   * Add a header part with the given text and wire it to the body's
   * trailing `<w:sectPr>`. Creates the sectPr if missing.
   *
   * Returns the relationship id assigned to the header part.
   */
  addHeader(text: string, type: HeaderFooterType = "default"): string {
    const partName = this.allocateHeaderPartName();
    const xml = buildHeaderXml(text);
    this.pkg.addPart({
      name: partName,
      contentType: WML_CONTENT_TYPES.header,
      data: new TextEncoder().encode(xml),
    });
    const docRels = this.pkg.partRelationships(this.partName);
    const rel = docRels.add({
      type: WML_RELATIONSHIPS.header,
      target: this.targetRelativeToDoc(partName),
    });
    const sectPr = this.ensureBodySectPr();
    addSectPrHeaderRef(sectPr, type, rel.id);
    this.dirty = true;
    return rel.id;
  }

  /** Like {@link addHeader} but for a footer. */
  addFooter(text: string, type: HeaderFooterType = "default"): string {
    const partName = this.allocateFooterPartName();
    const xml = buildFooterXml(text);
    this.pkg.addPart({
      name: partName,
      contentType: WML_CONTENT_TYPES.footer,
      data: new TextEncoder().encode(xml),
    });
    const docRels = this.pkg.partRelationships(this.partName);
    const rel = docRels.add({
      type: WML_RELATIONSHIPS.footer,
      target: this.targetRelativeToDoc(partName),
    });
    const sectPr = this.ensureBodySectPr();
    addSectPrFooterRef(sectPr, type, rel.id);
    this.dirty = true;
    return rel.id;
  }

  /** Set the page size on the body's trailing `<w:sectPr>`. */
  setPageSize(size: PageSize): void {
    const sectPr = this.ensureBodySectPr();
    setSectPrPageSize(sectPr, size);
    this.dirty = true;
  }

  /** Set the page margins on the body's trailing `<w:sectPr>`. */
  setPageMargins(margins: PageMargins): void {
    const sectPr = this.ensureBodySectPr();
    setSectPrPageMargins(sectPr, margins);
    this.dirty = true;
  }

  /** Switch the page orientation (portrait/landscape) on the body sectPr. */
  setPageOrientation(orientation: "portrait" | "landscape"): void {
    const sectPr = this.ensureBodySectPr();
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
    this.dirty = true;
  }

  private ensureBodySectPr(): XmlElement {
    if (this.docModel.body.sectPr) return this.docModel.body.sectPr;
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
    this.docModel.body.sectPr = sectPr;
    return sectPr;
  }

  private allocateHeaderPartName(): string {
    let n = 1;
    while (this.pkg.hasPart(`/word/header${n}.xml`)) n++;
    return `/word/header${n}.xml`;
  }

  private allocateFooterPartName(): string {
    let n = 1;
    while (this.pkg.hasPart(`/word/footer${n}.xml`)) n++;
    return `/word/footer${n}.xml`;
  }

  private targetRelativeToDoc(partName: string): string {
    if (partName.startsWith("/word/")) return partName.slice("/word/".length);
    return partName.startsWith("/") ? partName.slice(1) : partName;
  }

  /**
   * Serialize to a `Blob` with the `.docx` MIME type. Works in any
   * environment where `Blob` is a global (modern browsers and Node 18+).
   */
  toBlob(): Blob {
    const bytes = this.toUint8Array();
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

  /** Serialize the package back to `.docx` bytes. */
  toUint8Array(): Uint8Array {
    if (this.dirty) {
      this.flushDocument();
      this.dirty = false;
    }
    if (this.stylesDirty && this.stylesCache) {
      this.flushStyles(this.stylesCache);
      this.stylesDirty = false;
    }
    if (this.numberingDirty && this.numberingCache) {
      this.flushNumbering(this.numberingCache);
      this.numberingDirty = false;
    }
    if (this.commentsDirty && this.commentsCache) {
      this.flushComments(this.commentsCache);
      this.commentsDirty = false;
    }
    if (this.footnotesDirty && this.footnotesCache) {
      this.flushNotes(this.footnotesCache, FOOTNOTES_PART_NAME, "footnotes");
      this.footnotesDirty = false;
    }
    if (this.endnotesDirty && this.endnotesCache) {
      this.flushNotes(this.endnotesCache, ENDNOTES_PART_NAME, "endnotes");
      this.endnotesDirty = false;
    }
    return this.pkg.write();
  }

  private flushNotes(
    part: WmlFootnotesPart,
    partName: string,
    rootLocal: "footnotes" | "endnotes",
  ): void {
    const xmlDoc = writeFootnotesPart(part, rootLocal);
    const xmlText = serializeXml(xmlDoc);
    const xmlPart = this.pkg.getPart(partName);
    if (!xmlPart) return;
    xmlPart.data = new TextEncoder().encode(xmlText);
  }

  private flushComments(part: WmlCommentsPart): void {
    const xmlDoc = writeCommentsPart(part);
    const xmlText = serializeXml(xmlDoc);
    const commentsPart = this.pkg.getPart(COMMENTS_PART_NAME);
    if (!commentsPart) return;
    commentsPart.data = new TextEncoder().encode(xmlText);
  }

  private flushNumbering(part: WmlNumberingPart): void {
    const xmlDoc = writeNumberingPart(part);
    const xmlText = serializeXml(xmlDoc);
    const numberingPart = this.pkg.getPart(NUMBERING_PART_NAME);
    if (!numberingPart) return;
    numberingPart.data = new TextEncoder().encode(xmlText);
  }

  private flushDocument(): void {
    const xmlDoc = writeWmlDocument(this.docModel);
    const xmlText = serializeXml(xmlDoc);
    const part = this.pkg.getPart(this.partName);
    if (!part) {
      throw new Error(`Lost document part: ${this.partName}`);
    }
    part.data = new TextEncoder().encode(xmlText);
  }

  private flushStyles(stylesPart: WmlStylesPart): void {
    const xmlDoc = writeStylesPart(stylesPart);
    const xmlText = serializeXml(xmlDoc);
    const part = this.pkg.getPart(STYLES_PART_NAME);
    if (!part) return;
    part.data = new TextEncoder().encode(xmlText);
  }
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
  for (const rel of pkg.packageRelationships.byType(WML_RELATIONSHIPS.officeDocument)) {
    if (rel.targetMode !== "Internal") continue;
    const partName = normalizePartName(rel.target.startsWith("/") ? rel.target : `/${rel.target}`);
    if (partName === CONTENT_TYPES_PART_NAME) continue;
    const part = pkg.getPart(partName);
    if (part) return part;
  }
  return pkg.getPart(DOCUMENT_PART_FALLBACK);
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
