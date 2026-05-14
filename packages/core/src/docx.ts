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
  addSectPrFooterRef,
  addSectPrHeaderRef,
  buildAbstractNum,
  buildComment,
  buildCommentRangeEnd,
  buildCommentRangeStart,
  buildCommentReferenceRun,
  buildFooterXml,
  buildHeaderXml,
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
  findStyle,
  findText,
  type HeaderFooterType,
  MINIMAL_STYLES_XML,
  numAbstractRef,
  numId as readNumId,
  parseCommentsPart,
  type PageMargins,
  PAGE_SIZE_LETTER,
  type PageSize,
  parseNumberingPart,
  parseStylesPart,
  parseWmlDocument,
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
import type { XmlElement } from "@word-kit/ooxml-xml";

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

  /** Search the document for all matches of `query`. */
  findText(query: string | RegExp): TextMatch[] {
    return findText(this.docModel, query);
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

  /** Visible text of the document. Paragraphs are joined with `\n`. */
  get text(): string {
    return documentText(this.docModel);
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
    return this.pkg.write();
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
