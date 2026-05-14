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
  buildInlineDrawing,
  buildTextTable,
  type BuildTableOptions,
  documentText,
  extensionForImageContentType,
  findText,
  parseWmlDocument,
  replaceText,
  sniffImageContentType,
  type TextMatch,
  WML_NS,
  WML_RELATIONSHIPS,
  type WmlBlock,
  type WmlDocument,
  type WmlParagraph,
  type WmlRun,
  type WmlRunPiece,
  type WmlTable,
  writeWmlDocument,
} from "@word-kit/wml";

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
export class Docx {
  private readonly pkg: OpcPackage;
  private readonly docModel: WmlDocument;
  private readonly partName: string;
  private dirty = false;

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

  /** Serialize the package back to `.docx` bytes. */
  toUint8Array(): Uint8Array {
    if (this.dirty) {
      this.flushDocument();
      this.dirty = false;
    }
    return this.pkg.write();
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
