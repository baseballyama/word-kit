import type { XmlAttr, XmlElement, XmlNode } from "@word-kit/ooxml-xml";

/**
 * Top-level Word document model. Round-trip stable: every element that was
 * recognized is captured as a typed node, every unknown element is captured
 * as a {@link WmlRawBlock}/{@link WmlRawInline} so nothing is lost on save.
 */
export interface WmlDocument {
  /**
   * Attributes (typically `xmlns:*` namespace declarations and
   * `mc:Ignorable`) of the root `<w:document>` element.
   */
  readonly rootAttrs: readonly XmlAttr[];
  readonly body: WmlBody;
  /**
   * Any children of `<w:document>` other than `<w:body>` (extremely rare;
   * preserved verbatim).
   */
  readonly extras: readonly PassThrough[];
}

export interface WmlBody {
  blocks: WmlBlock[];
  /** Final `<w:sectPr>` at the end of the body, if any. Kept as raw XML. */
  sectPr?: XmlElement;
  /** Unknown body children, with their original child-list index. */
  extras: PassThrough[];
}

export type WmlBlock = WmlParagraph | WmlTable | WmlRawBlock;

export interface WmlTable {
  readonly kind: "table";
  /** `<w:tblPr>` retained as raw XML; typed property access lands later. */
  tblPr?: XmlElement;
  /** `<w:tblGrid>` retained as raw XML. */
  tblGrid?: XmlElement;
  rows: WmlTableRow[];
  /** Unknown children of `<w:tbl>`, with their original position. */
  extras: PassThrough[];
}

export interface WmlTableRow {
  /** `<w:trPr>` retained as raw XML. */
  trPr?: XmlElement;
  cells: WmlTableCell[];
  /** Unknown children of `<w:tr>`, with their original position. */
  extras: PassThrough[];
}

export interface WmlTableCell {
  /** `<w:tcPr>` retained as raw XML. */
  tcPr?: XmlElement;
  /** Paragraphs inside the cell. Nested tables / other content live in extras. */
  paragraphs: WmlParagraph[];
  /** Unknown children of `<w:tc>`, with their original position. */
  extras: PassThrough[];
}

export interface WmlParagraph {
  readonly kind: "paragraph";
  /**
   * `<w:pPr>` retained as the raw element. Typed property access is added
   * incrementally in higher milestones.
   */
  pPr?: XmlElement;
  children: WmlInline[];
  /** Children of `<w:p>` that we don't recognize (with original position). */
  extras: PassThrough[];
}

export interface WmlRawBlock {
  readonly kind: "raw";
  /** Original element verbatim (e.g. `<w:tbl>`, `<w:sdt>`, `<w:sectPr>`). */
  node: XmlElement;
}

export type WmlInline = WmlRun | WmlRawInline;

export interface WmlRun {
  readonly kind: "run";
  rPr?: XmlElement;
  pieces: WmlRunPiece[];
  /** Unknown children inside `<w:r>` (preserves original position). */
  extras: PassThrough[];
}

export interface WmlRawInline {
  readonly kind: "raw";
  node: XmlElement;
}

export type WmlRunPiece =
  | { kind: "text"; value: string; preserveSpace: boolean }
  | { kind: "delText"; value: string; preserveSpace: boolean }
  | { kind: "tab" }
  | {
      kind: "break";
      breakType?: "page" | "column" | "textWrapping";
      clear?: "none" | "left" | "right" | "all";
    }
  | { kind: "noBreakHyphen" }
  | { kind: "softHyphen" }
  | { kind: "lastRenderedPageBreak" }
  | { kind: "instrText"; value: string; preserveSpace: boolean }
  | { kind: "delInstrText"; value: string; preserveSpace: boolean }
  | { kind: "fieldChar"; charType: "begin" | "separate" | "end"; raw: XmlElement }
  | { kind: "symbol"; font: string; char: string }
  | { kind: "drawing"; node: XmlElement }
  | { kind: "pict"; node: XmlElement }
  | { kind: "raw"; node: XmlElement };

/**
 * A child node that the typed AST did not recognize, paired with its
 * original index in the parent's `children` array. On serialize, the writer
 * splices these back into the same position.
 */
export interface PassThrough {
  readonly slot: number;
  readonly node: XmlNode;
}
