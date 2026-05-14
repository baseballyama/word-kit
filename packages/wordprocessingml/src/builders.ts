import type { XmlAttr, XmlElement } from "@word-kit/ooxml-xml";
import { WML_NS } from "./namespaces.js";
import type {
  WmlParagraph,
  WmlRun,
  WmlRunPiece,
  WmlTable,
  WmlTableCell,
  WmlTableRow,
} from "./types.js";

/**
 * Construct a `<w:t>`-bearing run containing the given plain text.
 * `xml:space="preserve"` is applied automatically when the text has
 * leading or trailing whitespace.
 */
export function buildTextRun(text: string): WmlRun {
  const piece: WmlRunPiece = {
    kind: "text",
    value: text,
    preserveSpace: /^\s|\s$/.test(text),
  };
  return { kind: "run", pieces: [piece], extras: [] };
}

/** Construct a paragraph with a single text run. */
export function buildTextParagraph(text: string): WmlParagraph {
  return { kind: "paragraph", children: [buildTextRun(text)], extras: [] };
}

export interface BuildTableOptions {
  /**
   * Total table width in twips (1/20 of a point, 1/1440 of an inch).
   * Distributed evenly across columns. Defaults to 9000 twips (~6.25").
   */
  readonly totalWidthTwips?: number;
}

/**
 * Construct a basic table from a row-major matrix of strings. Each cell
 * becomes a single paragraph with one run containing the supplied text.
 * The number of columns is inferred from the longest row; shorter rows are
 * padded with empty cells.
 */
export function buildTextTable(
  rows: ReadonlyArray<ReadonlyArray<string>>,
  options: BuildTableOptions = {},
): WmlTable {
  const colCount = rows.reduce((max, row) => Math.max(max, row.length), 0);
  const totalWidth = options.totalWidthTwips ?? 9000;
  const colWidth = colCount > 0 ? Math.floor(totalWidth / colCount) : totalWidth;
  const tblGrid: XmlElement = {
    kind: "element",
    name: { uri: WML_NS, local: "tblGrid", prefix: "w" },
    attrs: [],
    children: Array.from({ length: colCount }, () => ({
      kind: "element" as const,
      name: { uri: WML_NS, local: "gridCol", prefix: "w" },
      attrs: [wmlAttr("w", String(colWidth))],
      children: [],
      xmlSpace: "default" as const,
      selfClosing: true,
    })),
    xmlSpace: "default",
    selfClosing: false,
  };
  const tblPr: XmlElement = {
    kind: "element",
    name: { uri: WML_NS, local: "tblPr", prefix: "w" },
    attrs: [],
    children: [
      {
        kind: "element",
        name: { uri: WML_NS, local: "tblW", prefix: "w" },
        attrs: [wmlAttr("w", String(totalWidth)), wmlAttr("type", "dxa")],
        children: [],
        xmlSpace: "default",
        selfClosing: true,
      },
      {
        kind: "element",
        name: { uri: WML_NS, local: "tblLook", prefix: "w" },
        attrs: [
          wmlAttr("val", "04A0"),
          wmlAttr("firstRow", "1"),
          wmlAttr("lastRow", "0"),
          wmlAttr("firstColumn", "1"),
          wmlAttr("lastColumn", "0"),
          wmlAttr("noHBand", "0"),
          wmlAttr("noVBand", "1"),
        ],
        children: [],
        xmlSpace: "default",
        selfClosing: true,
      },
    ],
    xmlSpace: "default",
    selfClosing: false,
  };
  const tableRows: WmlTableRow[] = rows.map((rowTexts): WmlTableRow => {
    const cells: WmlTableCell[] = [];
    for (let c = 0; c < colCount; c++) {
      const text = rowTexts[c] ?? "";
      cells.push({
        tcPr: cellPropertiesWithWidth(colWidth),
        paragraphs: [buildTextParagraph(text)],
        extras: [],
      });
    }
    return { cells, extras: [] };
  });
  return {
    kind: "table",
    tblPr,
    tblGrid,
    rows: tableRows,
    extras: [],
  };
}

function cellPropertiesWithWidth(widthTwips: number): XmlElement {
  return {
    kind: "element",
    name: { uri: WML_NS, local: "tcPr", prefix: "w" },
    attrs: [],
    children: [
      {
        kind: "element",
        name: { uri: WML_NS, local: "tcW", prefix: "w" },
        attrs: [wmlAttr("w", String(widthTwips)), wmlAttr("type", "dxa")],
        children: [],
        xmlSpace: "default",
        selfClosing: true,
      },
    ],
    xmlSpace: "default",
    selfClosing: false,
  };
}

export interface RunFormatting {
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly strike?: boolean;
  readonly underline?: "single" | "double" | "thick" | "dotted" | "wave" | "none";
  /** Hex RGB without leading `#`. */
  readonly color?: string;
  /** Hex RGB highlight color. */
  readonly highlight?: string;
  /** Font size in half-points (e.g. 24 = 12pt). */
  readonly fontSizeHalfPoints?: number;
  /** Font family applied to ASCII / hAnsi runs. */
  readonly font?: string;
  /** Font family applied to East Asian (CJK) text. */
  readonly fontEastAsia?: string;
}

/** Append a styled text run to an existing paragraph and return it. */
export function appendTextRun(
  paragraph: WmlParagraph,
  text: string,
  formatting: RunFormatting = {},
): WmlRun {
  const piece: WmlRunPiece = {
    kind: "text",
    value: text,
    preserveSpace: /^\s|\s$/.test(text),
  };
  const rPrChildren: XmlElement[] = [];
  if (formatting.font || formatting.fontEastAsia) {
    const attrs: XmlAttr[] = [];
    if (formatting.font) {
      attrs.push(wmlAttr("ascii", formatting.font));
      attrs.push(wmlAttr("hAnsi", formatting.font));
    }
    if (formatting.fontEastAsia) attrs.push(wmlAttr("eastAsia", formatting.fontEastAsia));
    rPrChildren.push(wmlEmpty("rFonts", attrs));
  }
  if (formatting.bold) rPrChildren.push(wmlEmpty("b", []));
  if (formatting.italic) rPrChildren.push(wmlEmpty("i", []));
  if (formatting.strike) rPrChildren.push(wmlEmpty("strike", []));
  if (formatting.underline) rPrChildren.push(wmlEmpty("u", [wmlAttr("val", formatting.underline)]));
  if (formatting.color) rPrChildren.push(wmlEmpty("color", [wmlAttr("val", formatting.color)]));
  if (formatting.highlight)
    rPrChildren.push(wmlEmpty("highlight", [wmlAttr("val", formatting.highlight)]));
  if (formatting.fontSizeHalfPoints !== undefined) {
    rPrChildren.push(wmlEmpty("sz", [wmlAttr("val", String(formatting.fontSizeHalfPoints))]));
    rPrChildren.push(wmlEmpty("szCs", [wmlAttr("val", String(formatting.fontSizeHalfPoints))]));
  }
  const run: WmlRun = rPrChildren.length
    ? {
        kind: "run",
        rPr: {
          kind: "element",
          name: { uri: WML_NS, local: "rPr", prefix: "w" },
          attrs: [],
          children: rPrChildren,
          xmlSpace: "default",
          selfClosing: false,
        },
        pieces: [piece],
        extras: [],
      }
    : { kind: "run", pieces: [piece], extras: [] };
  paragraph.children.push(run);
  return run;
}

export type ParagraphAlignment = "left" | "center" | "right" | "both" | "distribute";

/** Set the paragraph's `<w:jc w:val="..."/>` justification. */
export function setParagraphAlignment(p: WmlParagraph, alignment: ParagraphAlignment): void {
  const pPr = ensurePPr(p);
  const idx = pPr.children.findIndex(
    (c) => c.kind === "element" && c.name.uri === WML_NS && c.name.local === "jc",
  );
  const newEl = wmlEmpty("jc", [wmlAttr("val", alignment)]);
  if (idx >= 0) (pPr.children as XmlElement[])[idx] = newEl;
  else (pPr.children as XmlElement[]).push(newEl);
}

export interface ParagraphIndent {
  /** Left indent in twips. */
  readonly left?: number;
  /** Right indent in twips. */
  readonly right?: number;
  /** First-line indent in twips. */
  readonly firstLine?: number;
  /** Hanging indent in twips. */
  readonly hanging?: number;
}

/** Set `<w:ind>` on the paragraph's pPr. */
export function setParagraphIndent(p: WmlParagraph, indent: ParagraphIndent): void {
  const pPr = ensurePPr(p);
  const attrs: XmlAttr[] = [];
  if (indent.left !== undefined) attrs.push(wmlAttr("left", String(indent.left)));
  if (indent.right !== undefined) attrs.push(wmlAttr("right", String(indent.right)));
  if (indent.firstLine !== undefined) attrs.push(wmlAttr("firstLine", String(indent.firstLine)));
  if (indent.hanging !== undefined) attrs.push(wmlAttr("hanging", String(indent.hanging)));
  const idx = pPr.children.findIndex(
    (c) => c.kind === "element" && c.name.uri === WML_NS && c.name.local === "ind",
  );
  const newEl = wmlEmpty("ind", attrs);
  if (idx >= 0) (pPr.children as XmlElement[])[idx] = newEl;
  else (pPr.children as XmlElement[]).push(newEl);
}

export interface ParagraphSpacing {
  /** Space before the paragraph in twips. */
  readonly before?: number;
  /** Space after the paragraph in twips. */
  readonly after?: number;
  /** Line spacing value in twips. */
  readonly line?: number;
  /** Line spacing rule: `"auto"`, `"exact"`, or `"atLeast"`. */
  readonly lineRule?: "auto" | "exact" | "atLeast";
}

/** Set `<w:spacing>` on the paragraph's pPr. */
export function setParagraphSpacing(p: WmlParagraph, spacing: ParagraphSpacing): void {
  const pPr = ensurePPr(p);
  const attrs: XmlAttr[] = [];
  if (spacing.before !== undefined) attrs.push(wmlAttr("before", String(spacing.before)));
  if (spacing.after !== undefined) attrs.push(wmlAttr("after", String(spacing.after)));
  if (spacing.line !== undefined) attrs.push(wmlAttr("line", String(spacing.line)));
  if (spacing.lineRule !== undefined) attrs.push(wmlAttr("lineRule", spacing.lineRule));
  const idx = pPr.children.findIndex(
    (c) => c.kind === "element" && c.name.uri === WML_NS && c.name.local === "spacing",
  );
  const newEl = wmlEmpty("spacing", attrs);
  if (idx >= 0) (pPr.children as XmlElement[])[idx] = newEl;
  else (pPr.children as XmlElement[]).push(newEl);
}

function ensurePPr(p: WmlParagraph): XmlElement {
  if (p.pPr) return p.pPr;
  const pPr: XmlElement = {
    kind: "element",
    name: { uri: WML_NS, local: "pPr", prefix: "w" },
    attrs: [],
    children: [],
    xmlSpace: "default",
    selfClosing: false,
  };
  p.pPr = pPr;
  return pPr;
}

/** Read `<w:pStyle w:val>` from a paragraph's pPr, or `undefined`. */
export function getParagraphStyle(p: WmlParagraph): string | undefined {
  return refChildVal(p.pPr, "pStyle");
}

/** Read `<w:jc w:val>` from a paragraph's pPr as a typed value, or `undefined`. */
export function getParagraphAlignment(p: WmlParagraph): ParagraphAlignment | undefined {
  const v = refChildVal(p.pPr, "jc");
  if (v === "left" || v === "center" || v === "right" || v === "both" || v === "distribute")
    return v;
  return undefined;
}

/** Read the numbering reference (`<w:numPr>` → `numId` / `ilvl`) if present. */
export function getParagraphNumbering(
  p: WmlParagraph,
): { numId: number; ilvl: number } | undefined {
  if (!p.pPr) return undefined;
  const numPr = p.pPr.children.find(
    (c) => c.kind === "element" && c.name.uri === WML_NS && c.name.local === "numPr",
  );
  if (!numPr || numPr.kind !== "element") return undefined;
  const numIdEl = numPr.children.find(
    (c) => c.kind === "element" && c.name.uri === WML_NS && c.name.local === "numId",
  );
  const ilvlEl = numPr.children.find(
    (c) => c.kind === "element" && c.name.uri === WML_NS && c.name.local === "ilvl",
  );
  if (!numIdEl || numIdEl.kind !== "element") return undefined;
  const numIdAttr = numIdEl.attrs.find((a) => a.name.uri === WML_NS && a.name.local === "val");
  if (!numIdAttr) return undefined;
  const numIdVal = Number.parseInt(numIdAttr.value, 10);
  const ilvlAttr =
    ilvlEl?.kind === "element"
      ? ilvlEl.attrs.find((a) => a.name.uri === WML_NS && a.name.local === "val")
      : undefined;
  const ilvlVal = ilvlAttr ? Number.parseInt(ilvlAttr.value, 10) : 0;
  if (!Number.isFinite(numIdVal)) return undefined;
  return { numId: numIdVal, ilvl: Number.isFinite(ilvlVal) ? ilvlVal : 0 };
}

/**
 * Replace a paragraph's content with a single styled text run. Existing
 * runs and inline children are removed; existing pPr is kept.
 */
export function setParagraphText(
  p: WmlParagraph,
  text: string,
  formatting: RunFormatting = {},
): void {
  p.children = [];
  appendTextRun(p, text, formatting);
}

function refChildVal(parent: XmlElement | undefined, local: string): string | undefined {
  if (!parent) return undefined;
  for (const c of parent.children) {
    if (c.kind === "element" && c.name.uri === WML_NS && c.name.local === local) {
      const attr = c.attrs.find((a) => a.name.uri === WML_NS && a.name.local === "val");
      return attr?.value;
    }
  }
  return undefined;
}

function wmlEmpty(local: string, attrs: XmlAttr[]): XmlElement {
  return {
    kind: "element",
    name: { uri: WML_NS, local, prefix: "w" },
    attrs,
    children: [],
    xmlSpace: "default",
    selfClosing: true,
  };
}

function wmlAttr(local: string, value: string): XmlAttr {
  return {
    name: { uri: WML_NS, local, prefix: "w" },
    value,
    isNamespaceDecl: false,
  };
}
