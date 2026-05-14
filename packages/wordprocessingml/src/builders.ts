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

function wmlAttr(local: string, value: string): XmlAttr {
  return {
    name: { uri: WML_NS, local, prefix: "w" },
    value,
    isNamespaceDecl: false,
  };
}
