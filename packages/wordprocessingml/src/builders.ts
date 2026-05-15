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

export type TableBorderStyle = "none" | "single" | "double" | "dashed" | "dotted" | "thick";

export interface TableBordersOptions {
  /** Stroke style — defaults to `"single"`. Pass `"none"` to clear borders. */
  readonly style?: TableBorderStyle;
  /** Stroke width in eighths of a point — Word's `w:sz`. Defaults to 4 (½ pt). */
  readonly sizeEighthsOfPoint?: number;
  /** Hex RGB without a leading `#`. Defaults to `"auto"`. */
  readonly color?: string;
  /**
   * When `true` (default), also draws the interior gridlines
   * (`<w:insideH/>` and `<w:insideV/>`). Disable for an outer-only frame.
   */
  readonly inside?: boolean;
}

/**
 * Set uniform borders on every side of `table`. Replaces any existing
 * `<w:tblBorders>` block; the rest of `<w:tblPr>` (width, look, etc.)
 * is preserved.
 *
 * Example: a half-point single-line frame around the whole table —
 *   `setTableBorders(table, {});`
 *
 * Example: thick double border, outer only —
 *   `setTableBorders(table, { style: "double", sizeEighthsOfPoint: 12, inside: false });`
 *
 * Example: no borders at all —
 *   `setTableBorders(table, { style: "none" });`
 */
export function setTableBorders(table: WmlTable, options: TableBordersOptions = {}): void {
  const style = options.style ?? "single";
  const sz = String(options.sizeEighthsOfPoint ?? 4);
  const color = options.color ?? "auto";
  const inside = options.inside ?? true;

  const tblPr = table.tblPr ?? {
    kind: "element",
    name: { uri: WML_NS, local: "tblPr", prefix: "w" },
    attrs: [],
    children: [],
    xmlSpace: "default",
    selfClosing: false,
  };
  const children = tblPr.children as XmlElement[];
  // Drop any existing tblBorders before re-adding ours.
  for (let i = children.length - 1; i >= 0; i--) {
    const c = children[i];
    if (c && c.kind === "element" && c.name.uri === WML_NS && c.name.local === "tblBorders") {
      children.splice(i, 1);
    }
  }
  const borderAttrs = (): XmlAttr[] => [
    wmlAttr("val", style),
    wmlAttr("sz", sz),
    wmlAttr("space", "0"),
    wmlAttr("color", color),
  ];
  const sides = ["top", "left", "bottom", "right"];
  if (inside) sides.push("insideH", "insideV");
  const bordersEl: XmlElement = {
    kind: "element",
    name: { uri: WML_NS, local: "tblBorders", prefix: "w" },
    attrs: [],
    children: sides.map((s) => wmlEmpty(s, borderAttrs())),
    xmlSpace: "default",
    selfClosing: false,
  };
  // tblBorders must appear after tblW per the spec; we insert as the first
  // child after tblW if it's present, else just push to the end.
  const tblWIdx = children.findIndex((c) => c.kind === "element" && c.name.local === "tblW");
  if (tblWIdx >= 0) children.splice(tblWIdx + 1, 0, bordersEl);
  else children.unshift(bordersEl);

  if (!table.tblPr) table.tblPr = tblPr;
}

export interface TableCellShadingOptions {
  /**
   * Hex RGB fill colour without a leading `#` (e.g. `"DDDDDD"`). Defaults
   * to `"auto"` which lets Word pick from the active style.
   */
  readonly fill?: string;
  /**
   * Pattern overlaid on the fill. `"clear"` (default) is a flat fill;
   * `"solid"` is equivalent to `"clear"` with the foreground colour
   * forced; the rest are crosshatches commonly used for table emphasis.
   */
  readonly pattern?:
    | "clear"
    | "solid"
    | "horzStripe"
    | "vertStripe"
    | "diagStripe"
    | "diagCross"
    | "thinHorzStripe"
    | "thinVertStripe";
  /**
   * Pattern stroke colour (only meaningful when `pattern !== "clear"`).
   * Defaults to `"auto"`.
   */
  readonly color?: string;
}

/**
 * Apply cell shading to a table cell. Replaces any existing `<w:shd>` on
 * the cell's `<w:tcPr>`; other tcPr children (width, vertical alignment,
 * cell margins) are preserved.
 *
 * Example — light grey header cell:
 *   `setTableCellShading(headerCell, { fill: "E0E0E0" });`
 */
export function setTableCellShading(cell: WmlTableCell, options: TableCellShadingOptions): void {
  const fill = options.fill ?? "auto";
  const pattern = options.pattern ?? "clear";
  const color = options.color ?? "auto";
  const tcPr = cell.tcPr ?? {
    kind: "element",
    name: { uri: WML_NS, local: "tcPr", prefix: "w" },
    attrs: [],
    children: [],
    xmlSpace: "default",
    selfClosing: false,
  };
  const children = tcPr.children as XmlElement[];
  for (let i = children.length - 1; i >= 0; i--) {
    const c = children[i];
    if (c && c.kind === "element" && c.name.uri === WML_NS && c.name.local === "shd") {
      children.splice(i, 1);
    }
  }
  children.push(
    wmlEmpty("shd", [wmlAttr("val", pattern), wmlAttr("color", color), wmlAttr("fill", fill)]),
  );
  if (!cell.tcPr) cell.tcPr = tcPr;
}

/**
 * Vertical alignment for content inside a table cell. Maps to `<w:vAlign>`.
 */
export type TableCellVerticalAlign = "top" | "center" | "bottom";

/** Set the vertical alignment of content inside a cell. */
export function setTableCellVerticalAlign(cell: WmlTableCell, align: TableCellVerticalAlign): void {
  const tcPr = cell.tcPr ?? {
    kind: "element",
    name: { uri: WML_NS, local: "tcPr", prefix: "w" },
    attrs: [],
    children: [],
    xmlSpace: "default",
    selfClosing: false,
  };
  const children = tcPr.children as XmlElement[];
  for (let i = children.length - 1; i >= 0; i--) {
    const c = children[i];
    if (c && c.kind === "element" && c.name.uri === WML_NS && c.name.local === "vAlign") {
      children.splice(i, 1);
    }
  }
  children.push(wmlEmpty("vAlign", [wmlAttr("val", align)]));
  if (!cell.tcPr) cell.tcPr = tcPr;
}

export type TableRowHeightRule = "atLeast" | "exact" | "auto";

/**
 * Set an explicit row height in twips. `rule` defaults to `"atLeast"`,
 * meaning the row grows past `heightTwips` if its content overflows.
 * `"exact"` clips overflowing content; `"auto"` lets Word resize freely
 * and ignores `heightTwips` (Word's UI calls this "automatic").
 */
export function setTableRowHeight(
  row: WmlTableRow,
  heightTwips: number,
  rule: TableRowHeightRule = "atLeast",
): void {
  const trPr = row.trPr ?? {
    kind: "element",
    name: { uri: WML_NS, local: "trPr", prefix: "w" },
    attrs: [],
    children: [],
    xmlSpace: "default",
    selfClosing: false,
  };
  const children = trPr.children as XmlElement[];
  for (let i = children.length - 1; i >= 0; i--) {
    const c = children[i];
    if (c && c.kind === "element" && c.name.uri === WML_NS && c.name.local === "trHeight") {
      children.splice(i, 1);
    }
  }
  children.push(
    wmlEmpty("trHeight", [wmlAttr("val", String(heightTwips)), wmlAttr("hRule", rule)]),
  );
  if (!row.trPr) row.trPr = trPr;
}

/**
 * Mark a table row as a *header row* — its content is repeated at the
 * top of every page when the table breaks across pages. Word's UI
 * equivalent is "Repeat as header row at the top of each page".
 *
 * Pass `false` to clear the marker.
 */
export function setTableRowAsHeader(row: WmlTableRow, isHeader = true): void {
  const trPr = row.trPr ?? {
    kind: "element",
    name: { uri: WML_NS, local: "trPr", prefix: "w" },
    attrs: [],
    children: [],
    xmlSpace: "default",
    selfClosing: false,
  };
  const children = trPr.children as XmlElement[];
  for (let i = children.length - 1; i >= 0; i--) {
    const c = children[i];
    if (c && c.kind === "element" && c.name.uri === WML_NS && c.name.local === "tblHeader") {
      children.splice(i, 1);
    }
  }
  if (isHeader) children.push(wmlEmpty("tblHeader", []));
  if (!row.trPr) row.trPr = trPr;
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

/**
 * Apply formatting to an existing run, replacing any matching `<w:rPr>`
 * children. Properties left `undefined` are not touched, so this is a
 * patch — call `clearRunFormat` first if you want a full reset.
 *
 * Boolean flags (`bold`, `italic`, `strike`) accept `false` to remove
 * the corresponding `<w:b/>` / `<w:i/>` / `<w:strike/>` element.
 * `underline: "none"` removes `<w:u/>`. Setting a value to a string or
 * number replaces (or inserts) the relevant rPr child while preserving
 * the run's other rPr children, including formatting the library
 * doesn't yet model.
 */
export function setRunFormat(run: WmlRun, formatting: RunFormatting): void {
  if (
    formatting.bold === undefined &&
    formatting.italic === undefined &&
    formatting.strike === undefined &&
    formatting.underline === undefined &&
    formatting.color === undefined &&
    formatting.highlight === undefined &&
    formatting.fontSizeHalfPoints === undefined &&
    formatting.font === undefined &&
    formatting.fontEastAsia === undefined
  ) {
    return;
  }
  const rPr: XmlElement = run.rPr ?? {
    kind: "element",
    name: { uri: WML_NS, local: "rPr", prefix: "w" },
    attrs: [],
    children: [],
    xmlSpace: "default",
    selfClosing: false,
  };
  const children = rPr.children as XmlElement[];

  const removeLocal = (local: string): void => {
    for (let i = children.length - 1; i >= 0; i--) {
      const c = children[i];
      if (c && c.kind === "element" && c.name.uri === WML_NS && c.name.local === local) {
        children.splice(i, 1);
      }
    }
  };
  const setEmpty = (local: string, attrs: XmlAttr[]): void => {
    removeLocal(local);
    children.push(wmlEmpty(local, attrs));
  };

  if (formatting.bold !== undefined) {
    if (formatting.bold) setEmpty("b", []);
    else removeLocal("b");
  }
  if (formatting.italic !== undefined) {
    if (formatting.italic) setEmpty("i", []);
    else removeLocal("i");
  }
  if (formatting.strike !== undefined) {
    if (formatting.strike) setEmpty("strike", []);
    else removeLocal("strike");
  }
  if (formatting.underline !== undefined) {
    if (formatting.underline === "none") removeLocal("u");
    else setEmpty("u", [wmlAttr("val", formatting.underline)]);
  }
  if (formatting.color !== undefined) {
    setEmpty("color", [wmlAttr("val", formatting.color)]);
  }
  if (formatting.highlight !== undefined) {
    setEmpty("highlight", [wmlAttr("val", formatting.highlight)]);
  }
  if (formatting.fontSizeHalfPoints !== undefined) {
    setEmpty("sz", [wmlAttr("val", String(formatting.fontSizeHalfPoints))]);
    setEmpty("szCs", [wmlAttr("val", String(formatting.fontSizeHalfPoints))]);
  }
  if (formatting.font !== undefined || formatting.fontEastAsia !== undefined) {
    // Build/replace <w:rFonts>. Preserve any attrs the caller didn't override.
    const existing = children.find(
      (c): c is XmlElement => c.kind === "element" && c.name.local === "rFonts",
    );
    const attrs: XmlAttr[] = existing ? existing.attrs.slice() : [];
    const setAttr = (local: string, value: string): void => {
      const idx = attrs.findIndex((a) => a.name.local === local);
      const next = wmlAttr(local, value);
      if (idx >= 0) attrs[idx] = next;
      else attrs.push(next);
    };
    if (formatting.font !== undefined) {
      setAttr("ascii", formatting.font);
      setAttr("hAnsi", formatting.font);
    }
    if (formatting.fontEastAsia !== undefined) {
      setAttr("eastAsia", formatting.fontEastAsia);
    }
    removeLocal("rFonts");
    children.push(wmlEmpty("rFonts", attrs));
  }

  if (!run.rPr && children.length > 0) {
    run.rPr = rPr;
  }
}

/**
 * Drop the run's entire `<w:rPr>` block, leaving the run text unformatted
 * (it picks up the document/paragraph default style).
 */
export function clearRunFormat(run: WmlRun): void {
  delete run.rPr;
}

/**
 * Read back the formatting a run currently has on its `<w:rPr>`. Returns a
 * `RunFormatting` with only the keys that are actually present, so callers
 * can round-trip via `setRunFormat(other, getRunFormat(run))`.
 */
export function getRunFormat(run: WmlRun): RunFormatting {
  // Build a writable scratch object; only keys we actually observe will
  // be set, then cast to the readonly `RunFormatting` on return.
  const out: Record<string, unknown> = {};
  if (!run.rPr) return out as RunFormatting;
  for (const child of run.rPr.children) {
    if (child.kind !== "element" || child.name.uri !== WML_NS) continue;
    const local = child.name.local;
    if (local === "b") out.bold = true;
    else if (local === "i") out.italic = true;
    else if (local === "strike") out.strike = true;
    else if (local === "u") {
      const v = child.attrs.find((a) => a.name.local === "val")?.value;
      if (
        v === "single" ||
        v === "double" ||
        v === "thick" ||
        v === "dotted" ||
        v === "wave" ||
        v === "none"
      ) {
        out.underline = v;
      }
    } else if (local === "color") {
      const v = child.attrs.find((a) => a.name.local === "val")?.value;
      if (v !== undefined) out.color = v;
    } else if (local === "highlight") {
      const v = child.attrs.find((a) => a.name.local === "val")?.value;
      if (v !== undefined) out.highlight = v;
    } else if (local === "sz") {
      const v = child.attrs.find((a) => a.name.local === "val")?.value;
      const n = v !== undefined ? Number.parseInt(v, 10) : NaN;
      if (Number.isFinite(n)) out.fontSizeHalfPoints = n;
    } else if (local === "rFonts") {
      const ascii = child.attrs.find((a) => a.name.local === "ascii")?.value;
      const east = child.attrs.find((a) => a.name.local === "eastAsia")?.value;
      if (ascii !== undefined) out.font = ascii;
      if (east !== undefined) out.fontEastAsia = east;
    }
  }
  return out as RunFormatting;
}

/**
 * Collapse runs in `paragraph` that are immediately adjacent and share
 * the same `<w:rPr>`. Only runs whose pieces are entirely text /
 * delText / tab / break (no drawings, instrText, etc.) are eligible —
 * special pieces stop the merge boundary. Other inline children
 * (`<w:hyperlink>`, raw passthroughs, …) interrupt adjacency.
 *
 * Returns the number of merges performed. The merged run is the earlier
 * one — the later run is removed and its pieces appended.
 *
 * Useful after templating has fragmented a paragraph into many tiny
 * runs that share the same formatting (Word's spell-checker is the
 * common culprit).
 */
export function mergeAdjacentRuns(paragraph: WmlParagraph): number {
  let merges = 0;
  const out: typeof paragraph.children = [];
  for (const child of paragraph.children) {
    const prev = out[out.length - 1];
    if (
      child.kind === "run" &&
      prev &&
      prev.kind === "run" &&
      isMergeableRun(child) &&
      isMergeableRun(prev) &&
      sameRPr(prev, child)
    ) {
      prev.pieces.push(...child.pieces);
      merges++;
      continue;
    }
    out.push(child);
  }
  paragraph.children = out;
  return merges;
}

function isSafeRunPiece(p: WmlRunPiece): boolean {
  return p.kind === "text" || p.kind === "delText" || p.kind === "tab" || p.kind === "break";
}

function isMergeableRun(r: WmlRun): boolean {
  return r.extras.length === 0 && r.pieces.every(isSafeRunPiece);
}

function sameRPr(a: WmlRun, b: WmlRun): boolean {
  if (a.rPr === undefined && b.rPr === undefined) return true;
  if (a.rPr === undefined || b.rPr === undefined) return false;
  return elementsEqual(a.rPr, b.rPr);
}

function elementsEqual(a: XmlElement, b: XmlElement): boolean {
  if (a.name.uri !== b.name.uri || a.name.local !== b.name.local) return false;
  if (a.attrs.length !== b.attrs.length) return false;
  // Attribute order isn't semantic, but for rPr comparison Word treats
  // ordering as cosmetic — we compare attrs as an unordered set keyed by
  // namespaced name + value.
  const aKeys = a.attrs.map((x) => `${x.name.uri}|${x.name.local}=${x.value}`).toSorted();
  const bKeys = b.attrs.map((x) => `${x.name.uri}|${x.name.local}=${x.value}`).toSorted();
  for (let i = 0; i < aKeys.length; i++) if (aKeys[i] !== bKeys[i]) return false;
  if (a.children.length !== b.children.length) return false;
  for (let i = 0; i < a.children.length; i++) {
    const ca = a.children[i];
    const cb = b.children[i];
    if (!ca || !cb || ca.kind !== cb.kind) return false;
    if (ca.kind === "element" && cb.kind === "element") {
      if (!elementsEqual(ca, cb)) return false;
    } else if (ca.kind === "text" && cb.kind === "text") {
      if (ca.value !== cb.value) return false;
    } else if (ca.kind === "cdata" && cb.kind === "cdata") {
      if (ca.value !== cb.value) return false;
    } else {
      // Comments / PIs in rPr are exotic enough that we treat them as
      // "not equal" rather than try to compare in detail.
      return false;
    }
  }
  return true;
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

export interface ParagraphBordersOptions {
  /** Stroke style — defaults to `"single"`. Pass `"none"` to clear borders. */
  readonly style?: TableBorderStyle;
  /** Stroke width in eighths of a point. Defaults to 4 (½ pt). */
  readonly sizeEighthsOfPoint?: number;
  /** Hex RGB without a leading `#`. Defaults to `"auto"`. */
  readonly color?: string;
  /**
   * Distance in points (`w:space`) between text and border. Defaults to 4.
   * Valid range is 0–31 per the spec; values outside are clamped silently
   * by Word.
   */
  readonly spacePt?: number;
  /**
   * Sides to apply. Defaults to all four (`top`, `left`, `bottom`, `right`)
   * — pass a subset to draw, eg, only a bottom rule under headings.
   */
  readonly sides?: ReadonlyArray<"top" | "left" | "bottom" | "right">;
}

/**
 * Set borders on a paragraph's pPr. Replaces any existing
 * `<w:pBdr>` block.
 *
 * Common patterns:
 *
 * - Bottom rule under a heading:
 *   `setParagraphBorders(p, { sides: ["bottom"], sizeEighthsOfPoint: 12 });`
 *
 * - Full box around a callout paragraph:
 *   `setParagraphBorders(p, { color: "808080" });`
 *
 * - Clear an existing border block:
 *   `setParagraphBorders(p, { style: "none" });`
 */
export function setParagraphBorders(
  paragraph: WmlParagraph,
  options: ParagraphBordersOptions = {},
): void {
  const pPr = ensurePPr(paragraph);
  const style = options.style ?? "single";
  const sz = String(options.sizeEighthsOfPoint ?? 4);
  const color = options.color ?? "auto";
  const space = String(options.spacePt ?? 4);
  const sides = options.sides ?? ["top", "left", "bottom", "right"];
  const children = pPr.children as XmlElement[];
  for (let i = children.length - 1; i >= 0; i--) {
    const c = children[i];
    if (c && c.kind === "element" && c.name.uri === WML_NS && c.name.local === "pBdr") {
      children.splice(i, 1);
    }
  }
  const sideAttrs = (): XmlAttr[] => [
    wmlAttr("val", style),
    wmlAttr("sz", sz),
    wmlAttr("space", space),
    wmlAttr("color", color),
  ];
  const pBdr: XmlElement = {
    kind: "element",
    name: { uri: WML_NS, local: "pBdr", prefix: "w" },
    attrs: [],
    children: sides.map((s) => wmlEmpty(s, sideAttrs())),
    xmlSpace: "default",
    selfClosing: false,
  };
  // pBdr appears after pStyle and numPr per the schema; insert after the
  // last "structural" pPr child if any are present.
  const tailLocals = new Set(["pStyle", "keepNext", "keepLines", "numPr"]);
  let insertAt = 0;
  for (let i = 0; i < children.length; i++) {
    const c = children[i];
    if (c && c.kind === "element" && tailLocals.has(c.name.local)) insertAt = i + 1;
  }
  children.splice(insertAt, 0, pBdr);
}

export interface ParagraphShadingOptions {
  /** Hex RGB fill colour. Defaults to `"auto"`. */
  readonly fill?: string;
  /** Pattern overlaid on the fill. Defaults to `"clear"` (flat). */
  readonly pattern?:
    | "clear"
    | "solid"
    | "horzStripe"
    | "vertStripe"
    | "diagStripe"
    | "diagCross"
    | "thinHorzStripe"
    | "thinVertStripe";
  /** Pattern stroke colour. Defaults to `"auto"`. */
  readonly color?: string;
}

/** Apply background shading to a paragraph (Word's "highlight" — but applied
 * to the whole paragraph rather than a run). Replaces any existing
 * `<w:shd>` on pPr.
 */
export function setParagraphShading(
  paragraph: WmlParagraph,
  options: ParagraphShadingOptions = {},
): void {
  const pPr = ensurePPr(paragraph);
  const fill = options.fill ?? "auto";
  const pattern = options.pattern ?? "clear";
  const color = options.color ?? "auto";
  const children = pPr.children as XmlElement[];
  for (let i = children.length - 1; i >= 0; i--) {
    const c = children[i];
    if (c && c.kind === "element" && c.name.uri === WML_NS && c.name.local === "shd") {
      children.splice(i, 1);
    }
  }
  children.push(
    wmlEmpty("shd", [wmlAttr("val", pattern), wmlAttr("color", color), wmlAttr("fill", fill)]),
  );
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

/**
 * Set or clear a paragraph's `<w:pStyle w:val="…">`. Pass a styleId
 * (e.g. `"Heading1"`) to apply, or `undefined` to drop the style
 * reference altogether (paragraph falls back to Normal).
 *
 * Note that this only writes the `<w:pStyle>` reference; the styles
 * part is unaffected. Use {@link addStyle} (in `@word-kit/core`) first
 * if you're applying a custom style that doesn't exist yet.
 */
export function setParagraphStyle(p: WmlParagraph, styleId: string | undefined): void {
  if (styleId === undefined) {
    if (!p.pPr) return;
    const children = p.pPr.children as XmlElement[];
    for (let i = children.length - 1; i >= 0; i--) {
      const c = children[i];
      if (c && c.kind === "element" && c.name.uri === WML_NS && c.name.local === "pStyle") {
        children.splice(i, 1);
      }
    }
    return;
  }
  const pPr = ensurePPr(p);
  const children = pPr.children as XmlElement[];
  const idx = children.findIndex(
    (c) => c.kind === "element" && c.name.uri === WML_NS && c.name.local === "pStyle",
  );
  const newEl = wmlEmpty("pStyle", [wmlAttr("val", styleId)]);
  if (idx >= 0) children[idx] = newEl;
  else children.unshift(newEl); // pStyle conventionally appears first.
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

/**
 * Replace the text of a specific cell in a table with `text`. The cell's
 * existing paragraphs are replaced with a single paragraph containing one
 * styled text run. Row / column indices are 0-based; throws if out of
 * range.
 */
export function setTableCellText(
  table: WmlTable,
  row: number,
  col: number,
  text: string,
  formatting: RunFormatting = {},
): void {
  const tableRow = table.rows[row];
  if (!tableRow) {
    throw new Error(
      `setTableCellText: row ${row} is out of range (table has ${table.rows.length})`,
    );
  }
  const cell = tableRow.cells[col];
  if (!cell) {
    throw new Error(
      `setTableCellText: column ${col} is out of range (row ${row} has ${tableRow.cells.length})`,
    );
  }
  const paragraph = buildTextParagraph(text);
  cell.paragraphs = [paragraph];
  if (Object.keys(formatting).length > 0) {
    // appendTextRun already pushed the plain run; we need to wrap with the
    // styled run instead. Rebuild via appendTextRun on a fresh paragraph.
    paragraph.children = [];
    appendTextRun(paragraph, text, formatting);
  }
}

/**
 * Read the plain text of a single table cell (joining its paragraphs with
 * newlines). Throws if `row` / `col` is out of range.
 */
export function getTableCellText(table: WmlTable, row: number, col: number): string {
  const tableRow = table.rows[row];
  if (!tableRow) {
    throw new Error(`getTableCellText: row ${row} is out of range`);
  }
  const cell = tableRow.cells[col];
  if (!cell) {
    throw new Error(`getTableCellText: column ${col} is out of range`);
  }
  return cell.paragraphs
    .map((p) => {
      let acc = "";
      for (const child of p.children) {
        if (child.kind !== "run") continue;
        for (const piece of child.pieces) {
          if (piece.kind === "text" || piece.kind === "delText") acc += piece.value;
        }
      }
      return acc;
    })
    .join("\n");
}

/**
 * Append a new row at the end of a table. The new row's cell count
 * matches the existing rows. If `texts` is shorter, the remaining cells
 * are empty.
 */
export function appendTableRow(table: WmlTable, texts: readonly string[]): WmlTableRow {
  const expectedCols = table.rows[0]?.cells.length ?? texts.length;
  const cells: WmlTableCell[] = [];
  for (let c = 0; c < expectedCols; c++) {
    const text = texts[c] ?? "";
    cells.push({
      paragraphs: [buildTextParagraph(text)],
      extras: [],
    });
  }
  const row: WmlTableRow = { cells, extras: [] };
  table.rows.push(row);
  return row;
}

/**
 * Remove the row at `index` from a table (0-based). Returns true if
 * removed.
 */
export function removeTableRow(table: WmlTable, index: number): boolean {
  if (index < 0 || index >= table.rows.length) return false;
  table.rows.splice(index, 1);
  return true;
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
