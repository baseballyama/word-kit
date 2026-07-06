/**
 * `@office-kit/docx` is the public entry point for word-kit.
 *
 * @packageDocumentation
 */

export * from "./docx.js";
export { type ValidationIssue, validatePackage } from "./validator.js";
export {
  appendTableRow,
  appendTextRun,
  type BuildStyleOptions,
  type BuildTableOptions,
  appendChildElement,
  childElementsOf,
  clearRunFormat,
  type DocumentAppProperties,
  type DocumentCoreProperties,
  getElementAttr,
  getElementProp,
  getParagraphAlignment,
  getParagraphNumbering,
  getParagraphProp,
  getParagraphStyle,
  getRunFormat,
  getRunProp,
  makePropsElement,
  getTableCellText,
  mergeAdjacentRuns,
  type HeaderFooterType,
  MARGINS_NORMAL,
  PAGE_SIZE_A4,
  PAGE_SIZE_LETTER,
  type PageMargins,
  type PageSize,
  type ParagraphAlignment,
  type ParagraphBordersOptions,
  type ParagraphIndent,
  type ParagraphShadingOptions,
  type ParagraphSpacing,
  paragraphText,
  removeTableRow,
  type RunFormatting,
  setElementAttr,
  setElementOnOff,
  setElementValProp,
  setParagraphAlignment,
  setParagraphBorders,
  setParagraphIndent,
  setParagraphOnOff,
  setParagraphShading,
  setParagraphSpacing,
  setParagraphStyle,
  setParagraphText,
  setParagraphValProp,
  setRunFormat,
  setRunOnOff,
  setRunValProp,
  setTableBorders,
  setTableCellShading,
  setTableCellText,
  setTableCellVerticalAlign,
  setTableRowAsHeader,
  setTableRowHeight,
  type TableBorderStyle,
  type TableBordersOptions,
  type TableCellShadingOptions,
  type TableCellVerticalAlign,
  type TableRowHeightRule,
} from "../internal/wordprocessingml/index.js";
// Re-export types that are part of the public surface area so consumers can
// use them without importing from internal packages directly.
export type {
  TextMatch,
  WmlBlock,
  WmlBody,
  WmlDocument,
  WmlInline,
  WmlParagraph,
  WmlRun,
  WmlRunPiece,
  WmlTable,
  WmlTableCell,
  WmlTableRow,
} from "../internal/wordprocessingml/index.js";
// Raw XML AST types — the low-level escape hatch. `WmlRun.rPr`, `WmlTableCell.tcPr`,
// etc. are `XmlElement`, so consumers manipulating them (e.g. generic property
// setters) need these names.
export type { QName, XmlAttr, XmlElement, XmlNode } from "../internal/xml/index.js";
export { VERSION } from "./version.js";
