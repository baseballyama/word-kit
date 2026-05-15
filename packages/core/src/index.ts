/**
 * `@word-kit/core` is the public entry point for word-kit.
 *
 * @packageDocumentation
 */

export * from "./docx.js";
export { type ValidationIssue, validatePackage } from "./validator.js";
export {
  appendTableRow,
  appendTextRun,
  clearRunFormat,
  getParagraphAlignment,
  getParagraphNumbering,
  getParagraphStyle,
  getRunFormat,
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
  setParagraphAlignment,
  setParagraphBorders,
  setParagraphIndent,
  setParagraphShading,
  setParagraphSpacing,
  setParagraphText,
  setRunFormat,
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
} from "@word-kit/wml";
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
} from "@word-kit/wml";
export { VERSION } from "./version.js";
