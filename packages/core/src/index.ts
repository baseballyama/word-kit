/**
 * `@word-kit/core` is the public entry point for word-kit.
 *
 * @packageDocumentation
 */

export {
  type AddImageOptions,
  type AppendParagraphOptions,
  Docx,
  type DocxCreateOptions,
} from "./docx.js";
export {
  appendTextRun,
  type HeaderFooterType,
  MARGINS_NORMAL,
  PAGE_SIZE_A4,
  PAGE_SIZE_LETTER,
  type PageMargins,
  type PageSize,
  type ParagraphAlignment,
  type ParagraphIndent,
  type ParagraphSpacing,
  type RunFormatting,
  setParagraphAlignment,
  setParagraphIndent,
  setParagraphSpacing,
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
