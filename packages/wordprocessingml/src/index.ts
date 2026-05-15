/**
 * WordprocessingML semantic AST and parser/writer for word-kit.
 *
 * @packageDocumentation
 */

export {
  DML_NS,
  DML_PIC_NS,
  DML_WP_NS,
  MC_NS,
  REL_NS,
  WML_CONTENT_TYPES,
  WML_NS,
  WML_RELATIONSHIPS,
} from "./namespaces.js";
export {
  appendTableRow,
  appendTextRun,
  type BuildTableOptions,
  buildTextParagraph,
  buildTextRun,
  buildTextTable,
  clearRunFormat,
  getParagraphAlignment,
  getParagraphNumbering,
  getParagraphStyle,
  getRunFormat,
  getTableCellText,
  mergeAdjacentRuns,
  type ParagraphAlignment,
  type ParagraphBordersOptions,
  type ParagraphIndent,
  type ParagraphShadingOptions,
  type ParagraphSpacing,
  removeTableRow,
  type RunFormatting,
  setParagraphAlignment,
  setParagraphBorders,
  setParagraphIndent,
  setParagraphShading,
  setParagraphSpacing,
  setParagraphStyle,
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
} from "./builders.js";
export {
  buildInlineDrawing,
  extensionForImageContentType,
  type InlineDrawingOptions,
  sniffImageContentType,
} from "./drawing.js";
export {
  buildFieldRuns,
  buildPageNumberFooterXml,
  buildSimpleField,
  WORD_FIELDS,
} from "./fields.js";
export { type BuildHyperlinkOptions, buildHyperlink, buildHyperlinkRun } from "./hyperlink.js";
export {
  type BuildCommentOptions,
  buildComment,
  buildCommentRangeEnd,
  buildCommentRangeStart,
  buildCommentReferenceRun,
  EMPTY_COMMENTS_XML,
  parseCommentsPart,
  type WmlCommentsPart,
  writeCommentsPart,
} from "./comments.js";
export {
  type BuildFootnoteOptions,
  buildFootnote,
  buildFootnoteReferenceRun,
  parseFootnotesPart,
  SEED_ENDNOTES_XML,
  SEED_FOOTNOTES_XML,
  type WmlFootnotesPart,
  writeFootnotesPart,
} from "./footnotes.js";
export {
  abstractNumId,
  numAbstractRef,
  numId,
  parseNumberingPart,
  type WmlNumberingPart,
  writeNumberingPart,
} from "./numbering.js";
export {
  type BuildAbstractNumLevelOptions,
  type BuildAbstractNumOptions,
  buildAbstractNum,
  buildNum,
  buildNumPr,
  buildPPrWithNumPr,
  bulletAbstractNumLevels,
  decimalAbstractNumLevels,
  EMPTY_NUMBERING_XML,
  type NumberFormat,
} from "./numbering-builders.js";
export {
  APP_PROPERTIES_CONTENT_TYPE,
  APP_PROPERTIES_REL_TYPE,
  type DocumentAppProperties,
  EMPTY_APP_PROPERTIES_XML,
  parseAppProperties,
  writeAppProperties,
} from "./app-properties.js";
export {
  CORE_PROPERTIES_CONTENT_TYPE,
  CORE_PROPERTIES_REL_TYPE,
  type DocumentCoreProperties,
  EMPTY_CORE_PROPERTIES_XML,
  parseCoreProperties,
  writeCoreProperties,
} from "./core-properties.js";
export { parseParagraph, parseWmlDocument } from "./parser.js";
export {
  addSectPrFooterRef,
  addSectPrHeaderRef,
  buildFooterReference,
  buildFooterXml,
  buildHeaderReference,
  buildHeaderXml,
  buildPgMar,
  buildPgSz,
  buildSectPr,
  type HeaderFooterType,
  MARGINS_NORMAL,
  PAGE_SIZE_A4,
  PAGE_SIZE_LETTER,
  type PageMargins,
  type PageSize,
  setSectPrPageMargins,
  setSectPrPageSize,
} from "./section-builders.js";
export { type BuildStyleOptions, buildStyle, MINIMAL_STYLES_XML } from "./style-builders.js";
export {
  findStyle,
  isDefaultStyle,
  isQFormat,
  parseStylesPart,
  styleBasedOn,
  styleId,
  styleLink,
  styleName,
  styleNext,
  styleType,
  type WmlStylesPart,
  type WmlStyleType,
  writeStylesPart,
} from "./styles.js";
export { acceptAllRevisions, rejectAllRevisions } from "./revisions.js";
export {
  documentText,
  findInParagraph,
  findText,
  paragraphText,
  replaceInParagraph,
  replaceText,
  type TextMatch,
} from "./text-search.js";
export type {
  PassThrough,
  WmlBlock,
  WmlBody,
  WmlDocument,
  WmlInline,
  WmlParagraph,
  WmlRawBlock,
  WmlRawInline,
  WmlRun,
  WmlRunPiece,
  WmlTable,
  WmlTableCell,
  WmlTableRow,
} from "./types.js";
export { paragraphToElement, writeWmlDocument } from "./writer.js";
