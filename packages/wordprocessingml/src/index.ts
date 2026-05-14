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
  buildInlineDrawing,
  extensionForImageContentType,
  type InlineDrawingOptions,
  sniffImageContentType,
} from "./drawing.js";
export { parseWmlDocument } from "./parser.js";
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
} from "./types.js";
export { writeWmlDocument } from "./writer.js";
