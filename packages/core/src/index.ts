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
} from "@word-kit/wml";
export { VERSION } from "./version.js";
