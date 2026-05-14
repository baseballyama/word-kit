/**
 * Open Packaging Conventions (ECMA-376 Part 2) layer for word-kit.
 *
 * @packageDocumentation
 */

export {
  CONTENT_TYPES_PART_NAME,
  ContentTypesIndex,
  type DefaultEntry,
  type OverrideEntry,
} from "./content-types.js";
export { buildMinimalDocx } from "./minimal-docx.js";
export { OpcPackage, type OpcWriteOptions } from "./package.js";
export { RelationshipSet } from "./relationships.js";
export type {
  ContentType,
  Part,
  PartCompression,
  PartName,
  Relationship,
  RelationshipId,
  RelationshipTargetMode,
} from "./types.js";
export {
  lowerPartName,
  normalizePartName,
  partBaseName,
  partFolder,
  partNameToZipEntry,
  relativizeTarget,
  relsPartNameFor,
  resolveInternalTarget,
  zipEntryToPartName,
} from "./uri.js";
