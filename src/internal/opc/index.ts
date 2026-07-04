/**
 * Open Packaging Conventions (ECMA-376 Part 2) layer for word-kit.
 *
 * @packageDocumentation
 */

export {
  CONTENT_TYPES_PART_NAME,
  type ContentTypesIndex,
  createContentTypes,
  type DefaultEntry,
  emptyContentTypes,
  type OverrideEntry,
  parseContentTypesXml,
  removeContentTypeDefault,
  removeContentTypeOverride,
  resolveContentType,
  serializeContentTypesXml,
  setContentTypeDefault,
  setContentTypeOverride,
} from "./content-types.js";
export { buildMinimalDocx } from "./minimal-docx.js";
export {
  addPart,
  allRelationshipSources,
  emptyOpcPackage,
  getPart,
  hasPart,
  listParts,
  type OpcPackage,
  type OpcWriteOptions,
  packageRelationships,
  partRelationships,
  partsSize,
  readOpcPackage,
  removePart,
  writeOpcPackage,
} from "./package.js";
export {
  addRelationship,
  allRelationships,
  createRelationshipSet,
  emptyRelationshipSet,
  parseRelationshipsXml,
  relationshipById,
  relationshipsByType,
  type RelationshipSet,
  removeRelationship,
  serializeRelationshipsXml,
} from "./relationships.js";
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
