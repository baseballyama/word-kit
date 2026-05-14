/**
 * Canonical OPC part name. Must start with `/`, must not end with `/`.
 *
 * Example: `/word/document.xml`, `/[Content_Types].xml`, `/_rels/.rels`.
 */
export type PartName = string;

/** OPC relationship identifier (e.g. `rId1`). Case-sensitive per spec. */
export type RelationshipId = string;

/** Content-Type as written in `[Content_Types].xml` (case-insensitive per spec). */
export type ContentType = string;

export type RelationshipTargetMode = "Internal" | "External";

export interface Relationship {
  readonly id: RelationshipId;
  readonly type: string;
  readonly target: string;
  readonly targetMode: RelationshipTargetMode;
}

/** Compression method used for the ZIP entry that carries this part. */
export type PartCompression = "store" | "deflate";

export interface Part {
  readonly name: PartName;
  contentType: ContentType;
  data: Uint8Array;
  compression: PartCompression;
}
