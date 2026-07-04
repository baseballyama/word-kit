import {
  CONTENT_TYPES_PART_NAME,
  type ContentTypesIndex,
  emptyContentTypes,
  parseContentTypesXml,
  removeContentTypeOverride,
  resolveContentType,
  serializeContentTypesXml,
  setContentTypeDefault,
  setContentTypeOverride,
} from "./content-types.js";
import {
  emptyRelationshipSet,
  parseRelationshipsXml,
  type RelationshipSet,
  serializeRelationshipsXml,
} from "./relationships.js";
import type { ContentType, Part, PartCompression, PartName } from "./types.js";
import { normalizePartName, partFolder, partNameToZipEntry, relsPartNameFor } from "./uri.js";
import { readZip, writeZip, type WriteZipEntry } from "./zip.js";

const RELS_CONTENT_TYPE = "application/vnd.openxmlformats-package.relationships+xml";

interface PartEntry {
  name: PartName;
  contentType: ContentType;
  data: Uint8Array;
  compression: PartCompression;
  /** Original ZIP entry name preserved so we round-trip exact ZIP layout. */
  zipEntry: string;
}

export interface OpcWriteOptions {
  /**
   * Reserved for future use (eg. byte-stable round-trip of untouched
   * `.rels` and content types parts). Currently a no-op.
   */
  reuseUntouchedRelsBytes?: boolean;
}

/**
 * In-memory model of an OPC package (ECMA-376 Part 2).
 *
 * The package owns:
 *   - An ordered list of binary parts (including `[Content_Types].xml`,
 *     which is stored as a normal part).
 *   - A {@link ContentTypesIndex} mirror of `[Content_Types].xml`.
 *   - A map from source part name (or "package") to {@link RelationshipSet}.
 *
 * @remarks
 * Plain data type — no class. All operations are exposed as standalone
 * functions below (`readOpcPackage`, `addPart`, `getPart`, etc.) so
 * bundlers can tree-shake any operations the caller doesn't import.
 */
export interface OpcPackage {
  /** Map of part name → entry. */
  readonly parts: Map<PartName, PartEntry>;
  /** Insertion / source-ZIP order of parts, used to keep central-directory order stable. */
  readonly order: PartName[];
  /** Mirror of `[Content_Types].xml`. */
  readonly contentTypes: ContentTypesIndex;
  /** Relationship sets keyed by source part name (or `"package"`). */
  readonly relsBySource: Map<PartName | "package", RelationshipSet>;
  /** Original ZIP entry name we observed for each rels source. */
  readonly relsZipEntry: Map<PartName | "package", string>;
}

/** Construct an empty package with a fresh `[Content_Types].xml`. */
export function emptyOpcPackage(): OpcPackage {
  return {
    parts: new Map(),
    order: [],
    contentTypes: emptyContentTypes(),
    relsBySource: new Map(),
    relsZipEntry: new Map(),
  };
}

/** Parse an OPC package from a ZIP byte buffer. */
export function readOpcPackage(bytes: Uint8Array): OpcPackage {
  const entries = readZip(bytes);

  let contentTypes: ContentTypesIndex | undefined;
  for (const entry of entries) {
    if (entry.name === "[Content_Types].xml") {
      contentTypes = parseContentTypesXml(textDecode(entry.data));
      break;
    }
  }
  if (!contentTypes) {
    throw new Error("OPC package is missing [Content_Types].xml");
  }

  const pkg: OpcPackage = {
    parts: new Map(),
    order: [],
    contentTypes,
    relsBySource: new Map(),
    relsZipEntry: new Map(),
  };

  for (const entry of entries) {
    if (entry.name === "[Content_Types].xml") {
      pushPart(pkg, {
        name: CONTENT_TYPES_PART_NAME,
        contentType: "application/vnd.openxmlformats-package.content-types+xml",
        data: entry.data,
        compression: "deflate",
        zipEntry: entry.name,
      });
      continue;
    }

    const partName = `/${entry.name}`;
    const normalized = normalizePartName(partName);

    if (isRelsEntry(entry.name)) {
      const source = relsSourceOf(normalized);
      pkg.relsBySource.set(source, parseRelationshipsXml(textDecode(entry.data)));
      pkg.relsZipEntry.set(source, entry.name);
      pushPart(pkg, {
        name: normalized,
        contentType: RELS_CONTENT_TYPE,
        data: entry.data,
        compression: "deflate",
        zipEntry: entry.name,
      });
      continue;
    }

    const ct = resolveContentType(contentTypes, normalized) ?? "application/octet-stream";
    pushPart(pkg, {
      name: normalized,
      contentType: ct,
      data: entry.data,
      compression: "deflate",
      zipEntry: entry.name,
    });
  }

  return pkg;
}

/** Number of parts (including `[Content_Types].xml` and `.rels`). */
export function partsSize(pkg: OpcPackage): number {
  return pkg.parts.size;
}

/** Iterate parts in their original (or insertion) order. */
export function listParts(pkg: OpcPackage): readonly Part[] {
  return pkg.order.map((n) => {
    const entry = pkg.parts.get(n);
    if (!entry) throw new Error(`Inconsistent part order: ${n}`);
    return entry;
  });
}

/** Whether the package contains a part with the given name. */
export function hasPart(pkg: OpcPackage, name: string): boolean {
  return pkg.parts.has(normalizePartName(name));
}

/** Get a part by name, or `undefined`. */
export function getPart(pkg: OpcPackage, name: string): Part | undefined {
  return pkg.parts.get(normalizePartName(name));
}

/**
 * Add a new part to the package. The `[Content_Types].xml` index is
 * updated with an `<Override>` entry unless an existing `<Default>`
 * already resolves to the requested content type.
 */
export function addPart(
  pkg: OpcPackage,
  input: {
    name: string;
    contentType: ContentType;
    data: Uint8Array;
    compression?: PartCompression;
  },
): Part {
  const canonical = normalizePartName(input.name);
  if (pkg.parts.has(canonical)) {
    throw new Error(`Part already exists: ${canonical}`);
  }
  if (canonical === CONTENT_TYPES_PART_NAME) {
    throw new Error(
      "Use the content types functions to manipulate the content types part directly",
    );
  }
  const compression = input.compression ?? "deflate";
  const entry: PartEntry = {
    name: canonical,
    contentType: input.contentType,
    data: input.data,
    compression,
    zipEntry: partNameToZipEntry(canonical),
  };
  pushPart(pkg, entry);

  const resolved = resolveContentType(pkg.contentTypes, canonical);
  if (resolved !== input.contentType) {
    setContentTypeOverride(pkg.contentTypes, canonical, input.contentType);
  }
  return entry;
}

/** Remove a part by name. Returns true if removed. */
export function removePart(pkg: OpcPackage, name: string): boolean {
  const canonical = normalizePartName(name);
  if (canonical === CONTENT_TYPES_PART_NAME) {
    throw new Error("Cannot remove the content types part");
  }
  if (!pkg.parts.has(canonical)) return false;
  pkg.parts.delete(canonical);
  const idx = pkg.order.indexOf(canonical);
  if (idx >= 0) pkg.order.splice(idx, 1);
  removeContentTypeOverride(pkg.contentTypes, canonical);
  // Also drop any rels owned by this source.
  pkg.relsBySource.delete(canonical);
  pkg.relsZipEntry.delete(canonical);
  return true;
}

/** Package-level (`/_rels/.rels`) relationships, created on demand. */
export function packageRelationships(pkg: OpcPackage): RelationshipSet {
  let set = pkg.relsBySource.get("package");
  if (!set) {
    set = emptyRelationshipSet();
    pkg.relsBySource.set("package", set);
  }
  return set;
}

/** Relationships originating from a specific part, created on demand. */
export function partRelationships(pkg: OpcPackage, source: string): RelationshipSet {
  const canonical = normalizePartName(source);
  let set = pkg.relsBySource.get(canonical);
  if (!set) {
    set = emptyRelationshipSet();
    pkg.relsBySource.set(canonical, set);
  }
  return set;
}

/** All relationship sets that currently exist in the package. */
export function allRelationshipSources(
  pkg: OpcPackage,
): ReadonlyMap<PartName | "package", RelationshipSet> {
  return pkg.relsBySource;
}

/**
 * Serialize the package back to a ZIP byte buffer. Updates to content
 * types and relationships are re-emitted from their in-memory models;
 * payload parts pass through verbatim.
 */
export function writeOpcPackage(pkg: OpcPackage, _opts: OpcWriteOptions = {}): Uint8Array {
  const seen = new Set<string>();
  const out: WriteZipEntry[] = [];

  const ctBytes = textEncode(serializeContentTypesXml(pkg.contentTypes));

  for (const name of pkg.order) {
    const part = pkg.parts.get(name);
    if (!part) continue;
    if (part.name === CONTENT_TYPES_PART_NAME) {
      out.push({ name: "[Content_Types].xml", data: ctBytes, compression: part.compression });
      seen.add("[Content_Types].xml");
      continue;
    }
    if (isRelsEntry(part.zipEntry)) {
      const source = relsSourceOf(part.name);
      const set = pkg.relsBySource.get(source);
      if (set) {
        const xml = serializeRelationshipsXml(set);
        out.push({ name: part.zipEntry, data: textEncode(xml), compression: part.compression });
        seen.add(part.zipEntry);
        continue;
      }
    }
    out.push({ name: part.zipEntry, data: part.data, compression: part.compression });
    seen.add(part.zipEntry);
  }

  if (!seen.has("[Content_Types].xml")) {
    out.unshift({ name: "[Content_Types].xml", data: ctBytes, compression: "deflate" });
  }

  for (const [source, set] of pkg.relsBySource) {
    const zipEntry = partNameToZipEntry(relsPartNameFor(source));
    if (seen.has(zipEntry)) continue;
    if (set.relationships.length === 0) continue;
    const xml = serializeRelationshipsXml(set);
    out.push({ name: zipEntry, data: textEncode(xml), compression: "deflate" });
    seen.add(zipEntry);

    const relsPartName = relsPartNameFor(source);
    if (!pkg.parts.has(relsPartName)) {
      pushPart(pkg, {
        name: relsPartName,
        contentType: RELS_CONTENT_TYPE,
        data: textEncode(xml),
        compression: "deflate",
        zipEntry,
      });
    }
  }

  if (!pkg.contentTypes.defaults.some((d) => d.extension === "rels")) {
    setContentTypeDefault(pkg.contentTypes, "rels", RELS_CONTENT_TYPE);
  }

  return writeZip(out);
}

function pushPart(pkg: OpcPackage, part: PartEntry): void {
  pkg.parts.set(part.name, part);
  pkg.order.push(part.name);
}

function isRelsEntry(zipName: string): boolean {
  return zipName.endsWith(".rels") && (zipName === "_rels/.rels" || zipName.includes("/_rels/"));
}

function relsSourceOf(relsPartName: PartName): PartName | "package" {
  if (relsPartName === "/_rels/.rels") return "package";
  const folder = partFolder(relsPartName);
  const base = relsPartName.slice(folder.length);
  if (!base.endsWith(".rels")) return relsPartName;
  const parentFolder = folder.endsWith("_rels/") ? folder.slice(0, -"_rels/".length) : folder;
  return normalizePartName(`${parentFolder}${base.slice(0, -".rels".length)}`);
}

function textEncode(value: string): Uint8Array {
  return new TextEncoder().encode(value);
}

function textDecode(bytes: Uint8Array): string {
  return new TextDecoder("utf-8").decode(bytes);
}
