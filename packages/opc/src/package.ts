import { CONTENT_TYPES_PART_NAME, ContentTypesIndex } from "./content-types.js";
import { RelationshipSet } from "./relationships.js";
import type { ContentType, Part, PartCompression, PartName } from "./types.js";
import {
  lowerPartName,
  normalizePartName,
  partFolder,
  partNameToZipEntry,
  relsPartNameFor,
} from "./uri.js";
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
   * When true (default), parts known to be relationship containers
   * (`*.rels`) and the content types part are serialized only if their
   * in-memory state has been mutated. Pristine parts emit their original
   * bytes so re-saving an untouched package is bit-stable.
   */
  reuseUntouchedRelsBytes?: boolean;
}

/**
 * In-memory model of an OPC package (ECMA-376 Part 2).
 *
 * The package owns:
 *
 *   - An ordered list of binary parts (including `[Content_Types].xml`,
 *     which is stored as a normal part).
 *   - A {@link ContentTypesIndex} mirror of `[Content_Types].xml`.
 *   - A map from source part name (or "package") to {@link RelationshipSet}.
 *
 * On read, all three are populated from the ZIP archive. On write, they are
 * serialized back into the archive, replacing the corresponding parts.
 */
export class OpcPackage {
  private readonly parts: Map<PartName, PartEntry>;
  /** Original key order; used to keep ZIP central-directory order stable. */
  private readonly order: PartName[];
  private readonly contentTypesIndex: ContentTypesIndex;
  private readonly relsBySource: Map<PartName | "package", RelationshipSet>;
  /** ZIP entry name we observed for the relationships part of each source. */
  private readonly relsZipEntry: Map<PartName | "package", string>;

  private constructor(init: {
    parts: PartEntry[];
    contentTypesIndex: ContentTypesIndex;
    relsBySource: Map<PartName | "package", RelationshipSet>;
    relsZipEntry: Map<PartName | "package", string>;
  }) {
    this.parts = new Map();
    this.order = [];
    for (const part of init.parts) {
      this.parts.set(part.name, part);
      this.order.push(part.name);
    }
    this.contentTypesIndex = init.contentTypesIndex;
    this.relsBySource = init.relsBySource;
    this.relsZipEntry = init.relsZipEntry;
  }

  /** Construct an empty package with a fresh `[Content_Types].xml`. */
  static empty(): OpcPackage {
    return new OpcPackage({
      parts: [],
      contentTypesIndex: ContentTypesIndex.empty(),
      relsBySource: new Map(),
      relsZipEntry: new Map(),
    });
  }

  /** Parse an OPC package from a ZIP byte buffer. */
  static read(bytes: Uint8Array): OpcPackage {
    const entries = readZip(bytes);

    // First pass: locate [Content_Types].xml.
    let contentTypesIndex: ContentTypesIndex | undefined;
    for (const entry of entries) {
      if (entry.name === "[Content_Types].xml") {
        const xml = textDecode(entry.data);
        contentTypesIndex = ContentTypesIndex.fromXml(xml);
        break;
      }
    }
    if (!contentTypesIndex) {
      throw new Error("OPC package is missing [Content_Types].xml");
    }

    const parts: PartEntry[] = [];
    const relsBySource = new Map<PartName | "package", RelationshipSet>();
    const relsZipEntry = new Map<PartName | "package", string>();

    for (const entry of entries) {
      const partName = `/${entry.name}`;
      const normalized = normalizePartName(partName);
      if (entry.name === "[Content_Types].xml") {
        parts.push({
          name: CONTENT_TYPES_PART_NAME,
          contentType: "application/vnd.openxmlformats-package.content-types+xml",
          data: entry.data,
          compression: "deflate",
          zipEntry: entry.name,
        });
        continue;
      }

      if (isRelsEntry(entry.name)) {
        const source = relsSourceOf(normalized);
        const xml = textDecode(entry.data);
        relsBySource.set(source, RelationshipSet.fromXml(xml));
        relsZipEntry.set(source, entry.name);
        parts.push({
          name: normalized,
          contentType: RELS_CONTENT_TYPE,
          data: entry.data,
          compression: "deflate",
          zipEntry: entry.name,
        });
        continue;
      }

      const ct = contentTypesIndex.resolve(normalized) ?? "application/octet-stream";
      parts.push({
        name: normalized,
        contentType: ct,
        data: entry.data,
        compression: "deflate",
        zipEntry: entry.name,
      });
    }

    return new OpcPackage({ parts, contentTypesIndex, relsBySource, relsZipEntry });
  }

  /** Number of parts (including `[Content_Types].xml` and `.rels`). */
  get size(): number {
    return this.parts.size;
  }

  /** Iterate parts in their original (or insertion) order. */
  listParts(): readonly Part[] {
    return this.order.map((n) => {
      const entry = this.parts.get(n);
      if (!entry) throw new Error(`Inconsistent part order: ${n}`);
      return entry;
    });
  }

  hasPart(name: string): boolean {
    return this.parts.has(normalizePartName(name));
  }

  getPart(name: string): Part | undefined {
    return this.parts.get(normalizePartName(name));
  }

  /**
   * Add a new part to the package. The `[Content_Types].xml` index is
   * updated automatically with an `<Override>` entry unless an existing
   * `<Default>` already resolves to the requested content type.
   */
  addPart(input: {
    name: string;
    contentType: ContentType;
    data: Uint8Array;
    compression?: PartCompression;
  }): Part {
    const canonical = normalizePartName(input.name);
    if (this.parts.has(canonical)) {
      throw new Error(`Part already exists: ${canonical}`);
    }
    if (canonical === CONTENT_TYPES_PART_NAME) {
      throw new Error("Use ContentTypesIndex APIs to manipulate the content types part directly");
    }
    const compression = input.compression ?? "deflate";
    const entry: PartEntry = {
      name: canonical,
      contentType: input.contentType,
      data: input.data,
      compression,
      zipEntry: partNameToZipEntry(canonical),
    };
    this.parts.set(canonical, entry);
    this.order.push(canonical);

    const resolved = this.contentTypesIndex.resolve(canonical);
    if (resolved !== input.contentType) {
      this.contentTypesIndex.setOverride(canonical, input.contentType);
    }
    return entry;
  }

  removePart(name: string): boolean {
    const canonical = normalizePartName(name);
    if (canonical === CONTENT_TYPES_PART_NAME) {
      throw new Error("Cannot remove the content types part");
    }
    if (!this.parts.has(canonical)) return false;
    this.parts.delete(canonical);
    const idx = this.order.indexOf(canonical);
    if (idx >= 0) this.order.splice(idx, 1);
    this.contentTypesIndex.removeOverride(canonical);
    // Also drop any rels owned by this source.
    this.relsBySource.delete(canonical);
    this.relsZipEntry.delete(canonical);
    return true;
  }

  /** The mutable content-types index. */
  get contentTypes(): ContentTypesIndex {
    return this.contentTypesIndex;
  }

  /** Package-level (`/_rels/.rels`) relationships. Created on demand. */
  get packageRelationships(): RelationshipSet {
    let set = this.relsBySource.get("package");
    if (!set) {
      set = RelationshipSet.empty();
      this.relsBySource.set("package", set);
    }
    return set;
  }

  /** Relationships originating from a specific part. Created on demand. */
  partRelationships(source: string): RelationshipSet {
    const canonical = normalizePartName(source);
    let set = this.relsBySource.get(canonical);
    if (!set) {
      set = RelationshipSet.empty();
      this.relsBySource.set(canonical, set);
    }
    return set;
  }

  /** All relationship sets that currently hold at least one relationship. */
  allRelationships(): ReadonlyMap<PartName | "package", RelationshipSet> {
    return this.relsBySource;
  }

  /**
   * Serialize the package back to a ZIP byte buffer. Updates to content
   * types and relationships are re-emitted from their in-memory models;
   * payload parts pass through verbatim.
   */
  write(_opts: OpcWriteOptions = {}): Uint8Array {
    const seen = new Set<string>();
    const out: WriteZipEntry[] = [];

    const ctXml = this.contentTypesIndex.toXml();
    const ctBytes = textEncode(ctXml);

    // Emit parts in the original order, but substitute updated bytes for
    // [Content_Types].xml and `.rels` parts whose source is known.
    for (const name of this.order) {
      const part = this.parts.get(name);
      if (!part) continue;
      if (part.name === CONTENT_TYPES_PART_NAME) {
        out.push({ name: "[Content_Types].xml", data: ctBytes, compression: part.compression });
        seen.add("[Content_Types].xml");
        continue;
      }
      if (isRelsEntry(part.zipEntry)) {
        const source = relsSourceOf(part.name);
        const set = this.relsBySource.get(source);
        if (set) {
          const xml = set.toXml();
          out.push({ name: part.zipEntry, data: textEncode(xml), compression: part.compression });
          seen.add(part.zipEntry);
          continue;
        }
      }
      out.push({ name: part.zipEntry, data: part.data, compression: part.compression });
      seen.add(part.zipEntry);
    }

    // [Content_Types].xml must exist as the first new part if none was loaded.
    if (!seen.has("[Content_Types].xml")) {
      out.unshift({ name: "[Content_Types].xml", data: ctBytes, compression: "deflate" });
    }

    // Emit any newly created relationship sets that did not exist in the
    // source ZIP. Position them right after the source part for readability.
    for (const [source, set] of this.relsBySource) {
      const zipEntry = partNameToZipEntry(relsPartNameFor(source));
      if (seen.has(zipEntry)) continue;
      if (set.all.length === 0) continue;
      out.push({ name: zipEntry, data: textEncode(set.toXml()), compression: "deflate" });
      seen.add(zipEntry);

      // Register the rels part name so the content types resolve properly.
      const relsPartName = relsPartNameFor(source);
      if (!this.parts.has(relsPartName)) {
        const entry: PartEntry = {
          name: relsPartName,
          contentType: RELS_CONTENT_TYPE,
          data: textEncode(set.toXml()),
          compression: "deflate",
          zipEntry,
        };
        this.parts.set(relsPartName, entry);
        this.order.push(relsPartName);
      }
    }

    // Ensure the `rels` extension default is registered.
    if (!this.contentTypesIndex.allDefaults.some((d) => d.extension === "rels")) {
      this.contentTypesIndex.setDefault("rels", RELS_CONTENT_TYPE);
    }

    return writeZip(out);
  }
}

function isRelsEntry(zipName: string): boolean {
  return zipName.endsWith(".rels") && (zipName === "_rels/.rels" || zipName.includes("/_rels/"));
}

function relsSourceOf(relsPartName: PartName): PartName | "package" {
  if (relsPartName === "/_rels/.rels") return "package";
  // /foo/bar/_rels/baz.xml.rels  →  /foo/bar/baz.xml
  const folder = partFolder(relsPartName);
  const base = relsPartName.slice(folder.length);
  if (!base.endsWith(".rels")) return relsPartName; // shouldn't happen
  const parentFolder = folder.endsWith("_rels/") ? folder.slice(0, -"_rels/".length) : folder;
  return normalizePartName(`${parentFolder}${base.slice(0, -".rels".length)}`);
}

function textEncode(value: string): Uint8Array {
  return new TextEncoder().encode(value);
}

function textDecode(bytes: Uint8Array): string {
  return new TextDecoder("utf-8").decode(bytes);
}

// `lowerPartName` is re-exported here only so consumers do not need to dig
// into ./uri for case-insensitive comparisons.
export { lowerPartName };
