import { type MiniDocument, parseMiniXml, serializeMiniXml } from "./mini-xml.js";
import type { ContentType, PartName } from "./types.js";
import { lowerPartName, normalizePartName } from "./uri.js";

const CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types";

export interface DefaultEntry {
  /** Lowercase extension without leading dot (e.g. `"xml"`, `"rels"`). */
  readonly extension: string;
  readonly contentType: ContentType;
}

export interface OverrideEntry {
  readonly partName: PartName;
  readonly contentType: ContentType;
}

/**
 * In-memory model of `[Content_Types].xml`. Preserves the original entry
 * order so an unmodified package round-trips cleanly.
 *
 * @remarks
 * Plain data type — no class. Mutate via the standalone functions below
 * (`setDefault`, `setOverride`, etc.) so bundlers can tree-shake unused
 * operations.
 */
export interface ContentTypesIndex {
  readonly defaults: DefaultEntry[];
  readonly overrides: OverrideEntry[];
}

/** Construct an empty content types index. */
export function emptyContentTypes(): ContentTypesIndex {
  return { defaults: [], overrides: [] };
}

/**
 * Construct a content types index from explicit default + override lists.
 * Extensions are lower-cased; part names are normalized.
 */
export function createContentTypes(
  defaults: readonly DefaultEntry[] = [],
  overrides: readonly OverrideEntry[] = [],
): ContentTypesIndex {
  return {
    defaults: defaults.map((d) => ({ ...d, extension: d.extension.toLowerCase() })),
    overrides: overrides.map((o) => ({ ...o, partName: normalizePartName(o.partName) })),
  };
}

/** Parse a `[Content_Types].xml` payload into the index. */
export function parseContentTypesXml(xml: string): ContentTypesIndex {
  const doc = parseMiniXml(xml);
  if (doc.root !== "Types") {
    throw new Error(`[Content_Types].xml: expected <Types> root, got <${doc.root}>`);
  }
  const defaults: DefaultEntry[] = [];
  const overrides: OverrideEntry[] = [];
  for (const child of doc.children) {
    const attrs = Object.fromEntries(child.attrs);
    if (child.name === "Default") {
      const ext = attrs["Extension"];
      const ct = attrs["ContentType"];
      if (ext && ct) defaults.push({ extension: ext.toLowerCase(), contentType: ct });
    } else if (child.name === "Override") {
      const pn = attrs["PartName"];
      const ct = attrs["ContentType"];
      if (pn && ct) overrides.push({ partName: normalizePartName(pn), contentType: ct });
    }
  }
  return createContentTypes(defaults, overrides);
}

/** Serialize the content types index back to `[Content_Types].xml` payload. */
export function serializeContentTypesXml(index: ContentTypesIndex): string {
  const doc: MiniDocument = {
    root: "Types",
    rootAttrs: [["xmlns", CONTENT_TYPES_NS]],
    children: [
      ...index.defaults.map((d) => ({
        name: "Default",
        attrs: [
          ["Extension", d.extension],
          ["ContentType", d.contentType],
        ] as Array<[string, string]>,
      })),
      ...index.overrides.map((o) => ({
        name: "Override",
        attrs: [
          ["PartName", o.partName],
          ["ContentType", o.contentType],
        ] as Array<[string, string]>,
      })),
    ],
    standalone: "yes",
  };
  return serializeMiniXml(doc);
}

/**
 * Resolve the content type for a given part name. Override entries take
 * precedence over default entries (which match by file extension).
 */
export function resolveContentType(
  index: ContentTypesIndex,
  partName: PartName,
): ContentType | undefined {
  const canonical = normalizePartName(partName);
  const canonicalLower = lowerPartName(canonical);
  for (const o of index.overrides) {
    if (lowerPartName(o.partName) === canonicalLower) return o.contentType;
  }
  const dot = canonical.lastIndexOf(".");
  if (dot >= 0) {
    const ext = canonical.slice(dot + 1).toLowerCase();
    for (const d of index.defaults) {
      if (d.extension === ext) return d.contentType;
    }
  }
  return undefined;
}

/** Set or replace a `<Default>` entry for `extension`. */
export function setContentTypeDefault(
  index: ContentTypesIndex,
  extension: string,
  contentType: ContentType,
): void {
  const ext = extension.toLowerCase();
  const existing = index.defaults.findIndex((d) => d.extension === ext);
  if (existing >= 0) {
    index.defaults[existing] = { extension: ext, contentType };
  } else {
    index.defaults.push({ extension: ext, contentType });
  }
}

/** Remove the `<Default>` entry for `extension`. Returns true if removed. */
export function removeContentTypeDefault(index: ContentTypesIndex, extension: string): boolean {
  const ext = extension.toLowerCase();
  const idx = index.defaults.findIndex((d) => d.extension === ext);
  if (idx < 0) return false;
  index.defaults.splice(idx, 1);
  return true;
}

/** Set or replace an `<Override>` entry for `partName`. */
export function setContentTypeOverride(
  index: ContentTypesIndex,
  partName: PartName,
  contentType: ContentType,
): void {
  const canonical = normalizePartName(partName);
  const canonicalLower = lowerPartName(canonical);
  const existing = index.overrides.findIndex((o) => lowerPartName(o.partName) === canonicalLower);
  if (existing >= 0) {
    index.overrides[existing] = { partName: canonical, contentType };
  } else {
    index.overrides.push({ partName: canonical, contentType });
  }
}

/** Remove the `<Override>` entry for `partName`. Returns true if removed. */
export function removeContentTypeOverride(index: ContentTypesIndex, partName: PartName): boolean {
  const canonical = lowerPartName(normalizePartName(partName));
  const idx = index.overrides.findIndex((o) => lowerPartName(o.partName) === canonical);
  if (idx < 0) return false;
  index.overrides.splice(idx, 1);
  return true;
}

export const CONTENT_TYPES_PART_NAME: PartName = "/[Content_Types].xml";
