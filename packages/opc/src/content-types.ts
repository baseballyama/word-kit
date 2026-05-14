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
 */
export class ContentTypesIndex {
  private readonly defaults: DefaultEntry[];
  private readonly overrides: OverrideEntry[];

  constructor(defaults: readonly DefaultEntry[] = [], overrides: readonly OverrideEntry[] = []) {
    this.defaults = defaults.map((d) => ({ ...d, extension: d.extension.toLowerCase() }));
    this.overrides = overrides.map((o) => ({ ...o, partName: normalizePartName(o.partName) }));
  }

  static empty(): ContentTypesIndex {
    return new ContentTypesIndex();
  }

  static fromXml(xml: string): ContentTypesIndex {
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
    return new ContentTypesIndex(defaults, overrides);
  }

  toXml(): string {
    const doc: MiniDocument = {
      root: "Types",
      rootAttrs: [["xmlns", CONTENT_TYPES_NS]],
      children: [
        ...this.defaults.map((d) => ({
          name: "Default",
          attrs: [
            ["Extension", d.extension],
            ["ContentType", d.contentType],
          ] as Array<[string, string]>,
        })),
        ...this.overrides.map((o) => ({
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

  /** Resolve the content type for a given part name. Overrides take precedence. */
  resolve(partName: PartName): ContentType | undefined {
    const canonical = normalizePartName(partName);
    const canonicalLower = lowerPartName(canonical);
    for (const o of this.overrides) {
      if (lowerPartName(o.partName) === canonicalLower) return o.contentType;
    }
    const dot = canonical.lastIndexOf(".");
    if (dot >= 0) {
      const ext = canonical.slice(dot + 1).toLowerCase();
      for (const d of this.defaults) {
        if (d.extension === ext) return d.contentType;
      }
    }
    return undefined;
  }

  get allDefaults(): readonly DefaultEntry[] {
    return this.defaults;
  }

  get allOverrides(): readonly OverrideEntry[] {
    return this.overrides;
  }

  /** Set or replace a `<Default>` entry for `extension`. Returns this. */
  setDefault(extension: string, contentType: ContentType): this {
    const ext = extension.toLowerCase();
    const existing = this.defaults.findIndex((d) => d.extension === ext);
    if (existing >= 0) {
      this.defaults[existing] = { extension: ext, contentType };
    } else {
      this.defaults.push({ extension: ext, contentType });
    }
    return this;
  }

  /** Remove the `<Default>` entry for `extension`. Returns true if removed. */
  removeDefault(extension: string): boolean {
    const ext = extension.toLowerCase();
    const idx = this.defaults.findIndex((d) => d.extension === ext);
    if (idx < 0) return false;
    this.defaults.splice(idx, 1);
    return true;
  }

  /** Set or replace an `<Override>` entry for `partName`. */
  setOverride(partName: PartName, contentType: ContentType): this {
    const canonical = normalizePartName(partName);
    const canonicalLower = lowerPartName(canonical);
    const existing = this.overrides.findIndex((o) => lowerPartName(o.partName) === canonicalLower);
    if (existing >= 0) {
      this.overrides[existing] = { partName: canonical, contentType };
    } else {
      this.overrides.push({ partName: canonical, contentType });
    }
    return this;
  }

  /** Remove the `<Override>` entry for `partName`. Returns true if removed. */
  removeOverride(partName: PartName): boolean {
    const canonical = lowerPartName(normalizePartName(partName));
    const idx = this.overrides.findIndex((o) => lowerPartName(o.partName) === canonical);
    if (idx < 0) return false;
    this.overrides.splice(idx, 1);
    return true;
  }
}

export const CONTENT_TYPES_PART_NAME: PartName = "/[Content_Types].xml";
