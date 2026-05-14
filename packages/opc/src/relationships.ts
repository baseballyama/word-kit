import { type MiniDocument, parseMiniXml, serializeMiniXml } from "./mini-xml.js";
import type { Relationship, RelationshipId, RelationshipTargetMode } from "./types.js";

const RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

/**
 * Ordered, mutable collection of relationships scoped to a single source.
 *
 * The source may be the package itself (for `/_rels/.rels`) or an individual
 * part (for `<folder>/_rels/<name>.rels`). The set does not know which it is;
 * the package layer owns that mapping.
 *
 * Relationship target values are stored exactly as written in the `.rels`
 * file, including relative paths like `media/image1.png`. Resolution to an
 * absolute part name is the caller's responsibility (see `uri.ts`).
 */
export class RelationshipSet {
  private readonly relationships: Relationship[];
  private nextAutoId = 1;

  constructor(initial: readonly Relationship[] = []) {
    this.relationships = initial.map((r) => ({ ...r }));
    this.refreshAutoId();
  }

  static empty(): RelationshipSet {
    return new RelationshipSet();
  }

  static fromXml(xml: string): RelationshipSet {
    const doc = parseMiniXml(xml);
    if (doc.root !== "Relationships") {
      throw new Error(`.rels: expected <Relationships> root, got <${doc.root}>`);
    }
    const rels: Relationship[] = [];
    for (const child of doc.children) {
      if (child.name !== "Relationship") continue;
      const attrs = Object.fromEntries(child.attrs);
      const id = attrs["Id"];
      const type = attrs["Type"];
      const target = attrs["Target"];
      const targetMode: RelationshipTargetMode =
        attrs["TargetMode"] === "External" ? "External" : "Internal";
      if (id && type && target !== undefined) {
        rels.push({ id, type, target, targetMode });
      }
    }
    return new RelationshipSet(rels);
  }

  toXml(): string {
    const doc: MiniDocument = {
      root: "Relationships",
      rootAttrs: [["xmlns", RELATIONSHIPS_NS]],
      children: this.relationships.map((r) => ({
        name: "Relationship",
        attrs: [
          ["Id", r.id],
          ["Type", r.type],
          ["Target", r.target],
          ...(r.targetMode === "External"
            ? ([["TargetMode", "External"]] as Array<[string, string]>)
            : []),
        ] as Array<[string, string]>,
      })),
      standalone: "yes",
    };
    return serializeMiniXml(doc);
  }

  get all(): readonly Relationship[] {
    return this.relationships;
  }

  byId(id: RelationshipId): Relationship | undefined {
    return this.relationships.find((r) => r.id === id);
  }

  byType(type: string): readonly Relationship[] {
    return this.relationships.filter((r) => r.type === type);
  }

  /**
   * Add a relationship. If `id` is omitted, the lowest unused `rId<N>` is
   * assigned (per ECMA-376 Part 2 §9.3 — IDs are arbitrary; we keep them
   * stable and reuse gaps only after explicit removal).
   */
  add(input: {
    id?: RelationshipId;
    type: string;
    target: string;
    targetMode?: RelationshipTargetMode;
  }): Relationship {
    const id = input.id ?? this.allocateId();
    if (this.byId(id)) {
      throw new Error(`Relationship id already in use: ${id}`);
    }
    const rel: Relationship = {
      id,
      type: input.type,
      target: input.target,
      targetMode: input.targetMode ?? "Internal",
    };
    this.relationships.push(rel);
    return rel;
  }

  remove(id: RelationshipId): boolean {
    const idx = this.relationships.findIndex((r) => r.id === id);
    if (idx < 0) return false;
    this.relationships.splice(idx, 1);
    return true;
  }

  private allocateId(): RelationshipId {
    while (this.byId(`rId${this.nextAutoId}`)) {
      this.nextAutoId++;
    }
    return `rId${this.nextAutoId++}`;
  }

  private refreshAutoId(): void {
    let max = 0;
    for (const r of this.relationships) {
      const m = /^rId(\d+)$/.exec(r.id);
      if (m && m[1]) {
        const n = Number.parseInt(m[1], 10);
        if (Number.isFinite(n) && n > max) max = n;
      }
    }
    this.nextAutoId = max + 1;
  }
}
