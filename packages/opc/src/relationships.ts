import { type MiniDocument, parseMiniXml, serializeMiniXml } from "./mini-xml.js";
import type { Relationship, RelationshipId, RelationshipTargetMode } from "./types.js";

const RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

/**
 * Ordered, mutable collection of relationships scoped to a single source.
 *
 * The source may be the package itself (for `/_rels/.rels`) or an individual
 * part (for `<folder>/_rels/<name>.rels`). The state does not know which it
 * is; the package layer owns that mapping.
 *
 * Relationship target values are stored exactly as written in the `.rels`
 * file, including relative paths like `media/image1.png`. Resolution to an
 * absolute part name is the caller's responsibility (see `uri.ts`).
 *
 * @remarks
 * Plain data type — no class. Mutate via the standalone functions below
 * (`addRelationship`, `removeRelationship`, etc.) so that bundlers can
 * tree-shake unused operations.
 */
export interface RelationshipSet {
  readonly relationships: Relationship[];
  /** Next `rId<N>` candidate to try when auto-allocating. Internal. */
  nextAutoId: number;
}

/** Construct an empty relationship set. */
export function emptyRelationshipSet(): RelationshipSet {
  return { relationships: [], nextAutoId: 1 };
}

/** Construct a relationship set seeded from an initial array. */
export function createRelationshipSet(initial: readonly Relationship[] = []): RelationshipSet {
  const set: RelationshipSet = {
    relationships: initial.map((r) => ({ ...r })),
    nextAutoId: 1,
  };
  refreshAutoId(set);
  return set;
}

/** Parse a `.rels` XML payload into a relationship set. */
export function parseRelationshipsXml(xml: string): RelationshipSet {
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
  return createRelationshipSet(rels);
}

/** Serialize a relationship set to `.rels` XML. */
export function serializeRelationshipsXml(set: RelationshipSet): string {
  const doc: MiniDocument = {
    root: "Relationships",
    rootAttrs: [["xmlns", RELATIONSHIPS_NS]],
    children: set.relationships.map((r) => ({
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

/** Return all relationships in a set in their original insertion order. */
export function allRelationships(set: RelationshipSet): readonly Relationship[] {
  return set.relationships;
}

/** Find a relationship by id. */
export function relationshipById(
  set: RelationshipSet,
  id: RelationshipId,
): Relationship | undefined {
  return set.relationships.find((r) => r.id === id);
}

/** Filter relationships by type URI. */
export function relationshipsByType(set: RelationshipSet, type: string): readonly Relationship[] {
  return set.relationships.filter((r) => r.type === type);
}

/**
 * Add a relationship to the set. If `id` is omitted, the lowest unused
 * `rId<N>` is assigned. Throws if the supplied `id` is already in use.
 */
export function addRelationship(
  set: RelationshipSet,
  input: {
    id?: RelationshipId;
    type: string;
    target: string;
    targetMode?: RelationshipTargetMode;
  },
): Relationship {
  const id = input.id ?? allocateRelationshipId(set);
  if (relationshipById(set, id)) {
    throw new Error(`Relationship id already in use: ${id}`);
  }
  const rel: Relationship = {
    id,
    type: input.type,
    target: input.target,
    targetMode: input.targetMode ?? "Internal",
  };
  set.relationships.push(rel);
  return rel;
}

/** Remove a relationship by id. Returns true if found and removed. */
export function removeRelationship(set: RelationshipSet, id: RelationshipId): boolean {
  const idx = set.relationships.findIndex((r) => r.id === id);
  if (idx < 0) return false;
  set.relationships.splice(idx, 1);
  return true;
}

function allocateRelationshipId(set: RelationshipSet): RelationshipId {
  while (relationshipById(set, `rId${set.nextAutoId}`)) {
    set.nextAutoId++;
  }
  return `rId${set.nextAutoId++}`;
}

function refreshAutoId(set: RelationshipSet): void {
  let max = 0;
  for (const r of set.relationships) {
    const m = /^rId(\d+)$/.exec(r.id);
    if (m?.[1]) {
      const n = Number.parseInt(m[1], 10);
      if (Number.isFinite(n) && n > max) max = n;
    }
  }
  set.nextAutoId = max + 1;
}
