import { parseXml, type XmlElement } from "@word-kit/ooxml-xml";
import type { OpcPackage } from "@word-kit/opc";
import { WML_NS, WML_RELATIONSHIPS } from "@word-kit/wml";

/**
 * A single problem the validator detected. `level` is `"error"` for issues
 * Word will reject ("needs repair") and `"warning"` for issues Word can
 * recover from but that suggest a bug in the producing code.
 */
export interface ValidationIssue {
  readonly level: "error" | "warning";
  readonly code: string;
  readonly message: string;
  readonly partName?: string;
}

/**
 * Run structural checks against an OPC package. Catches the most common
 * "Word needs to repair the file" causes:
 *
 * - Internal rels whose target part is missing
 * - commentReference w:id values without a matching <w:comment> in
 *   comments.xml
 * - footnoteReference / endnoteReference ids without a matching entry
 * - bookmarkStart ids without a matching bookmarkEnd
 * - image relationships whose media part is missing
 */
export function validatePackage(pkg: OpcPackage): ValidationIssue[] {
  const issues: ValidationIssue[] = [];

  // 1) Every internal rel target must exist.
  const allRelSources: Array<{ source: "package" | string; setOwner: string }> = [
    { source: "package", setOwner: "/_rels/.rels" },
  ];
  for (const part of pkg.listParts()) {
    if (part.name === "/[Content_Types].xml") continue;
    allRelSources.push({ source: part.name, setOwner: relsPartFor(part.name) });
  }
  const seenSources = new Set<string>();
  for (const entry of allRelSources) {
    if (seenSources.has(entry.setOwner)) continue;
    seenSources.add(entry.setOwner);
    const set =
      entry.source === "package" ? pkg.packageRelationships : pkg.partRelationships(entry.source);
    for (const rel of set.all) {
      if (rel.targetMode === "External") continue;
      const absolute = resolveRel(entry.source, rel.target);
      if (!pkg.hasPart(absolute)) {
        issues.push({
          level: "error",
          code: "rel-target-missing",
          message: `Relationship ${rel.id} of type ${rel.type} points at missing part ${absolute}`,
          partName: entry.setOwner,
        });
      }
    }
  }

  // 2) Walk /word/document.xml and verify cross-part references.
  const docPart = pkg.getPart("/word/document.xml");
  if (!docPart) return issues; // Already flagged by rel check above
  const docXml = new TextDecoder("utf-8").decode(docPart.data);
  let docTree: ReturnType<typeof parseXml>;
  try {
    docTree = parseXml(docXml);
  } catch (e) {
    issues.push({
      level: "error",
      code: "document-xml-parse",
      message: `word/document.xml failed to parse: ${(e as Error).message}`,
      partName: "/word/document.xml",
    });
    return issues;
  }

  const commentRefIds = collectAttrIds(docTree.root, "commentReference");
  const footnoteRefIds = collectAttrIds(docTree.root, "footnoteReference");
  const endnoteRefIds = collectAttrIds(docTree.root, "endnoteReference");
  const bookmarkStartIds = collectAttrIds(docTree.root, "bookmarkStart");
  const bookmarkEndIds = collectAttrIds(docTree.root, "bookmarkEnd");

  if (commentRefIds.size > 0) {
    const commentsPart = pkg.getPart("/word/comments.xml");
    if (!commentsPart) {
      issues.push({
        level: "error",
        code: "comments-part-missing",
        message: `document.xml references comment ids but /word/comments.xml is missing`,
      });
    } else {
      const ids = collectCommentIds(commentsPart.data, "comment");
      for (const id of commentRefIds) {
        if (!ids.has(id)) {
          issues.push({
            level: "error",
            code: "comment-id-missing",
            message: `commentReference w:id="${id}" has no matching <w:comment>`,
            partName: "/word/comments.xml",
          });
        }
      }
    }
  }

  if (footnoteRefIds.size > 0) {
    const fnPart = pkg.getPart("/word/footnotes.xml");
    if (!fnPart) {
      issues.push({
        level: "error",
        code: "footnotes-part-missing",
        message: `document.xml references footnote ids but /word/footnotes.xml is missing`,
      });
    } else {
      const ids = collectCommentIds(fnPart.data, "footnote");
      for (const id of footnoteRefIds) {
        if (!ids.has(id)) {
          issues.push({
            level: "error",
            code: "footnote-id-missing",
            message: `footnoteReference w:id="${id}" has no matching <w:footnote>`,
            partName: "/word/footnotes.xml",
          });
        }
      }
    }
  }

  if (endnoteRefIds.size > 0) {
    const enPart = pkg.getPart("/word/endnotes.xml");
    if (!enPart) {
      issues.push({
        level: "error",
        code: "endnotes-part-missing",
        message: `document.xml references endnote ids but /word/endnotes.xml is missing`,
      });
    } else {
      const ids = collectCommentIds(enPart.data, "endnote");
      for (const id of endnoteRefIds) {
        if (!ids.has(id)) {
          issues.push({
            level: "error",
            code: "endnote-id-missing",
            message: `endnoteReference w:id="${id}" has no matching <w:endnote>`,
            partName: "/word/endnotes.xml",
          });
        }
      }
    }
  }

  // bookmarkStart without matching End is a warning (Word tolerates it).
  for (const id of bookmarkStartIds) {
    if (!bookmarkEndIds.has(id)) {
      issues.push({
        level: "warning",
        code: "bookmark-end-missing",
        message: `bookmarkStart w:id="${id}" has no matching bookmarkEnd`,
        partName: "/word/document.xml",
      });
    }
  }

  // 3) Image rels: target media part must exist (already covered by rel
  // check, but flag with a more specific message).
  const docRels = pkg.partRelationships("/word/document.xml");
  for (const rel of docRels.byType(WML_RELATIONSHIPS.image)) {
    if (rel.targetMode === "External") continue;
    const partName = resolveRel("/word/document.xml", rel.target);
    if (!pkg.hasPart(partName)) {
      issues.push({
        level: "error",
        code: "image-missing",
        message: `image relationship ${rel.id} points at missing part ${partName}`,
        partName: "/word/document.xml",
      });
    }
  }

  return issues;
}

function collectAttrIds(root: XmlElement, localName: string): Set<string> {
  const out = new Set<string>();
  const visit = (el: XmlElement): void => {
    if (el.name.uri === WML_NS && el.name.local === localName) {
      const id = el.attrs.find((a) => a.name.uri === WML_NS && a.name.local === "id")?.value;
      if (id !== undefined) out.add(id);
    }
    for (const c of el.children) if (c.kind === "element") visit(c);
  };
  visit(root);
  return out;
}

function collectCommentIds(bytes: Uint8Array, elementLocal: string): Set<string> {
  const xml = new TextDecoder("utf-8").decode(bytes);
  let tree: ReturnType<typeof parseXml>;
  try {
    tree = parseXml(xml);
  } catch {
    return new Set();
  }
  const out = new Set<string>();
  for (const child of tree.root.children) {
    if (child.kind !== "element") continue;
    if (child.name.uri === WML_NS && child.name.local === elementLocal) {
      const id = child.attrs.find((a) => a.name.uri === WML_NS && a.name.local === "id")?.value;
      if (id !== undefined) out.add(id);
    }
  }
  return out;
}

function resolveRel(source: "package" | string, target: string): string {
  if (target.startsWith("/")) return target;
  if (source === "package") {
    return `/${target}`;
  }
  const folder = source.slice(0, source.lastIndexOf("/") + 1);
  return normalize(`${folder}${target}`);
}

function normalize(path: string): string {
  const parts = path.split("/").filter((p) => p !== "");
  const out: string[] = [];
  for (const p of parts) {
    if (p === ".") continue;
    if (p === "..") {
      out.pop();
      continue;
    }
    out.push(p);
  }
  return `/${out.join("/")}`;
}

function relsPartFor(partName: string): string {
  const folder = partName.slice(0, partName.lastIndexOf("/") + 1);
  const file = partName.slice(folder.length);
  return `${folder}_rels/${file}.rels`;
}
