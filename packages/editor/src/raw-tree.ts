/**
 * Raw XML node access for the universal editor.
 *
 * DrawingML, OMML (math), VML and any other markup the semantic AST doesn't
 * model is preserved as raw `XmlElement` subtrees inside `document.xml` (in run
 * pieces, raw inlines, and raw blocks). This module collects those subtrees and
 * all their descendants so the raw-XML inspector can present — and the raw
 * commands can edit — every element, then flush via the document AST on save.
 */

import {
  childElementsOf,
  type Docx,
  getRawPartRoot,
  type XmlElement,
  xmlPartNames,
} from "@office-kit/docx";

/** A node in the raw-XML inspector tree. */
export interface RawNode {
  readonly element: XmlElement;
  /** Local element name (e.g. `blip`, `oMath`, `shape`). */
  readonly local: string;
  /** Namespace prefix as authored (e.g. `a`, `m`, `v`, `w`). */
  readonly prefix: string;
  readonly children: RawNode[];
}

function toRawNode(element: XmlElement): RawNode {
  return {
    element,
    local: element.name.local,
    prefix: element.name.prefix ?? "",
    children: childElementsOf(element).map(toRawNode),
  };
}

/** Every XML part in the package the part-level raw editor can open. */
export function xmlParts(doc: Docx): string[] {
  return xmlPartNames(doc);
}

/**
 * The raw-XML tree for a single part (fontTable / settings / styles / … ), for
 * the part inspector. Edits go through the `rawpart.*` commands, which mark the
 * part dirty so it is re-serialized on save.
 */
export function partRawTree(doc: Docx, partName: string): RawNode | undefined {
  const root = getRawPartRoot(doc, partName);
  return root ? toRawNode(root) : undefined;
}

/** Every top-level raw XML subtree in the document body, as inspector trees. */
export function rawTrees(doc: Docx): RawNode[] {
  const roots: XmlElement[] = [];
  const visitParagraph = (children: { kind: string }[]): void => {
    for (const child of children as Array<
      | { kind: "run"; pieces: Array<{ kind: string; node?: XmlElement }> }
      | { kind: "raw"; node: XmlElement }
      | { kind: string }
    >) {
      if (child.kind === "run" && "pieces" in child) {
        for (const piece of child.pieces) {
          if (
            (piece.kind === "drawing" || piece.kind === "pict" || piece.kind === "raw") &&
            piece.node
          ) {
            roots.push(piece.node);
          }
        }
      } else if (child.kind === "raw" && "node" in child) {
        roots.push(child.node);
      }
    }
  };
  for (const block of doc.document.body.blocks) {
    if (block.kind === "paragraph") visitParagraph(block.children);
    else if (block.kind === "table") {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          for (const p of cell.paragraphs) visitParagraph(p.children);
        }
      }
    } else if (block.kind === "raw") {
      roots.push(block.node);
    }
  }
  return roots.map(toRawNode);
}

/** Flatten the raw trees to every element (for search / coverage reach). */
export function allRawElements(doc: Docx): XmlElement[] {
  const out: XmlElement[] = [];
  const walk = (n: RawNode): void => {
    out.push(n.element);
    for (const c of n.children) walk(c);
  };
  for (const t of rawTrees(doc)) walk(t);
  return out;
}
