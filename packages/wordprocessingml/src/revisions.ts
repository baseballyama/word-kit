import type { XmlElement, XmlNode } from "@word-kit/ooxml-xml";
import { WML_NS } from "./namespaces.js";
import type {
  WmlBlock,
  WmlBody,
  WmlDocument,
  WmlInline,
  WmlParagraph,
  WmlRawInline,
  WmlRun,
  WmlRunPiece,
  WmlTable,
  WmlTableCell,
  WmlTableRow,
} from "./types.js";

/**
 * Accept all tracked-change insertions/deletions in the document.
 *
 * - `<w:ins>` content is kept (the wrapper is dropped).
 * - `<w:del>` content is removed.
 * - `<w:delText>` inside a remaining run (i.e. user manually unwrapped a
 *   del) is treated as deleted and dropped.
 *
 * Returns the total number of revisions resolved.
 */
export function acceptAllRevisions(doc: WmlDocument): number {
  const ctx = { count: 0 };
  walkBody(doc.body, (b) => processBlock(b, ctx, "accept"));
  return ctx.count;
}

/**
 * Reject all tracked-change insertions/deletions in the document.
 *
 * - `<w:ins>` content is removed.
 * - `<w:del>` content is kept (the wrapper is dropped, and any
 *   `<w:delText>` inside becomes a normal `<w:t>`).
 *
 * Returns the total number of revisions resolved.
 */
export function rejectAllRevisions(doc: WmlDocument): number {
  const ctx = { count: 0 };
  walkBody(doc.body, (b) => processBlock(b, ctx, "reject"));
  return ctx.count;
}

interface Ctx {
  count: number;
}
type Mode = "accept" | "reject";

function walkBody(body: WmlBody, visit: (b: WmlBlock) => void): void {
  for (const b of body.blocks) visit(b);
}

function processBlock(block: WmlBlock, ctx: Ctx, mode: Mode): void {
  if (block.kind === "paragraph") {
    processParagraph(block, ctx, mode);
    return;
  }
  if (block.kind === "table") {
    processTable(block, ctx, mode);
    return;
  }
}

function processTable(table: WmlTable, ctx: Ctx, mode: Mode): void {
  for (const row of table.rows) processRow(row, ctx, mode);
}

function processRow(row: WmlTableRow, ctx: Ctx, mode: Mode): void {
  for (const cell of row.cells) processCell(cell, ctx, mode);
}

function processCell(cell: WmlTableCell, ctx: Ctx, mode: Mode): void {
  for (const p of cell.paragraphs) processParagraph(p, ctx, mode);
}

function processParagraph(p: WmlParagraph, ctx: Ctx, mode: Mode): void {
  const next: WmlInline[] = [];
  for (const child of p.children) {
    if (child.kind === "raw" && isInsOrDel(child)) {
      const isIns = isElement(child.node, "ins");
      ctx.count++;
      if (mode === "accept" && isIns) {
        next.push(...extractInsContent(child));
      } else if (mode === "reject" && !isIns) {
        next.push(...extractDelContent(child));
      }
      // else: drop entirely
      continue;
    }
    if (child.kind === "run") {
      // Strip delText pieces if accepting; convert to text if rejecting.
      child.pieces = child.pieces
        .map((piece): WmlRunPiece | undefined => {
          if (piece.kind === "delText") {
            ctx.count++;
            if (mode === "accept") return undefined; // drop
            return {
              kind: "text",
              value: piece.value,
              preserveSpace: piece.preserveSpace,
            };
          }
          if (piece.kind === "delInstrText") {
            ctx.count++;
            if (mode === "accept") return undefined;
            return {
              kind: "instrText",
              value: piece.value,
              preserveSpace: piece.preserveSpace,
            };
          }
          return piece;
        })
        .filter((p2): p2 is WmlRunPiece => p2 !== undefined);
      next.push(child);
      continue;
    }
    next.push(child);
  }
  p.children = next;
}

function extractInsContent(raw: WmlRawInline): WmlInline[] {
  return promoteRunChildren(raw.node);
}

function extractDelContent(raw: WmlRawInline): WmlInline[] {
  // For <w:del>, runs inside use <w:delText>; convert them to text on
  // re-promotion since reject means "as if the deletion never happened".
  const inlines = promoteRunChildren(raw.node);
  for (const inline of inlines) {
    if (inline.kind !== "run") continue;
    inline.pieces = inline.pieces.map((piece): WmlRunPiece => {
      if (piece.kind === "delText") {
        return {
          kind: "text",
          value: piece.value,
          preserveSpace: piece.preserveSpace,
        };
      }
      if (piece.kind === "delInstrText") {
        return {
          kind: "instrText",
          value: piece.value,
          preserveSpace: piece.preserveSpace,
        };
      }
      return piece;
    });
  }
  return inlines;
}

function promoteRunChildren(insOrDel: XmlElement): WmlInline[] {
  const result: WmlInline[] = [];
  for (const child of insOrDel.children) {
    if (child.kind !== "element") continue;
    if (child.name.uri === WML_NS && child.name.local === "r") {
      result.push(parseRunFromXml(child));
    } else {
      // Unknown nested element (e.g. nested ins/del, sdt) — preserve as raw.
      result.push({ kind: "raw", node: child });
    }
  }
  return result;
}

function parseRunFromXml(r: XmlElement): WmlRun {
  let rPr: XmlElement | undefined;
  const pieces: WmlRunPiece[] = [];
  for (const c of r.children) {
    if (c.kind !== "element") continue;
    if (c.name.uri === WML_NS && c.name.local === "rPr") {
      rPr = c;
      continue;
    }
    // Promote known text-bearing children minimally.
    if (c.name.uri === WML_NS && (c.name.local === "t" || c.name.local === "delText")) {
      const value = readText(c);
      const preserveSpace = c.attrs.some(
        (a) =>
          a.name.uri === "http://www.w3.org/XML/1998/namespace" &&
          a.name.local === "space" &&
          a.value === "preserve",
      );
      pieces.push({
        kind: c.name.local === "delText" ? "delText" : "text",
        value,
        preserveSpace,
      });
      continue;
    }
    pieces.push({ kind: "raw", node: c });
  }
  return {
    kind: "run",
    ...(rPr ? { rPr } : {}),
    pieces,
    extras: [],
  };
}

function readText(el: XmlElement): string {
  let acc = "";
  for (const c of el.children) {
    if (c.kind === "text") acc += c.value;
    else if (c.kind === "cdata") acc += c.value;
  }
  return acc;
}

function isInsOrDel(raw: WmlRawInline): boolean {
  return (
    raw.node.kind === "element" &&
    raw.node.name.uri === WML_NS &&
    (raw.node.name.local === "ins" || raw.node.name.local === "del")
  );
}

function isElement(node: XmlNode, local: string): node is XmlElement {
  return node.kind === "element" && node.name.uri === WML_NS && node.name.local === local;
}
