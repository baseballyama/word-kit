import type { XmlAttr, XmlDocument, XmlElement, XmlNode } from "@word-kit/ooxml-xml";
import { XML_NAMESPACE } from "@word-kit/ooxml-xml";
import { WML_NS } from "./namespaces.js";
import type {
  PassThrough,
  WmlBlock,
  WmlBody,
  WmlDocument as WmlDocumentType,
  WmlInline,
  WmlParagraph,
  WmlRun,
  WmlRunPiece,
  WmlTable,
  WmlTableCell,
  WmlTableRow,
} from "./types.js";

/**
 * Serialize a {@link WmlDocumentType} back to a raw XML document AST,
 * suitable for handing to `@word-kit/ooxml-xml`'s serializer.
 *
 * Pass-through nodes are spliced back into their original slot indices so
 * the output remains structurally equivalent to the input when nothing was
 * edited.
 */
export function writeWmlDocument(wml: WmlDocumentType): XmlDocument {
  const bodyElement = writeBody(wml.body);
  const documentChildren = spliceWithExtras([bodyElement], wml.extras);
  const root: XmlElement = {
    kind: "element",
    name: { uri: WML_NS, local: "document", prefix: "w" },
    attrs: wml.rootAttrs,
    children: documentChildren,
    xmlSpace: "default",
    selfClosing: documentChildren.length === 0,
  };
  return {
    declaration: { version: "1.0", encoding: "UTF-8", standalone: "yes" },
    prologue: [],
    root,
    epilogue: [],
  };
}

function writeBody(body: WmlBody): XmlElement {
  const recognized: XmlNode[] = [];
  for (const block of body.blocks) {
    recognized.push(blockToElement(block));
  }
  if (body.sectPr) recognized.push(body.sectPr);
  const children = spliceWithExtras(recognized, body.extras);
  return {
    kind: "element",
    name: { uri: WML_NS, local: "body", prefix: "w" },
    attrs: [],
    children,
    xmlSpace: "default",
    selfClosing: children.length === 0,
  };
}

function blockToElement(block: WmlBlock): XmlElement {
  switch (block.kind) {
    case "raw":
      return block.node;
    case "paragraph":
      return paragraphToElement(block);
    case "table":
      return tableToElement(block);
    default: {
      const unhandled: never = block;
      throw new Error(`Unhandled block kind: ${(unhandled as { kind: string }).kind}`);
    }
  }
}

function tableToElement(t: WmlTable): XmlElement {
  const recognized: XmlNode[] = [];
  if (t.tblPr) recognized.push(t.tblPr);
  if (t.tblGrid) recognized.push(t.tblGrid);
  for (const row of t.rows) {
    recognized.push(tableRowToElement(row));
  }
  const children = spliceWithExtras(recognized, t.extras);
  return {
    kind: "element",
    name: { uri: WML_NS, local: "tbl", prefix: "w" },
    attrs: [],
    children,
    xmlSpace: "default",
    selfClosing: children.length === 0,
  };
}

function tableRowToElement(row: WmlTableRow): XmlElement {
  const recognized: XmlNode[] = [];
  if (row.trPr) recognized.push(row.trPr);
  for (const cell of row.cells) {
    recognized.push(tableCellToElement(cell));
  }
  const children = spliceWithExtras(recognized, row.extras);
  return {
    kind: "element",
    name: { uri: WML_NS, local: "tr", prefix: "w" },
    attrs: [],
    children,
    xmlSpace: "default",
    selfClosing: children.length === 0,
  };
}

function tableCellToElement(cell: WmlTableCell): XmlElement {
  const recognized: XmlNode[] = [];
  if (cell.tcPr) recognized.push(cell.tcPr);
  for (const p of cell.paragraphs) {
    recognized.push(paragraphToElement(p));
  }
  const children = spliceWithExtras(recognized, cell.extras);
  // A cell with no paragraphs at all is invalid; emit one empty <w:p/>.
  const ensured =
    children.length === 0
      ? [
          {
            kind: "element" as const,
            name: { uri: WML_NS, local: "p", prefix: "w" },
            attrs: [],
            children: [],
            xmlSpace: "default" as const,
            selfClosing: true,
          },
        ]
      : children;
  return {
    kind: "element",
    name: { uri: WML_NS, local: "tc", prefix: "w" },
    attrs: [],
    children: ensured,
    xmlSpace: "default",
    selfClosing: false,
  };
}

/**
 * Convert a {@link WmlParagraph} back into a `<w:p>` XML element. Useful
 * when paragraphs live outside the main document body (e.g. inside
 * `<w:hdr>`/`<w:ftr>` or a `<w:footnote>`).
 */
export function paragraphToElement(p: WmlParagraph): XmlElement {
  const recognized: XmlNode[] = [];
  if (p.pPr) recognized.push(p.pPr);
  for (const inline of p.children) {
    recognized.push(inlineToElement(inline));
  }
  const children = spliceWithExtras(recognized, p.extras);
  return {
    kind: "element",
    name: { uri: WML_NS, local: "p", prefix: "w" },
    attrs: [],
    children,
    xmlSpace: "default",
    selfClosing: children.length === 0,
  };
}

function inlineToElement(inline: WmlInline): XmlElement {
  if (inline.kind === "raw") return inline.node;
  return runToElement(inline);
}

function runToElement(run: WmlRun): XmlElement {
  const recognized: XmlNode[] = [];
  if (run.rPr) recognized.push(run.rPr);
  for (const piece of run.pieces) {
    recognized.push(runPieceToElement(piece));
  }
  const children = spliceWithExtras(recognized, run.extras);
  return {
    kind: "element",
    name: { uri: WML_NS, local: "r", prefix: "w" },
    attrs: [],
    children,
    xmlSpace: "default",
    selfClosing: children.length === 0,
  };
}

function runPieceToElement(piece: WmlRunPiece): XmlElement {
  switch (piece.kind) {
    case "text":
      return wmlTextElement("t", piece.value, piece.preserveSpace);
    case "delText":
      return wmlTextElement("delText", piece.value, piece.preserveSpace);
    case "instrText":
      return wmlTextElement("instrText", piece.value, piece.preserveSpace);
    case "delInstrText":
      return wmlTextElement("delInstrText", piece.value, piece.preserveSpace);
    case "tab":
      return wmlEmptyElement("tab");
    case "break":
      return wmlBreakElement(piece);
    case "noBreakHyphen":
      return wmlEmptyElement("noBreakHyphen");
    case "softHyphen":
      return wmlEmptyElement("softHyphen");
    case "lastRenderedPageBreak":
      return wmlEmptyElement("lastRenderedPageBreak");
    case "fieldChar":
      return piece.raw;
    case "symbol":
      return wmlElement("sym", [
        {
          name: { uri: WML_NS, local: "font", prefix: "w" },
          value: piece.font,
          isNamespaceDecl: false,
        },
        {
          name: { uri: WML_NS, local: "char", prefix: "w" },
          value: piece.char,
          isNamespaceDecl: false,
        },
      ]);
    case "drawing":
    case "pict":
      return piece.node;
    case "raw":
      return piece.node;
    default: {
      const unhandled: never = piece;
      throw new Error(`Unhandled run piece kind: ${(unhandled as { kind: string }).kind}`);
    }
  }
}

function wmlTextElement(local: string, value: string, preserveSpace: boolean): XmlElement {
  const attrs: XmlAttr[] = preserveSpace
    ? [
        {
          name: { uri: XML_NAMESPACE, local: "space", prefix: "xml" },
          value: "preserve",
          isNamespaceDecl: false,
        },
      ]
    : [];
  return {
    kind: "element",
    name: { uri: WML_NS, local, prefix: "w" },
    attrs,
    children: [{ kind: "text", value }],
    xmlSpace: preserveSpace ? "preserve" : "default",
    selfClosing: false,
  };
}

function wmlEmptyElement(local: string): XmlElement {
  return {
    kind: "element",
    name: { uri: WML_NS, local, prefix: "w" },
    attrs: [],
    children: [],
    xmlSpace: "default",
    selfClosing: true,
  };
}

function wmlElement(local: string, attrs: readonly XmlAttr[]): XmlElement {
  return {
    kind: "element",
    name: { uri: WML_NS, local, prefix: "w" },
    attrs,
    children: [],
    xmlSpace: "default",
    selfClosing: true,
  };
}

function wmlBreakElement(piece: {
  breakType?: "page" | "column" | "textWrapping";
  clear?: "none" | "left" | "right" | "all";
}): XmlElement {
  const attrs: XmlAttr[] = [];
  if (piece.breakType) {
    attrs.push({
      name: { uri: WML_NS, local: "type", prefix: "w" },
      value: piece.breakType,
      isNamespaceDecl: false,
    });
  }
  if (piece.clear) {
    attrs.push({
      name: { uri: WML_NS, local: "clear", prefix: "w" },
      value: piece.clear,
      isNamespaceDecl: false,
    });
  }
  return {
    kind: "element",
    name: { uri: WML_NS, local: "br", prefix: "w" },
    attrs,
    children: [],
    xmlSpace: "default",
    selfClosing: true,
  };
}

function spliceWithExtras(recognized: XmlNode[], extras: readonly PassThrough[]): XmlNode[] {
  if (extras.length === 0) return recognized;
  const totalLength = recognized.length + extras.length;
  const result: (XmlNode | undefined)[] = Array.from({ length: totalLength });
  for (const e of extras) {
    if (e.slot >= 0 && e.slot < totalLength) result[e.slot] = e.node;
  }
  let r = 0;
  for (let i = 0; i < totalLength; i++) {
    if (result[i] === undefined) {
      if (r < recognized.length) {
        const node = recognized[r++];
        if (node) result[i] = node;
      }
    }
  }
  while (r < recognized.length) {
    const node = recognized[r++];
    if (node) result.push(node);
  }
  return result.filter((n): n is XmlNode => n !== undefined);
}
