import type { XmlDocument, XmlElement, XmlNode } from "@word-kit/ooxml-xml";
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
 * Parse a `word/document.xml` XML document AST into a structured WML
 * document. Anything not recognized by this layer is preserved as a raw
 * passthrough so it can be re-emitted unchanged.
 */
export function parseWmlDocument(xmlDoc: XmlDocument): WmlDocumentType {
  const root = xmlDoc.root;
  if (!isWmlElement(root, "document")) {
    throw new Error(`Expected <w:document> root, got <${qnameToString(root)}>`);
  }
  const body = findChildElement(root, "body");
  if (!body) {
    throw new Error("Expected <w:body> in document");
  }
  const extras = collectExtras(root.children, (n) => isWmlElement(n, "body"));
  return {
    rootAttrs: root.attrs,
    body: parseBody(body),
    extras,
  };
}

function parseBody(body: XmlElement): WmlBody {
  const blocks: WmlBlock[] = [];
  const extras: PassThrough[] = [];
  let sectPr: XmlElement | undefined;
  const lastChildIdx = body.children.length - 1;

  for (let i = 0; i < body.children.length; i++) {
    const child = body.children[i];
    if (!child) continue;
    if (child.kind !== "element") {
      // Body should not normally contain text, but preserve if present.
      extras.push({ slot: i, node: child });
      continue;
    }
    if (isWmlElement(child, "p")) {
      blocks.push(parseParagraph(child));
      continue;
    }
    if (isWmlElement(child, "tbl")) {
      blocks.push(parseTable(child));
      continue;
    }
    if (isWmlElement(child, "sectPr") && i === lastChildIdx) {
      sectPr = child;
      continue;
    }
    // Other known (sdt, ...) or unknown: keep as raw block.
    blocks.push({ kind: "raw", node: child });
  }

  return {
    blocks,
    ...(sectPr ? { sectPr } : {}),
    extras,
  };
}

/**
 * Parse a `<w:p>` element into a structured paragraph. Useful for processing
 * paragraphs that appear outside the main document body — for example
 * inside `<w:hdr>`, `<w:ftr>`, `<w:comment>`, or `<w:footnote>`.
 */
export function parseParagraph(p: XmlElement): WmlParagraph {
  let pPr: XmlElement | undefined;
  const children: WmlInline[] = [];
  const extras: PassThrough[] = [];

  for (let i = 0; i < p.children.length; i++) {
    const child = p.children[i];
    if (!child) continue;
    if (child.kind !== "element") {
      extras.push({ slot: i, node: child });
      continue;
    }
    if (isWmlElement(child, "pPr")) {
      pPr = child;
      continue;
    }
    if (isWmlElement(child, "r")) {
      children.push(parseRun(child));
      continue;
    }
    // Anything else under <w:p> (hyperlink, sdt, bookmarkStart/End, ins, del…)
    // is kept verbatim until later milestones structure it.
    children.push({ kind: "raw", node: child });
  }

  return {
    kind: "paragraph",
    ...(pPr ? { pPr } : {}),
    children,
    extras,
  };
}

function parseRun(r: XmlElement): WmlRun {
  let rPr: XmlElement | undefined;
  const pieces: WmlRunPiece[] = [];
  const extras: PassThrough[] = [];

  for (let i = 0; i < r.children.length; i++) {
    const child = r.children[i];
    if (!child) continue;
    if (child.kind !== "element") {
      extras.push({ slot: i, node: child });
      continue;
    }
    if (isWmlElement(child, "rPr")) {
      rPr = child;
      continue;
    }
    const piece = parseRunPiece(child);
    if (piece) {
      pieces.push(piece);
    } else {
      pieces.push({ kind: "raw", node: child });
    }
  }

  return {
    kind: "run",
    ...(rPr ? { rPr } : {}),
    pieces,
    extras,
  };
}

function parseRunPiece(el: XmlElement): WmlRunPiece | undefined {
  if (el.name.uri !== WML_NS) return undefined;
  const attrs = attrMap(el);
  const xmlSpace = attrs.get("xml:space");
  switch (el.name.local) {
    case "t": {
      const text = textContent(el);
      return { kind: "text", value: text, preserveSpace: xmlSpace === "preserve" };
    }
    case "delText": {
      const text = textContent(el);
      return { kind: "delText", value: text, preserveSpace: xmlSpace === "preserve" };
    }
    case "tab":
      return { kind: "tab" };
    case "br": {
      const type = el.attrs.find((a) => a.name.uri === WML_NS && a.name.local === "type")?.value;
      const clear = el.attrs.find((a) => a.name.uri === WML_NS && a.name.local === "clear")?.value;
      return {
        kind: "break",
        ...(type === "page" || type === "column" || type === "textWrapping"
          ? { breakType: type }
          : {}),
        ...(clear === "none" || clear === "left" || clear === "right" || clear === "all"
          ? { clear }
          : {}),
      };
    }
    case "noBreakHyphen":
      return { kind: "noBreakHyphen" };
    case "softHyphen":
      return { kind: "softHyphen" };
    case "lastRenderedPageBreak":
      return { kind: "lastRenderedPageBreak" };
    case "instrText":
      return { kind: "instrText", value: textContent(el), preserveSpace: xmlSpace === "preserve" };
    case "delInstrText":
      return {
        kind: "delInstrText",
        value: textContent(el),
        preserveSpace: xmlSpace === "preserve",
      };
    case "fldChar": {
      const charType = el.attrs.find(
        (a) => a.name.uri === WML_NS && a.name.local === "fldCharType",
      )?.value;
      if (charType === "begin" || charType === "separate" || charType === "end") {
        return { kind: "fieldChar", charType, raw: el };
      }
      return undefined;
    }
    case "sym": {
      const font = el.attrs.find((a) => a.name.uri === WML_NS && a.name.local === "font")?.value;
      const char = el.attrs.find((a) => a.name.uri === WML_NS && a.name.local === "char")?.value;
      if (font && char) {
        return { kind: "symbol", font, char };
      }
      return undefined;
    }
    case "drawing":
      return { kind: "drawing", node: el };
    case "pict":
      return { kind: "pict", node: el };
    default:
      return undefined;
  }
}

function parseTable(tbl: XmlElement): WmlTable {
  let tblPr: XmlElement | undefined;
  let tblGrid: XmlElement | undefined;
  const rows: WmlTableRow[] = [];
  const extras: PassThrough[] = [];
  for (let i = 0; i < tbl.children.length; i++) {
    const child = tbl.children[i];
    if (!child) continue;
    if (child.kind !== "element") {
      extras.push({ slot: i, node: child });
      continue;
    }
    if (isWmlElement(child, "tblPr")) {
      tblPr = child;
      continue;
    }
    if (isWmlElement(child, "tblGrid")) {
      tblGrid = child;
      continue;
    }
    if (isWmlElement(child, "tr")) {
      rows.push(parseTableRow(child));
      continue;
    }
    extras.push({ slot: i, node: child });
  }
  return {
    kind: "table",
    ...(tblPr ? { tblPr } : {}),
    ...(tblGrid ? { tblGrid } : {}),
    rows,
    extras,
  };
}

function parseTableRow(tr: XmlElement): WmlTableRow {
  let trPr: XmlElement | undefined;
  const cells: WmlTableCell[] = [];
  const extras: PassThrough[] = [];
  for (let i = 0; i < tr.children.length; i++) {
    const child = tr.children[i];
    if (!child) continue;
    if (child.kind !== "element") {
      extras.push({ slot: i, node: child });
      continue;
    }
    if (isWmlElement(child, "trPr")) {
      trPr = child;
      continue;
    }
    if (isWmlElement(child, "tc")) {
      cells.push(parseTableCell(child));
      continue;
    }
    extras.push({ slot: i, node: child });
  }
  return {
    ...(trPr ? { trPr } : {}),
    cells,
    extras,
  };
}

function parseTableCell(tc: XmlElement): WmlTableCell {
  let tcPr: XmlElement | undefined;
  const paragraphs: WmlParagraph[] = [];
  const extras: PassThrough[] = [];
  for (let i = 0; i < tc.children.length; i++) {
    const child = tc.children[i];
    if (!child) continue;
    if (child.kind !== "element") {
      extras.push({ slot: i, node: child });
      continue;
    }
    if (isWmlElement(child, "tcPr")) {
      tcPr = child;
      continue;
    }
    if (isWmlElement(child, "p")) {
      paragraphs.push(parseParagraph(child));
      continue;
    }
    // Nested tables, sdt, etc. are kept as extras for now.
    extras.push({ slot: i, node: child });
  }
  return {
    ...(tcPr ? { tcPr } : {}),
    paragraphs,
    extras,
  };
}

function textContent(el: XmlElement): string {
  let acc = "";
  for (const child of el.children) {
    if (child.kind === "text") acc += child.value;
    else if (child.kind === "cdata") acc += child.value;
  }
  return acc;
}

function attrMap(el: XmlElement): Map<string, string> {
  const m = new Map<string, string>();
  for (const a of el.attrs) {
    const key = a.name.prefix ? `${a.name.prefix}:${a.name.local}` : a.name.local;
    m.set(key, a.value);
  }
  return m;
}

function isWmlElement(node: XmlNode, local: string): node is XmlElement {
  return node.kind === "element" && node.name.uri === WML_NS && node.name.local === local;
}

function qnameToString(el: XmlElement | XmlNode): string {
  if (el.kind !== "element") return `(${el.kind})`;
  return el.name.prefix ? `${el.name.prefix}:${el.name.local}` : el.name.local;
}

function findChildElement(parent: XmlElement, local: string): XmlElement | undefined {
  for (const c of parent.children) {
    if (c.kind === "element" && c.name.uri === WML_NS && c.name.local === local) return c;
  }
  return undefined;
}

function collectExtras(nodes: readonly XmlNode[], skip: (n: XmlNode) => boolean): PassThrough[] {
  const extras: PassThrough[] = [];
  for (let i = 0; i < nodes.length; i++) {
    const n = nodes[i];
    if (!n) continue;
    if (skip(n)) continue;
    extras.push({ slot: i, node: n });
  }
  return extras;
}
