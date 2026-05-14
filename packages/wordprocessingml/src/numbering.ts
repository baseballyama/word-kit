import type { XmlAttr, XmlDocument, XmlElement, XmlNode } from "@word-kit/ooxml-xml";
import { WML_NS } from "./namespaces.js";
import type { PassThrough } from "./types.js";

/**
 * In-memory model of `word/numbering.xml`. abstractNum and num entries are
 * kept as raw XML elements; convenience accessors extract derived values.
 */
export interface WmlNumberingPart {
  readonly rootAttrs: readonly XmlAttr[];
  abstractNums: XmlElement[];
  nums: XmlElement[];
  extras: PassThrough[];
}

export function parseNumberingPart(doc: XmlDocument): WmlNumberingPart {
  const root = doc.root;
  if (root.name.uri !== WML_NS || root.name.local !== "numbering") {
    throw new Error(`Expected <w:numbering> root, got <${root.name.local}>`);
  }
  const abstractNums: XmlElement[] = [];
  const nums: XmlElement[] = [];
  const extras: PassThrough[] = [];
  for (let i = 0; i < root.children.length; i++) {
    const child = root.children[i];
    if (!child) continue;
    if (child.kind !== "element") {
      extras.push({ slot: i, node: child });
      continue;
    }
    if (child.name.uri === WML_NS && child.name.local === "abstractNum") {
      abstractNums.push(child);
      continue;
    }
    if (child.name.uri === WML_NS && child.name.local === "num") {
      nums.push(child);
      continue;
    }
    extras.push({ slot: i, node: child });
  }
  return {
    rootAttrs: root.attrs,
    abstractNums,
    nums,
    extras,
  };
}

export function writeNumberingPart(part: WmlNumberingPart): XmlDocument {
  const recognized: XmlNode[] = [...part.abstractNums, ...part.nums];
  const children = spliceWithExtras(recognized, part.extras);
  const root: XmlElement = {
    kind: "element",
    name: { uri: WML_NS, local: "numbering", prefix: "w" },
    attrs: part.rootAttrs,
    children,
    xmlSpace: "default",
    selfClosing: children.length === 0,
  };
  return {
    declaration: { version: "1.0", encoding: "UTF-8", standalone: "yes" },
    prologue: [],
    root,
    epilogue: [],
  };
}

export function abstractNumId(el: XmlElement): number | undefined {
  const v = wmlAttr(el, "abstractNumId");
  if (!v) return undefined;
  const n = Number.parseInt(v, 10);
  return Number.isFinite(n) ? n : undefined;
}

export function numId(el: XmlElement): number | undefined {
  const v = wmlAttr(el, "numId");
  if (!v) return undefined;
  const n = Number.parseInt(v, 10);
  return Number.isFinite(n) ? n : undefined;
}

/** Resolve the `abstractNumId` referenced by a `<w:num>` element. */
export function numAbstractRef(numEl: XmlElement): number | undefined {
  for (const c of numEl.children) {
    if (c.kind === "element" && c.name.uri === WML_NS && c.name.local === "abstractNumId") {
      const v = wmlAttr(c, "val");
      if (!v) return undefined;
      const n = Number.parseInt(v, 10);
      return Number.isFinite(n) ? n : undefined;
    }
  }
  return undefined;
}

function wmlAttr(el: XmlElement, local: string): string | undefined {
  return el.attrs.find((a) => a.name.uri === WML_NS && a.name.local === local)?.value;
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
    if (result[i] === undefined && r < recognized.length) {
      const node = recognized[r++];
      if (node) result[i] = node;
    }
  }
  while (r < recognized.length) {
    const node = recognized[r++];
    if (node) result.push(node);
  }
  return result.filter((n): n is XmlNode => n !== undefined);
}
