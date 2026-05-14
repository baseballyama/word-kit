import type { XmlAttr, XmlDocument, XmlElement, XmlNode } from "@word-kit/ooxml-xml";
import { WML_NS } from "./namespaces.js";
import type { PassThrough } from "./types.js";

export type WmlStyleType = "paragraph" | "character" | "table" | "numbering";

/**
 * In-memory model of `word/styles.xml`. Each individual `<w:style>` is kept
 * as its raw XML element so round-trip fidelity is exact; convenience
 * accessors (e.g. {@link styleId}, {@link styleBasedOn}) read derived
 * properties from the element on demand.
 */
export interface WmlStylesPart {
  readonly rootAttrs: readonly XmlAttr[];
  docDefaults?: XmlElement;
  latentStyles?: XmlElement;
  styles: XmlElement[];
  /** Children of `<w:styles>` we didn't structure, with original position. */
  extras: PassThrough[];
}

export function parseStylesPart(doc: XmlDocument): WmlStylesPart {
  const root = doc.root;
  if (root.name.uri !== WML_NS || root.name.local !== "styles") {
    throw new Error(`Expected <w:styles> root, got <${qnameToString(root)}>`);
  }
  let docDefaults: XmlElement | undefined;
  let latentStyles: XmlElement | undefined;
  const styles: XmlElement[] = [];
  const extras: PassThrough[] = [];
  for (let i = 0; i < root.children.length; i++) {
    const child = root.children[i];
    if (!child) continue;
    if (child.kind !== "element") {
      extras.push({ slot: i, node: child });
      continue;
    }
    if (isWml(child, "docDefaults")) {
      docDefaults = child;
      continue;
    }
    if (isWml(child, "latentStyles")) {
      latentStyles = child;
      continue;
    }
    if (isWml(child, "style")) {
      styles.push(child);
      continue;
    }
    extras.push({ slot: i, node: child });
  }
  return {
    rootAttrs: root.attrs,
    ...(docDefaults ? { docDefaults } : {}),
    ...(latentStyles ? { latentStyles } : {}),
    styles,
    extras,
  };
}

export function writeStylesPart(part: WmlStylesPart): XmlDocument {
  const recognized: XmlNode[] = [];
  if (part.docDefaults) recognized.push(part.docDefaults);
  if (part.latentStyles) recognized.push(part.latentStyles);
  for (const s of part.styles) recognized.push(s);
  const children = spliceWithExtras(recognized, part.extras);
  const root: XmlElement = {
    kind: "element",
    name: { uri: WML_NS, local: "styles", prefix: "w" },
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

/** Read the `w:styleId` attribute of a `<w:style>` element. */
export function styleId(s: XmlElement): string {
  return wmlAttr(s, "styleId") ?? "";
}

/** Read the `w:type` attribute as a typed value, or `undefined` if missing. */
export function styleType(s: XmlElement): WmlStyleType | undefined {
  const v = wmlAttr(s, "type");
  return v === "paragraph" || v === "character" || v === "table" || v === "numbering"
    ? v
    : undefined;
}

/** Whether the style is marked as `w:default="1"`. */
export function isDefaultStyle(s: XmlElement): boolean {
  const v = wmlAttr(s, "default");
  return v === "1" || v === "true";
}

/** Read `<w:basedOn w:val="..."/>` if present. */
export function styleBasedOn(s: XmlElement): string | undefined {
  return refChildVal(s, "basedOn");
}

/** Read `<w:next w:val="..."/>` if present. */
export function styleNext(s: XmlElement): string | undefined {
  return refChildVal(s, "next");
}

/** Read `<w:link w:val="..."/>` if present. */
export function styleLink(s: XmlElement): string | undefined {
  return refChildVal(s, "link");
}

/** Read `<w:name w:val="..."/>` if present. */
export function styleName(s: XmlElement): string | undefined {
  return refChildVal(s, "name");
}

/** Whether the style has a `<w:qFormat/>` child (Quick Format). */
export function isQFormat(s: XmlElement): boolean {
  return s.children.some(
    (c) => c.kind === "element" && c.name.uri === WML_NS && c.name.local === "qFormat",
  );
}

/** Find a style by id within a parsed styles part. */
export function findStyle(part: WmlStylesPart, id: string): XmlElement | undefined {
  return part.styles.find((s) => styleId(s) === id);
}

function refChildVal(parent: XmlElement, local: string): string | undefined {
  for (const c of parent.children) {
    if (c.kind === "element" && c.name.uri === WML_NS && c.name.local === local) {
      return wmlAttr(c, "val");
    }
  }
  return undefined;
}

function wmlAttr(el: XmlElement, local: string): string | undefined {
  return el.attrs.find((a) => a.name.uri === WML_NS && a.name.local === local)?.value;
}

function isWml(node: XmlNode, local: string): node is XmlElement {
  return node.kind === "element" && node.name.uri === WML_NS && node.name.local === local;
}

function qnameToString(el: XmlElement): string {
  return el.name.prefix ? `${el.name.prefix}:${el.name.local}` : el.name.local;
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
