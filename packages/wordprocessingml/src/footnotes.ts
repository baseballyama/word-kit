import type { XmlAttr, XmlDocument, XmlElement, XmlNode } from "@word-kit/ooxml-xml";
import { WML_NS } from "./namespaces.js";
import type { PassThrough } from "./types.js";

export interface WmlFootnotesPart {
  readonly rootAttrs: readonly XmlAttr[];
  footnotes: XmlElement[];
  extras: PassThrough[];
}

export function parseFootnotesPart(doc: XmlDocument): WmlFootnotesPart {
  const root = doc.root;
  if (
    root.name.uri !== WML_NS ||
    (root.name.local !== "footnotes" && root.name.local !== "endnotes")
  ) {
    throw new Error(`Expected <w:footnotes> or <w:endnotes> root, got <${root.name.local}>`);
  }
  const footnotes: XmlElement[] = [];
  const extras: PassThrough[] = [];
  for (let i = 0; i < root.children.length; i++) {
    const child = root.children[i];
    if (!child) continue;
    if (child.kind !== "element") {
      extras.push({ slot: i, node: child });
      continue;
    }
    if (
      child.name.uri === WML_NS &&
      (child.name.local === "footnote" || child.name.local === "endnote")
    ) {
      footnotes.push(child);
      continue;
    }
    extras.push({ slot: i, node: child });
  }
  return { rootAttrs: root.attrs, footnotes, extras };
}

export function writeFootnotesPart(
  part: WmlFootnotesPart,
  rootLocal: "footnotes" | "endnotes" = "footnotes",
): XmlDocument {
  const recognized: XmlNode[] = [...part.footnotes];
  const children = spliceWithExtras(recognized, part.extras);
  const root: XmlElement = {
    kind: "element",
    name: { uri: WML_NS, local: rootLocal, prefix: "w" },
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

export interface BuildFootnoteOptions {
  readonly id: number;
  readonly type?: "separator" | "continuationSeparator" | "normal";
  readonly text?: string;
}

/** Build a `<w:footnote>` (or `<w:endnote>` shape) element. */
export function buildFootnote(
  options: BuildFootnoteOptions,
  elementLocal: "footnote" | "endnote" = "footnote",
): XmlElement {
  const attrs: XmlAttr[] = [wmlAttr("id", String(options.id))];
  if (options.type && options.type !== "normal") attrs.push(wmlAttr("type", options.type));
  if (options.type === "separator") {
    return {
      kind: "element",
      name: { uri: WML_NS, local: elementLocal, prefix: "w" },
      attrs,
      children: [paragraphContaining(wmlEmptyEl("separator", []))],
      xmlSpace: "default",
      selfClosing: false,
    };
  }
  if (options.type === "continuationSeparator") {
    return {
      kind: "element",
      name: { uri: WML_NS, local: elementLocal, prefix: "w" },
      attrs,
      children: [paragraphContaining(wmlEmptyEl("continuationSeparator", []))],
      xmlSpace: "default",
      selfClosing: false,
    };
  }
  const text = options.text ?? "";
  return {
    kind: "element",
    name: { uri: WML_NS, local: elementLocal, prefix: "w" },
    attrs,
    children: [
      paragraphContaining(
        referenceMarkRun(elementLocal === "endnote" ? "endnoteRef" : "footnoteRef", text),
      ),
    ],
    xmlSpace: "default",
    selfClosing: false,
  };
}

function paragraphContaining(child: XmlElement): XmlElement {
  return {
    kind: "element",
    name: { uri: WML_NS, local: "p", prefix: "w" },
    attrs: [],
    children: [
      {
        kind: "element",
        name: { uri: WML_NS, local: "pPr", prefix: "w" },
        attrs: [],
        children: [wmlEmptyEl("pStyle", [wmlAttr("val", "FootnoteText")])],
        xmlSpace: "default",
        selfClosing: false,
      },
      child,
    ],
    xmlSpace: "default",
    selfClosing: false,
  };
}

function referenceMarkRun(refLocal: "footnoteRef" | "endnoteRef", text: string): XmlElement {
  const children: XmlElement[] = [
    {
      kind: "element",
      name: { uri: WML_NS, local: "rPr", prefix: "w" },
      attrs: [],
      children: [wmlEmptyEl("rStyle", [wmlAttr("val", "FootnoteReference")])],
      xmlSpace: "default",
      selfClosing: false,
    },
    wmlEmptyEl(refLocal, []),
  ];
  if (text) {
    const space = /^\s|\s$/.test(text);
    children.push({
      kind: "element",
      name: { uri: WML_NS, local: "t", prefix: "w" },
      attrs: space
        ? [
            {
              name: {
                uri: "http://www.w3.org/XML/1998/namespace",
                local: "space",
                prefix: "xml",
              },
              value: "preserve",
              isNamespaceDecl: false,
            },
          ]
        : [],
      children: [{ kind: "text", value: ` ${text}` }],
      xmlSpace: "preserve",
      selfClosing: false,
    });
  }
  return {
    kind: "element",
    name: { uri: WML_NS, local: "r", prefix: "w" },
    attrs: [],
    children,
    xmlSpace: "default",
    selfClosing: false,
  };
}

/** Build a `<w:r><w:footnoteReference w:id="N"/></w:r>` inline used in body paragraphs. */
export function buildFootnoteReferenceRun(
  id: number,
  refLocal: "footnoteReference" | "endnoteReference" = "footnoteReference",
): XmlElement {
  return {
    kind: "element",
    name: { uri: WML_NS, local: "r", prefix: "w" },
    attrs: [],
    children: [
      {
        kind: "element",
        name: { uri: WML_NS, local: "rPr", prefix: "w" },
        attrs: [],
        children: [wmlEmptyEl("rStyle", [wmlAttr("val", "FootnoteReference")])],
        xmlSpace: "default",
        selfClosing: false,
      },
      wmlEmptyEl(refLocal, [wmlAttr("id", String(id))]),
    ],
    xmlSpace: "default",
    selfClosing: false,
  };
}

/** Minimal footnotes.xml skeleton with the standard separator entries (id=-1 and id=0). */
export const SEED_FOOTNOTES_XML = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
  '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>',
  '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>',
  "</w:footnotes>",
].join("");

/** Minimal endnotes.xml skeleton with standard separators. */
export const SEED_ENDNOTES_XML = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
  '<w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>',
  '<w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>',
  "</w:endnotes>",
].join("");

function wmlAttr(local: string, value: string): XmlAttr {
  return { name: { uri: WML_NS, local, prefix: "w" }, value, isNamespaceDecl: false };
}

function wmlEmptyEl(local: string, attrs: XmlAttr[]): XmlElement {
  return {
    kind: "element",
    name: { uri: WML_NS, local, prefix: "w" },
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
