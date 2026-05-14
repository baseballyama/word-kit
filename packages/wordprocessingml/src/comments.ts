import type { XmlAttr, XmlDocument, XmlElement, XmlNode } from "@word-kit/ooxml-xml";
import { WML_NS } from "./namespaces.js";
import type { PassThrough } from "./types.js";

export interface WmlCommentsPart {
  readonly rootAttrs: readonly XmlAttr[];
  comments: XmlElement[];
  extras: PassThrough[];
}

export function parseCommentsPart(doc: XmlDocument): WmlCommentsPart {
  const root = doc.root;
  if (root.name.uri !== WML_NS || root.name.local !== "comments") {
    throw new Error(`Expected <w:comments> root, got <${root.name.local}>`);
  }
  const comments: XmlElement[] = [];
  const extras: PassThrough[] = [];
  for (let i = 0; i < root.children.length; i++) {
    const child = root.children[i];
    if (!child) continue;
    if (child.kind !== "element") {
      extras.push({ slot: i, node: child });
      continue;
    }
    if (child.name.uri === WML_NS && child.name.local === "comment") {
      comments.push(child);
      continue;
    }
    extras.push({ slot: i, node: child });
  }
  return { rootAttrs: root.attrs, comments, extras };
}

export function writeCommentsPart(part: WmlCommentsPart): XmlDocument {
  const recognized: XmlNode[] = [...part.comments];
  const children = spliceWithExtras(recognized, part.extras);
  const root: XmlElement = {
    kind: "element",
    name: { uri: WML_NS, local: "comments", prefix: "w" },
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

export interface BuildCommentOptions {
  readonly id: number;
  readonly author: string;
  readonly initials?: string;
  readonly date?: string;
  readonly text: string;
}

/** Build a `<w:comment>` element with one paragraph of body text. */
export function buildComment(options: BuildCommentOptions): XmlElement {
  const escapedText = options.text;
  const attrs: XmlAttr[] = [wmlAttr("id", String(options.id)), wmlAttr("author", options.author)];
  if (options.date) attrs.push(wmlAttr("date", options.date));
  if (options.initials) attrs.push(wmlAttr("initials", options.initials));
  const space = /^\s|\s$/.test(escapedText) ? "preserve" : undefined;
  const textAttrs: XmlAttr[] = space
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
    : [];
  return {
    kind: "element",
    name: { uri: WML_NS, local: "comment", prefix: "w" },
    attrs,
    children: [
      {
        kind: "element",
        name: { uri: WML_NS, local: "p", prefix: "w" },
        attrs: [],
        children: [
          {
            kind: "element",
            name: { uri: WML_NS, local: "r", prefix: "w" },
            attrs: [],
            children: [
              {
                kind: "element",
                name: { uri: WML_NS, local: "t", prefix: "w" },
                attrs: textAttrs,
                children: [{ kind: "text", value: escapedText }],
                xmlSpace: space === "preserve" ? "preserve" : "default",
                selfClosing: false,
              },
            ],
            xmlSpace: "default",
            selfClosing: false,
          },
        ],
        xmlSpace: "default",
        selfClosing: false,
      },
    ],
    xmlSpace: "default",
    selfClosing: false,
  };
}

/** Build a `<w:commentRangeStart w:id="N"/>` marker. */
export function buildCommentRangeStart(id: number): XmlElement {
  return wmlEmptyEl("commentRangeStart", [wmlAttr("id", String(id))]);
}

/** Build a `<w:commentRangeEnd w:id="N"/>` marker. */
export function buildCommentRangeEnd(id: number): XmlElement {
  return wmlEmptyEl("commentRangeEnd", [wmlAttr("id", String(id))]);
}

/** Build a `<w:r><w:commentReference w:id="N"/></w:r>` icon run. */
export function buildCommentReferenceRun(id: number): XmlElement {
  return {
    kind: "element",
    name: { uri: WML_NS, local: "r", prefix: "w" },
    attrs: [],
    children: [
      {
        kind: "element",
        name: { uri: WML_NS, local: "rPr", prefix: "w" },
        attrs: [],
        children: [wmlEmptyEl("rStyle", [wmlAttr("val", "CommentReference")])],
        xmlSpace: "default",
        selfClosing: false,
      },
      wmlEmptyEl("commentReference", [wmlAttr("id", String(id))]),
    ],
    xmlSpace: "default",
    selfClosing: false,
  };
}

/** Minimal empty `comments.xml` skeleton. */
export const EMPTY_COMMENTS_XML = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
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
