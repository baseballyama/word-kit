import type { XmlAttr, XmlElement } from "@word-kit/ooxml-xml";
import { WML_NS } from "./namespaces.js";

export interface BuildHyperlinkOptions {
  /**
   * If set, the hyperlink targets an external URL. The caller is
   * responsible for having created a `r:id` relationship.
   */
  readonly relId?: string;
  /** If set, the hyperlink targets an internal bookmark. */
  readonly anchor?: string;
  /** Optional tooltip text. */
  readonly tooltip?: string;
  /** The run children to wrap (visible text). */
  readonly runs: readonly XmlElement[];
}

/**
 * Construct a `<w:hyperlink>` element wrapping the given visible runs.
 *
 * Either `relId` (external URL via a relationship) or `anchor` (internal
 * bookmark name) must be supplied.
 */
export function buildHyperlink(options: BuildHyperlinkOptions): XmlElement {
  const attrs: XmlAttr[] = [];
  if (options.relId) {
    attrs.push({
      name: {
        uri: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        local: "id",
        prefix: "r",
      },
      value: options.relId,
      isNamespaceDecl: false,
    });
  }
  if (options.anchor) {
    attrs.push(wmlAttr("anchor", options.anchor));
  }
  if (options.tooltip) {
    attrs.push(wmlAttr("tooltip", options.tooltip));
  }
  attrs.push(wmlAttr("history", "1"));
  return {
    kind: "element",
    name: { uri: WML_NS, local: "hyperlink", prefix: "w" },
    attrs,
    children: options.runs,
    xmlSpace: "default",
    selfClosing: false,
  };
}

/** Build a run styled like a Word hyperlink (blue, underlined). */
export function buildHyperlinkRun(text: string): XmlElement {
  const space = /^\s|\s$/.test(text) ? "preserve" : undefined;
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
    name: { uri: WML_NS, local: "r", prefix: "w" },
    attrs: [],
    children: [
      {
        kind: "element",
        name: { uri: WML_NS, local: "rPr", prefix: "w" },
        attrs: [],
        children: [wmlEmptyEl("rStyle", [wmlAttr("val", "Hyperlink")])],
        xmlSpace: "default",
        selfClosing: false,
      },
      {
        kind: "element",
        name: { uri: WML_NS, local: "t", prefix: "w" },
        attrs: textAttrs,
        children: [{ kind: "text", value: text }],
        xmlSpace: space === "preserve" ? "preserve" : "default",
        selfClosing: false,
      },
    ],
    xmlSpace: "default",
    selfClosing: false,
  };
}

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
