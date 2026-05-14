import type { XmlAttr, XmlElement } from "@word-kit/ooxml-xml";
import { WML_NS } from "./namespaces.js";
import type { WmlStyleType } from "./styles.js";

export interface BuildStyleOptions {
  readonly type: WmlStyleType;
  readonly styleId: string;
  readonly name?: string;
  readonly basedOn?: string;
  readonly next?: string;
  readonly link?: string;
  readonly default?: boolean;
  readonly customStyle?: boolean;
  readonly qFormat?: boolean;
  readonly uiPriority?: number;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly fontSizeHalfPoints?: number;
  /** Hex RGB color value (no leading `#`). */
  readonly color?: string;
  /** Justification: left / center / right / both / distribute. */
  readonly alignment?: "left" | "center" | "right" | "both" | "distribute";
}

/**
 * Construct a `<w:style>` XML element from a plain options record. The
 * resulting element is what {@link WmlStylesPart.styles} expects.
 */
export function buildStyle(options: BuildStyleOptions): XmlElement {
  const attrs: XmlAttr[] = [wmlAttr("type", options.type), wmlAttr("styleId", options.styleId)];
  if (options.default) attrs.push(wmlAttr("default", "1"));
  if (options.customStyle) attrs.push(wmlAttr("customStyle", "1"));

  const children: XmlElement[] = [];
  if (options.name) children.push(wmlSelfClosing("name", [wmlAttr("val", options.name)]));
  if (options.basedOn) children.push(wmlSelfClosing("basedOn", [wmlAttr("val", options.basedOn)]));
  if (options.next) children.push(wmlSelfClosing("next", [wmlAttr("val", options.next)]));
  if (options.link) children.push(wmlSelfClosing("link", [wmlAttr("val", options.link)]));
  if (options.uiPriority !== undefined) {
    children.push(wmlSelfClosing("uiPriority", [wmlAttr("val", String(options.uiPriority))]));
  }
  if (options.qFormat) children.push(wmlSelfClosing("qFormat", []));

  const pPrChildren: XmlElement[] = [];
  if (options.alignment) {
    pPrChildren.push(wmlSelfClosing("jc", [wmlAttr("val", options.alignment)]));
  }
  if (pPrChildren.length > 0) {
    children.push(wmlElement("pPr", [], pPrChildren));
  }

  const rPrChildren: XmlElement[] = [];
  if (options.bold) rPrChildren.push(wmlSelfClosing("b", []));
  if (options.italic) rPrChildren.push(wmlSelfClosing("i", []));
  if (options.fontSizeHalfPoints !== undefined) {
    rPrChildren.push(wmlSelfClosing("sz", [wmlAttr("val", String(options.fontSizeHalfPoints))]));
    rPrChildren.push(wmlSelfClosing("szCs", [wmlAttr("val", String(options.fontSizeHalfPoints))]));
  }
  if (options.color) {
    rPrChildren.push(wmlSelfClosing("color", [wmlAttr("val", options.color)]));
  }
  if (rPrChildren.length > 0) {
    children.push(wmlElement("rPr", [], rPrChildren));
  }

  return {
    kind: "element",
    name: { uri: WML_NS, local: "style", prefix: "w" },
    attrs,
    children,
    xmlSpace: "default",
    selfClosing: children.length === 0,
  };
}

/**
 * Minimal `styles.xml` content: docDefaults + Normal paragraph style +
 * default character/table/numbering placeholders. Matches what Word emits
 * when you File → New → save.
 */
export const MINIMAL_STYLES_XML = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
  "<w:docDefaults>",
  "<w:rPrDefault>",
  "<w:rPr>",
  '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Times New Roman"/>',
  '<w:sz w:val="22"/>',
  '<w:szCs w:val="22"/>',
  '<w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/>',
  "</w:rPr>",
  "</w:rPrDefault>",
  "<w:pPrDefault>",
  "<w:pPr>",
  '<w:spacing w:after="160" w:line="259" w:lineRule="auto"/>',
  "</w:pPr>",
  "</w:pPrDefault>",
  "</w:docDefaults>",
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">',
  '<w:name w:val="Normal"/>',
  "<w:qFormat/>",
  "</w:style>",
  '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">',
  '<w:name w:val="Default Paragraph Font"/>',
  '<w:uiPriority w:val="1"/>',
  "<w:semiHidden/>",
  "<w:unhideWhenUsed/>",
  "</w:style>",
  '<w:style w:type="table" w:default="1" w:styleId="TableNormal">',
  '<w:name w:val="Normal Table"/>',
  '<w:uiPriority w:val="99"/>',
  "<w:semiHidden/>",
  "<w:unhideWhenUsed/>",
  "<w:tblPr>",
  '<w:tblInd w:w="0" w:type="dxa"/>',
  "<w:tblCellMar>",
  '<w:top w:w="0" w:type="dxa"/>',
  '<w:left w:w="108" w:type="dxa"/>',
  '<w:bottom w:w="0" w:type="dxa"/>',
  '<w:right w:w="108" w:type="dxa"/>',
  "</w:tblCellMar>",
  "</w:tblPr>",
  "</w:style>",
  '<w:style w:type="numbering" w:default="1" w:styleId="NoList">',
  '<w:name w:val="No List"/>',
  '<w:uiPriority w:val="99"/>',
  "<w:semiHidden/>",
  "<w:unhideWhenUsed/>",
  "</w:style>",
  "</w:styles>",
].join("");

function wmlAttr(local: string, value: string): XmlAttr {
  return { name: { uri: WML_NS, local, prefix: "w" }, value, isNamespaceDecl: false };
}

function wmlSelfClosing(local: string, attrs: XmlAttr[]): XmlElement {
  return {
    kind: "element",
    name: { uri: WML_NS, local, prefix: "w" },
    attrs,
    children: [],
    xmlSpace: "default",
    selfClosing: true,
  };
}

function wmlElement(local: string, attrs: XmlAttr[], children: XmlElement[]): XmlElement {
  return {
    kind: "element",
    name: { uri: WML_NS, local, prefix: "w" },
    attrs,
    children,
    xmlSpace: "default",
    selfClosing: false,
  };
}
