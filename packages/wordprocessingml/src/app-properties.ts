import type { XmlDocument, XmlElement } from "@word-kit/ooxml-xml";

const APP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
const VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

export interface DocumentAppProperties {
  /** Application name (e.g. `"word-kit"`, `"Microsoft Office Word"`). */
  application?: string;
  /** Application version (e.g. `"0.1.0"`). */
  appVersion?: string;
  /** Reported page count. */
  pages?: number;
  /** Reported word count. */
  words?: number;
  /** Reported character count (including spaces). */
  characters?: number;
  /** Reported character count (excluding spaces). */
  charactersWithSpaces?: number;
  /** Reported paragraph count. */
  paragraphs?: number;
  /** Reported line count. */
  lines?: number;
  /** Document template name (e.g. `"Normal.dotm"`). */
  template?: string;
  /** Company / organization. */
  company?: string;
  /** Manager / supervisor. */
  manager?: string;
}

export const APP_PROPERTIES_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.extended-properties+xml";
export const APP_PROPERTIES_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";

export function parseAppProperties(doc: XmlDocument): DocumentAppProperties {
  const root = doc.root;
  if (root.name.uri !== APP_NS || root.name.local !== "Properties") {
    throw new Error(`Expected <Properties> root, got <${root.name.local}>`);
  }
  const out: DocumentAppProperties = {};
  for (const child of root.children) {
    if (child.kind !== "element") continue;
    if (child.name.uri !== APP_NS) continue;
    const value = elementText(child);
    switch (child.name.local) {
      case "Application":
        out.application = value;
        break;
      case "AppVersion":
        out.appVersion = value;
        break;
      case "Pages": {
        const n = toNumber(value);
        if (n !== undefined) out.pages = n;
        break;
      }
      case "Words": {
        const n = toNumber(value);
        if (n !== undefined) out.words = n;
        break;
      }
      case "Characters": {
        const n = toNumber(value);
        if (n !== undefined) out.characters = n;
        break;
      }
      case "CharactersWithSpaces": {
        const n = toNumber(value);
        if (n !== undefined) out.charactersWithSpaces = n;
        break;
      }
      case "Paragraphs": {
        const n = toNumber(value);
        if (n !== undefined) out.paragraphs = n;
        break;
      }
      case "Lines": {
        const n = toNumber(value);
        if (n !== undefined) out.lines = n;
        break;
      }
      case "Template":
        out.template = value;
        break;
      case "Company":
        out.company = value;
        break;
      case "Manager":
        out.manager = value;
        break;
      default:
        break;
    }
  }
  return out;
}

export function writeAppProperties(props: DocumentAppProperties): XmlDocument {
  const children: XmlElement[] = [];
  if (props.template !== undefined) children.push(textElement("Template", props.template));
  if (props.application !== undefined) children.push(textElement("Application", props.application));
  if (props.appVersion !== undefined) children.push(textElement("AppVersion", props.appVersion));
  if (props.pages !== undefined) children.push(textElement("Pages", String(props.pages)));
  if (props.words !== undefined) children.push(textElement("Words", String(props.words)));
  if (props.characters !== undefined)
    children.push(textElement("Characters", String(props.characters)));
  if (props.charactersWithSpaces !== undefined)
    children.push(textElement("CharactersWithSpaces", String(props.charactersWithSpaces)));
  if (props.paragraphs !== undefined)
    children.push(textElement("Paragraphs", String(props.paragraphs)));
  if (props.lines !== undefined) children.push(textElement("Lines", String(props.lines)));
  if (props.company !== undefined) children.push(textElement("Company", props.company));
  if (props.manager !== undefined) children.push(textElement("Manager", props.manager));

  const root: XmlElement = {
    kind: "element",
    name: { uri: APP_NS, local: "Properties", prefix: "" },
    attrs: [
      {
        name: { uri: "http://www.w3.org/2000/xmlns/", local: "xmlns", prefix: "" },
        value: APP_NS,
        isNamespaceDecl: true,
      },
      {
        name: { uri: "http://www.w3.org/2000/xmlns/", local: "vt", prefix: "xmlns" },
        value: VT_NS,
        isNamespaceDecl: true,
      },
    ],
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

function textElement(local: string, value: string): XmlElement {
  return {
    kind: "element",
    name: { uri: APP_NS, local, prefix: "" },
    attrs: [],
    children: [{ kind: "text", value }],
    xmlSpace: "default",
    selfClosing: false,
  };
}

function elementText(el: XmlElement): string {
  let acc = "";
  for (const c of el.children) {
    if (c.kind === "text") acc += c.value;
    else if (c.kind === "cdata") acc += c.value;
  }
  return acc;
}

function toNumber(s: string): number | undefined {
  const n = Number.parseInt(s, 10);
  return Number.isFinite(n) ? n : undefined;
}

export const EMPTY_APP_PROPERTIES_XML = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"/>',
].join("");
