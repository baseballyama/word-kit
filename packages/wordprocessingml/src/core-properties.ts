import type { XmlAttr, XmlDocument, XmlElement } from "@word-kit/ooxml-xml";

const CP_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
const DC_NS = "http://purl.org/dc/elements/1.1/";
const DCTERMS_NS = "http://purl.org/dc/terms/";
const XSI_NS = "http://www.w3.org/2001/XMLSchema-instance";

export interface DocumentCoreProperties {
  /** Dublin Core title (`dc:title`). */
  title?: string;
  /** Dublin Core creator (`dc:creator`). */
  creator?: string;
  /** Dublin Core subject (`dc:subject`). */
  subject?: string;
  /** Dublin Core description (`dc:description`). */
  description?: string;
  /** Comma-separated keywords (`cp:keywords`). */
  keywords?: string;
  /** Last person who modified (`cp:lastModifiedBy`). */
  lastModifiedBy?: string;
  /** Document category (`cp:category`). */
  category?: string;
  /** Created timestamp in ISO 8601 (`dcterms:created`). */
  created?: string;
  /** Modified timestamp in ISO 8601 (`dcterms:modified`). */
  modified?: string;
  /** Revision number (`cp:revision`). */
  revision?: string;
  /** Content status (`cp:contentStatus`). */
  contentStatus?: string;
}

const CONTENT_TYPE = "application/vnd.openxmlformats-package.core-properties+xml";
const REL_TYPE =
  "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";

export { CONTENT_TYPE as CORE_PROPERTIES_CONTENT_TYPE, REL_TYPE as CORE_PROPERTIES_REL_TYPE };

/**
 * Parse `docProps/core.xml` into a {@link DocumentCoreProperties} record.
 */
export function parseCoreProperties(doc: XmlDocument): DocumentCoreProperties {
  const root = doc.root;
  if (root.name.uri !== CP_NS || root.name.local !== "coreProperties") {
    throw new Error(`Expected <cp:coreProperties> root, got <${root.name.local}>`);
  }
  const out: DocumentCoreProperties = {};
  for (const child of root.children) {
    if (child.kind !== "element") continue;
    const value = elementText(child);
    if (child.name.uri === DC_NS) {
      if (child.name.local === "title") out.title = value;
      else if (child.name.local === "creator") out.creator = value;
      else if (child.name.local === "subject") out.subject = value;
      else if (child.name.local === "description") out.description = value;
    } else if (child.name.uri === CP_NS) {
      if (child.name.local === "keywords") out.keywords = value;
      else if (child.name.local === "lastModifiedBy") out.lastModifiedBy = value;
      else if (child.name.local === "category") out.category = value;
      else if (child.name.local === "revision") out.revision = value;
      else if (child.name.local === "contentStatus") out.contentStatus = value;
    } else if (child.name.uri === DCTERMS_NS) {
      if (child.name.local === "created") out.created = value;
      else if (child.name.local === "modified") out.modified = value;
    }
  }
  return out;
}

export function writeCoreProperties(props: DocumentCoreProperties): XmlDocument {
  const rootAttrs: XmlAttr[] = [
    nsDecl("cp", CP_NS),
    nsDecl("dc", DC_NS),
    nsDecl("dcterms", DCTERMS_NS),
    nsDecl("xsi", XSI_NS),
  ];

  const children: XmlElement[] = [];
  if (props.title !== undefined) children.push(textElement("dc", DC_NS, "title", props.title));
  if (props.creator !== undefined)
    children.push(textElement("dc", DC_NS, "creator", props.creator));
  if (props.subject !== undefined)
    children.push(textElement("dc", DC_NS, "subject", props.subject));
  if (props.description !== undefined)
    children.push(textElement("dc", DC_NS, "description", props.description));
  if (props.keywords !== undefined)
    children.push(textElement("cp", CP_NS, "keywords", props.keywords));
  if (props.lastModifiedBy !== undefined)
    children.push(textElement("cp", CP_NS, "lastModifiedBy", props.lastModifiedBy));
  if (props.category !== undefined)
    children.push(textElement("cp", CP_NS, "category", props.category));
  if (props.revision !== undefined)
    children.push(textElement("cp", CP_NS, "revision", props.revision));
  if (props.contentStatus !== undefined)
    children.push(textElement("cp", CP_NS, "contentStatus", props.contentStatus));
  if (props.created !== undefined) children.push(dctermsElement("created", props.created));
  if (props.modified !== undefined) children.push(dctermsElement("modified", props.modified));

  const root: XmlElement = {
    kind: "element",
    name: { uri: CP_NS, local: "coreProperties", prefix: "cp" },
    attrs: rootAttrs,
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

function dctermsElement(local: string, value: string): XmlElement {
  return {
    kind: "element",
    name: { uri: DCTERMS_NS, local, prefix: "dcterms" },
    attrs: [
      {
        name: { uri: XSI_NS, local: "type", prefix: "xsi" },
        value: "dcterms:W3CDTF",
        isNamespaceDecl: false,
      },
    ],
    children: [{ kind: "text", value }],
    xmlSpace: "default",
    selfClosing: false,
  };
}

function textElement(prefix: string, uri: string, local: string, value: string): XmlElement {
  return {
    kind: "element",
    name: { uri, local, prefix },
    attrs: [],
    children: [{ kind: "text", value }],
    xmlSpace: "default",
    selfClosing: false,
  };
}

function nsDecl(prefix: string, uri: string): XmlAttr {
  return {
    name: { uri: "http://www.w3.org/2000/xmlns/", local: prefix, prefix: "xmlns" },
    value: uri,
    isNamespaceDecl: true,
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

export const EMPTY_CORE_PROPERTIES_XML = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"/>',
].join("");
