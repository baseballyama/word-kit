import { encodeAttrValue, encodeText } from "./entities.js";
import type { XmlDocument, XmlElement, XmlNode } from "./types.js";

export interface SerializeOptions {
  /**
   * Line separator to insert after the XML declaration. Defaults to `\r\n`
   * which matches what Word emits.
   */
  readonly newlineAfterDeclaration?: "\r\n" | "\n" | "";
}

export function serializeXml(doc: XmlDocument, options: SerializeOptions = {}): string {
  const nl = options.newlineAfterDeclaration ?? "\r\n";
  const parts: string[] = [];
  if (doc.declaration) {
    const d = doc.declaration;
    parts.push(`<?xml version="${d.version}"`);
    if (d.encoding !== undefined) parts.push(` encoding="${d.encoding}"`);
    if (d.standalone !== undefined) parts.push(` standalone="${d.standalone}"`);
    parts.push("?>");
    parts.push(nl);
  }
  for (const node of doc.prologue) {
    parts.push(serializeNode(node));
  }
  parts.push(serializeElement(doc.root));
  for (const node of doc.epilogue) {
    parts.push(serializeNode(node));
  }
  return parts.join("");
}

function serializeElement(el: XmlElement): string {
  const tagName = el.name.prefix ? `${el.name.prefix}:${el.name.local}` : el.name.local;
  const attrStr = el.attrs
    .map((a) => {
      const attrTag = a.name.prefix ? `${a.name.prefix}:${a.name.local}` : a.name.local;
      return ` ${attrTag}="${encodeAttrValue(a.value, '"')}"`;
    })
    .join("");
  if (el.children.length === 0 && el.selfClosing) {
    return `<${tagName}${attrStr}/>`;
  }
  if (el.children.length === 0) {
    return `<${tagName}${attrStr}></${tagName}>`;
  }
  const inner = el.children.map(serializeNode).join("");
  return `<${tagName}${attrStr}>${inner}</${tagName}>`;
}

function serializeNode(node: XmlNode): string {
  switch (node.kind) {
    case "element":
      return serializeElement(node);
    case "text":
      return encodeText(node.value);
    case "cdata":
      return `<![CDATA[${node.value}]]>`;
    case "comment":
      return `<!--${node.value}-->`;
    case "pi":
      return node.data ? `<?${node.target} ${node.data}?>` : `<?${node.target}?>`;
    default:
      // Exhaustive switch should never reach here.
      return "";
  }
}
