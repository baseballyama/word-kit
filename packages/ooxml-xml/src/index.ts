/**
 * Namespace-aware XML parser/serializer for word-kit.
 *
 * @packageDocumentation
 */

export { decodeEntities, encodeAttrValue, encodeText } from "./entities.js";
export { type ParseOptions, parseXml, XmlParseError } from "./parser.js";
export { type SerializeOptions, serializeXml } from "./serializer.js";
export {
  type QName,
  XML_NAMESPACE,
  type XmlAttr,
  type XmlCData,
  type XmlComment,
  type XmlDeclaration,
  type XmlDocument,
  type XmlElement,
  type XmlNode,
  type XmlPI,
  type XmlText,
  XMLNS_NAMESPACE,
} from "./types.js";
