/**
 * Qualified XML name: namespace URI + local part + originating prefix.
 *
 * `prefix` is preserved as authored so a serialized output can use the same
 * shorthand the source did. The empty string represents the default
 * namespace (i.e. `xmlns="..."`).
 */
export interface QName {
  readonly uri: string;
  readonly local: string;
  readonly prefix: string;
}

/**
 * A single XML attribute on an element.
 *
 * `isNamespaceDecl` is `true` when this attribute is an `xmlns` or
 * `xmlns:prefix` declaration; the parser keeps these in the attribute list
 * so output order matches the source.
 */
export interface XmlAttr {
  readonly name: QName;
  readonly value: string;
  readonly isNamespaceDecl: boolean;
}

export type XmlNode = XmlElement | XmlText | XmlCData | XmlComment | XmlPI;

export interface XmlElement {
  readonly kind: "element";
  readonly name: QName;
  readonly attrs: readonly XmlAttr[];
  readonly children: readonly XmlNode[];
  /**
   * Effective value of `xml:space` inside this element (`"preserve"` or
   * `"default"`). Inherited from ancestors if not set locally.
   */
  readonly xmlSpace: "default" | "preserve";
  /** When true the element was authored as `<foo/>`. */
  readonly selfClosing: boolean;
}

export interface XmlText {
  readonly kind: "text";
  readonly value: string;
}

export interface XmlCData {
  readonly kind: "cdata";
  readonly value: string;
}

export interface XmlComment {
  readonly kind: "comment";
  readonly value: string;
}

export interface XmlPI {
  readonly kind: "pi";
  readonly target: string;
  readonly data: string;
}

export interface XmlDeclaration {
  readonly version: string;
  readonly encoding?: string;
  readonly standalone?: "yes" | "no";
}

export interface XmlDocument {
  readonly declaration?: XmlDeclaration;
  /** PIs and comments that appear before the root element. */
  readonly prologue: readonly XmlNode[];
  readonly root: XmlElement;
  /** PIs and comments that appear after the root element. */
  readonly epilogue: readonly XmlNode[];
}

/** Standard XML namespace URI for the `xml:` prefix. */
export const XML_NAMESPACE = "http://www.w3.org/XML/1998/namespace";
/** Standard XML namespace URI for the `xmlns:` prefix. */
export const XMLNS_NAMESPACE = "http://www.w3.org/2000/xmlns/";
