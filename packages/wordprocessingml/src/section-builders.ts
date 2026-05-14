import type { XmlAttr, XmlElement } from "@word-kit/ooxml-xml";
import { WML_NS } from "./namespaces.js";

export type HeaderFooterType = "default" | "first" | "even";

/** Build a `<w:headerReference>` (`local="headerReference"`) element. */
export function buildHeaderReference(type: HeaderFooterType, relId: string): XmlElement {
  return buildReference("headerReference", type, relId);
}

/** Build a `<w:footerReference>` element. */
export function buildFooterReference(type: HeaderFooterType, relId: string): XmlElement {
  return buildReference("footerReference", type, relId);
}

function buildReference(local: string, type: HeaderFooterType, relId: string): XmlElement {
  return {
    kind: "element",
    name: { uri: WML_NS, local, prefix: "w" },
    attrs: [
      wmlAttr("type", type),
      {
        name: {
          uri: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
          local: "id",
          prefix: "r",
        },
        value: relId,
        isNamespaceDecl: false,
      },
    ],
    children: [],
    xmlSpace: "default",
    selfClosing: true,
  };
}

export interface PageSize {
  /** Page width in twips. Letter = 12240, A4 = 11906. */
  readonly widthTwips: number;
  /** Page height in twips. Letter = 15840, A4 = 16838. */
  readonly heightTwips: number;
  /** Orientation. Defaults to `"portrait"`. */
  readonly orientation?: "portrait" | "landscape";
}

export interface PageMargins {
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
  readonly left: number;
  readonly header: number;
  readonly footer: number;
  readonly gutter: number;
}

/** Build a `<w:pgSz>` element. */
export function buildPgSz(size: PageSize): XmlElement {
  const attrs = [wmlAttr("w", String(size.widthTwips)), wmlAttr("h", String(size.heightTwips))];
  if (size.orientation === "landscape") {
    attrs.push(wmlAttr("orient", "landscape"));
  }
  return {
    kind: "element",
    name: { uri: WML_NS, local: "pgSz", prefix: "w" },
    attrs,
    children: [],
    xmlSpace: "default",
    selfClosing: true,
  };
}

/** Build a `<w:pgMar>` element. */
export function buildPgMar(margins: PageMargins): XmlElement {
  return {
    kind: "element",
    name: { uri: WML_NS, local: "pgMar", prefix: "w" },
    attrs: [
      wmlAttr("top", String(margins.top)),
      wmlAttr("right", String(margins.right)),
      wmlAttr("bottom", String(margins.bottom)),
      wmlAttr("left", String(margins.left)),
      wmlAttr("header", String(margins.header)),
      wmlAttr("footer", String(margins.footer)),
      wmlAttr("gutter", String(margins.gutter)),
    ],
    children: [],
    xmlSpace: "default",
    selfClosing: true,
  };
}

/** Build a `<w:sectPr>` with the given page size and margins. */
export function buildSectPr(opts: {
  pgSz?: PageSize;
  pgMar?: PageMargins;
  headerRefs?: ReadonlyArray<{ type: HeaderFooterType; relId: string }>;
  footerRefs?: ReadonlyArray<{ type: HeaderFooterType; relId: string }>;
}): XmlElement {
  const children: XmlElement[] = [];
  for (const r of opts.headerRefs ?? []) {
    children.push(buildHeaderReference(r.type, r.relId));
  }
  for (const r of opts.footerRefs ?? []) {
    children.push(buildFooterReference(r.type, r.relId));
  }
  if (opts.pgSz) children.push(buildPgSz(opts.pgSz));
  if (opts.pgMar) children.push(buildPgMar(opts.pgMar));
  return {
    kind: "element",
    name: { uri: WML_NS, local: "sectPr", prefix: "w" },
    attrs: [],
    children,
    xmlSpace: "default",
    selfClosing: false,
  };
}

/**
 * Replace or add the `<w:pgSz>` child of a `<w:sectPr>`.
 * Mutates the input element.
 */
export function setSectPrPageSize(sectPr: XmlElement, size: PageSize): void {
  const newEl = buildPgSz(size);
  const idx = sectPr.children.findIndex(
    (c) => c.kind === "element" && c.name.uri === WML_NS && c.name.local === "pgSz",
  );
  const children = sectPr.children as XmlElement["children"][number][];
  if (idx >= 0) {
    (children as XmlElement[])[idx] = newEl;
  } else {
    (children as XmlElement[]).push(newEl);
  }
}

/** Replace or add the `<w:pgMar>` child of a `<w:sectPr>`. */
export function setSectPrPageMargins(sectPr: XmlElement, margins: PageMargins): void {
  const newEl = buildPgMar(margins);
  const idx = sectPr.children.findIndex(
    (c) => c.kind === "element" && c.name.uri === WML_NS && c.name.local === "pgMar",
  );
  const children = sectPr.children as XmlElement[];
  if (idx >= 0) {
    children[idx] = newEl;
  } else {
    children.push(newEl);
  }
}

/** Append a header/footer reference to a `<w:sectPr>`. */
export function addSectPrHeaderRef(
  sectPr: XmlElement,
  type: HeaderFooterType,
  relId: string,
): void {
  (sectPr.children as XmlElement[]).push(buildHeaderReference(type, relId));
}

/** Append a footer reference to a `<w:sectPr>`. */
export function addSectPrFooterRef(
  sectPr: XmlElement,
  type: HeaderFooterType,
  relId: string,
): void {
  (sectPr.children as XmlElement[]).push(buildFooterReference(type, relId));
}

/** Minimal `header1.xml` body containing a single paragraph with text. */
export function buildHeaderXml(text: string): string {
  const escaped = text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const space = /^\s|\s$/.test(text) ? ' xml:space="preserve"' : "";
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
    `<w:p><w:r><w:t${space}>${escaped}</w:t></w:r></w:p>`,
    "</w:hdr>",
  ].join("");
}

/** Minimal `footer1.xml` body containing a single paragraph with text. */
export function buildFooterXml(text: string): string {
  const escaped = text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const space = /^\s|\s$/.test(text) ? ' xml:space="preserve"' : "";
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
    `<w:p><w:r><w:t${space}>${escaped}</w:t></w:r></w:p>`,
    "</w:ftr>",
  ].join("");
}

/** Standard Letter page size in twips. */
export const PAGE_SIZE_LETTER: PageSize = {
  widthTwips: 12240,
  heightTwips: 15840,
  orientation: "portrait",
};

/** Standard A4 page size in twips. */
export const PAGE_SIZE_A4: PageSize = {
  widthTwips: 11906,
  heightTwips: 16838,
  orientation: "portrait",
};

/** Word's "Normal" margins in twips (1 inch top/bottom, 1 inch left/right). */
export const MARGINS_NORMAL: PageMargins = {
  top: 1440,
  right: 1440,
  bottom: 1440,
  left: 1440,
  header: 720,
  footer: 720,
  gutter: 0,
};

function wmlAttr(local: string, value: string): XmlAttr {
  return { name: { uri: WML_NS, local, prefix: "w" }, value, isNamespaceDecl: false };
}
