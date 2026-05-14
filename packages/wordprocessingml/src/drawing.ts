import { parseXml, type XmlElement } from "@word-kit/ooxml-xml";

export interface InlineDrawingOptions {
  /** Relationship id (`rId…`) of the embedded image. */
  readonly relId: string;
  /** Image width in EMU (1 inch = 914400 EMU). */
  readonly widthEmu: number;
  /** Image height in EMU. */
  readonly heightEmu: number;
  /** docPr id (must be unique per document; callers typically allocate sequentially). */
  readonly docPrId: number;
  /** docPr name (defaults to `Picture {id}`). */
  readonly name?: string;
  /** Alt text on the docPr element. */
  readonly altText?: string;
}

const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/**
 * Construct a `<w:drawing>` element for an inline image. The result mirrors
 * what Microsoft Word emits for a freshly-inserted picture (including the
 * `noChangeAspect` lock so users cannot accidentally stretch the image).
 */
export function buildInlineDrawing(options: InlineDrawingOptions): XmlElement {
  const name = options.name ?? `Picture ${options.docPrId}`;
  const descr = options.altText
    ? ` descr="${escapeAttr(options.altText)}"`
    : "";
  const xml = parseXml(
    [
      `<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" `,
      `xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:pic="${PIC_NS}" xmlns:r="${R_NS}">`,
      `<wp:inline distT="0" distB="0" distL="0" distR="0">`,
      `<wp:extent cx="${options.widthEmu}" cy="${options.heightEmu}"/>`,
      `<wp:effectExtent l="0" t="0" r="0" b="0"/>`,
      `<wp:docPr id="${options.docPrId}" name="${escapeAttr(name)}"${descr}/>`,
      `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>`,
      `<a:graphic>`,
      `<a:graphicData uri="${PIC_NS}">`,
      `<pic:pic>`,
      `<pic:nvPicPr>`,
      `<pic:cNvPr id="0" name="${escapeAttr(name)}"/>`,
      `<pic:cNvPicPr><a:picLocks noChangeAspect="1" noChangeArrowheads="1"/></pic:cNvPicPr>`,
      `</pic:nvPicPr>`,
      `<pic:blipFill>`,
      `<a:blip r:embed="${escapeAttr(options.relId)}"/>`,
      `<a:stretch><a:fillRect/></a:stretch>`,
      `</pic:blipFill>`,
      `<pic:spPr>`,
      `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${options.widthEmu}" cy="${options.heightEmu}"/></a:xfrm>`,
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`,
      `</pic:spPr>`,
      `</pic:pic>`,
      `</a:graphicData>`,
      `</a:graphic>`,
      `</wp:inline>`,
      `</w:drawing>`,
    ].join(""),
  );
  return xml.root;
}

function escapeAttr(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/**
 * Detect the MIME type of a binary image from its leading bytes.
 *
 * Returns `undefined` if the format cannot be identified.
 */
export function sniffImageContentType(bytes: Uint8Array): string | undefined {
  if (bytes.length >= 8 && bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) {
    return "image/png";
  }
  if (bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) {
    return "image/jpeg";
  }
  if (
    bytes.length >= 6 &&
    bytes[0] === 0x47 &&
    bytes[1] === 0x49 &&
    bytes[2] === 0x46 &&
    bytes[3] === 0x38
  ) {
    return "image/gif";
  }
  if (bytes.length >= 2 && bytes[0] === 0x42 && bytes[1] === 0x4d) {
    return "image/bmp";
  }
  if (
    bytes.length >= 4 &&
    bytes[0] === 0x49 &&
    bytes[1] === 0x49 &&
    bytes[2] === 0x2a &&
    bytes[3] === 0x00
  ) {
    return "image/tiff";
  }
  if (
    bytes.length >= 4 &&
    bytes[0] === 0x4d &&
    bytes[1] === 0x4d &&
    bytes[2] === 0x00 &&
    bytes[3] === 0x2a
  ) {
    return "image/tiff";
  }
  if (bytes.length >= 64) {
    const head = new TextDecoder().decode(bytes.subarray(0, Math.min(256, bytes.length)));
    if (/<svg[\s>]/i.test(head) || /^<\?xml[^>]*>\s*<svg/i.test(head)) {
      return "image/svg+xml";
    }
  }
  return undefined;
}

const EXTENSION_FOR_CONTENT_TYPE: Record<string, string> = {
  "image/png": "png",
  "image/jpeg": "jpg",
  "image/gif": "gif",
  "image/bmp": "bmp",
  "image/tiff": "tiff",
  "image/svg+xml": "svg",
};

/** Suggest a file extension (no leading dot) for an image content type. */
export function extensionForImageContentType(contentType: string): string {
  return EXTENSION_FOR_CONTENT_TYPE[contentType] ?? "bin";
}
