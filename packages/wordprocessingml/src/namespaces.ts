/**
 * Canonical namespace URIs used in WordprocessingML.
 *
 * Prefix bindings vary by document, but URIs are stable. Always compare on
 * `QName.uri`, never on prefix.
 */
export const WML_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
export const DML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
export const DML_WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
export const DML_PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture";
export const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";
export const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

export const WML_RELATIONSHIPS = {
  officeDocument:
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
  styles: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
  numbering: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
  settings: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
  webSettings: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings",
  fontTable: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable",
  theme: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
  header: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
  footer: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
  comments: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
  footnotes: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
  endnotes: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
  hyperlink: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
  image: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
} as const;

export const WML_CONTENT_TYPES = {
  document: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
  styles: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
  numbering: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
  settings: "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
  webSettings: "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml",
  fontTable: "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
  theme: "application/vnd.openxmlformats-officedocument.theme+xml",
  header: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
  footer: "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
  comments: "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
  footnotes: "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
  endnotes: "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
} as const;
