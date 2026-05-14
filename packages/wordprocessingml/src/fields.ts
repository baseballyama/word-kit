import type { XmlAttr, XmlElement } from "@word-kit/ooxml-xml";
import { WML_NS } from "./namespaces.js";

/**
 * Common Word field instructions. Pass a literal string for arbitrary
 * instructions; these are just well-named constants.
 */
export const WORD_FIELDS = {
  PAGE: "PAGE",
  NUMPAGES: "NUMPAGES",
  DATE: "DATE",
  TIME: "TIME",
  FILENAME: "FILENAME",
  AUTHOR: "AUTHOR",
  TITLE: "TITLE",
  SECTION: "SECTION",
  HYPERLINK: "HYPERLINK",
  REF: "REF",
  TOC: "TOC",
} as const;

/**
 * Build a sequence of runs implementing a complex field — `<w:fldChar
 * w:fldCharType="begin"/>` + `<w:instrText>` + `<w:fldChar w:fldCharType=
 * "separate"/>` + display run + `<w:fldChar w:fldCharType="end"/>`. The
 * returned array is intended to be spliced into a paragraph's children as
 * pass-through raw inlines.
 *
 * `displayText` is what Word shows when the field hasn't been refreshed
 * (e.g. "1" for PAGE). Word recomputes the cached value on first open;
 * we just need something there.
 */
export function buildFieldRuns(instruction: string, displayText: string): XmlElement[] {
  return [
    runWithFldChar("begin"),
    runWithInstrText(instruction),
    runWithFldChar("separate"),
    runWithDisplay(displayText),
    runWithFldChar("end"),
  ];
}

/**
 * Build a single `<w:fldSimple w:instr="...">…</w:fldSimple>` element.
 * Simpler than the begin/separate/end runs and adequate for most fields,
 * but Word may rewrite it to the explicit form on save.
 */
export function buildSimpleField(instruction: string, displayText: string): XmlElement {
  return {
    kind: "element",
    name: { uri: WML_NS, local: "fldSimple", prefix: "w" },
    attrs: [wmlAttr("instr", instruction)],
    children: [runWithDisplay(displayText)],
    xmlSpace: "default",
    selfClosing: false,
  };
}

function runWithFldChar(type: "begin" | "separate" | "end"): XmlElement {
  return {
    kind: "element",
    name: { uri: WML_NS, local: "r", prefix: "w" },
    attrs: [],
    children: [
      {
        kind: "element",
        name: { uri: WML_NS, local: "fldChar", prefix: "w" },
        attrs: [wmlAttr("fldCharType", type)],
        children: [],
        xmlSpace: "default",
        selfClosing: true,
      },
    ],
    xmlSpace: "default",
    selfClosing: false,
  };
}

function runWithInstrText(instruction: string): XmlElement {
  const preserveSpace = /^\s|\s$/.test(instruction);
  return {
    kind: "element",
    name: { uri: WML_NS, local: "r", prefix: "w" },
    attrs: [],
    children: [
      {
        kind: "element",
        name: { uri: WML_NS, local: "instrText", prefix: "w" },
        attrs: preserveSpace
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
          : [],
        children: [{ kind: "text", value: instruction }],
        xmlSpace: preserveSpace ? "preserve" : "default",
        selfClosing: false,
      },
    ],
    xmlSpace: "default",
    selfClosing: false,
  };
}

function runWithDisplay(text: string): XmlElement {
  const preserveSpace = /^\s|\s$/.test(text);
  return {
    kind: "element",
    name: { uri: WML_NS, local: "r", prefix: "w" },
    attrs: [],
    children: [
      {
        kind: "element",
        name: { uri: WML_NS, local: "t", prefix: "w" },
        attrs: preserveSpace
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
          : [],
        children: [{ kind: "text", value: text }],
        xmlSpace: preserveSpace ? "preserve" : "default",
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

function escapeXmlText(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/**
 * Footer body XML containing one paragraph with a PAGE field.
 *
 * The displayed text starts at 1 (Word recomputes on open). `prefix` and
 * `suffix` wrap the page number to allow formats like "Page 1" or
 * "- 1 -".
 */
export function buildPageNumberFooterXml(prefix = "", suffix = ""): string {
  const preservePrefix = /^\s|\s$/.test(prefix);
  const preserveSuffix = /^\s|\s$/.test(suffix);
  const prefixRun = prefix
    ? `<w:r><w:t${preservePrefix ? ' xml:space="preserve"' : ""}>${escapeXmlText(prefix)}</w:t></w:r>`
    : "";
  const suffixRun = suffix
    ? `<w:r><w:t${preserveSuffix ? ' xml:space="preserve"' : ""}>${escapeXmlText(suffix)}</w:t></w:r>`
    : "";
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
    "<w:p>",
    '<w:pPr><w:jc w:val="center"/></w:pPr>',
    prefixRun,
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>',
    "<w:r><w:instrText>PAGE</w:instrText></w:r>",
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>',
    "<w:r><w:t>1</w:t></w:r>",
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    suffixRun,
    "</w:p>",
    "</w:ftr>",
  ].join("");
}
