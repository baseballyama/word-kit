// Verify that user-supplied strings are XML-escaped on write and
// decoded on read across every place the public API accepts text.
// A bug here would either:
//  - corrupt the produced docx (Word reports "needs repair"),
//  - or open the door to a content-control / structured-storage
//    injection by letting a `</w:t><w:r>...` payload smuggle real markup
//    through a placeholder substitution.
//
// Each test feeds the string `"a<b>&\"'c"` (ampersand, less-than,
// greater-than, double quote, single quote) into a different API and
// then checks that:
//  1. validate() returns no errors after a save+reopen, and
//  2. the round-tripped text comes back exactly equal to the input.

import { describe, expect, it } from "vitest";
import {
  addBookmark,
  addBulletList,
  addComment,
  addEndnote,
  addFooter,
  addFootnote,
  addHeader,
  addHyperlink,
  addNumberedList,
  addStyle,
  addTable,
  appendField,
  appendHeading,
  appendMergeField,
  appendParagraph,
  appendTextRun,
  bookmarks,
  commentsPart,
  createDocx,
  endnotesPart,
  footers,
  footnotesPart,
  headers,
  openDocx,
  paragraphs,
  paragraphText,
  setCoreProperties,
  setHyperlinkUrl,
  setTitle,
  text,
  toUint8Array,
  validate,
} from "./index.js";

const NASTY = "a<b>&\"'c";
const NASTY_AMP = "x&y";

function reopen(doc: ReturnType<typeof createDocx>): ReturnType<typeof createDocx> {
  return openDocx(toUint8Array(doc));
}

function expectClean(doc: ReturnType<typeof createDocx>): ReturnType<typeof createDocx> {
  const reopened = reopen(doc);
  const errors = validate(reopened).filter((i) => i.level === "error");
  if (errors.length > 0) {
    throw new Error(
      `validate() returned errors:\n${errors.map((e) => `  [${e.code}] ${e.message}`).join("\n")}`,
    );
  }
  return reopened;
}

describe("XML escaping survives save+reopen", () => {
  it("appendParagraph", () => {
    const doc = createDocx({ paragraphs: [NASTY] });
    const final = expectClean(doc);
    expect(text(final)).toContain(NASTY);
  });

  it("appendTextRun (with formatting)", () => {
    const doc = createDocx({ paragraphs: [""] });
    appendParagraph(doc, "");
    const target = paragraphs(doc).at(-1)!;
    appendTextRun(target, NASTY, { bold: true });
    const final = expectClean(doc);
    expect(text(final)).toContain(NASTY);
  });

  it("appendHeading", () => {
    const doc = createDocx({ paragraphs: [] });
    appendHeading(doc, NASTY, 1);
    const final = expectClean(doc);
    expect(text(final)).toContain(NASTY);
  });

  it("addBulletList", () => {
    const doc = createDocx({ paragraphs: [] });
    addBulletList(doc, [NASTY, NASTY_AMP]);
    const final = expectClean(doc);
    const t = text(final);
    expect(t).toContain(NASTY);
    expect(t).toContain(NASTY_AMP);
  });

  it("addNumberedList", () => {
    const doc = createDocx({ paragraphs: [] });
    addNumberedList(doc, [NASTY]);
    const final = expectClean(doc);
    expect(text(final)).toContain(NASTY);
  });

  it("addTable cell text", () => {
    const doc = createDocx({ paragraphs: [] });
    addTable(doc, [[NASTY, NASTY_AMP]]);
    const final = expectClean(doc);
    const xmlText = new TextDecoder().decode(
      final.opc.parts.get("/word/document.xml")?.data ?? new Uint8Array(),
    );
    // Mandatory escapes (per XML spec): `<` → `&lt;`, `&` → `&amp;`. `>`,
    // `"`, `'` are legal literal in element text and we don't bother
    // escaping them.
    expect(xmlText).toContain("&lt;b>&amp;");
    expect(xmlText).toContain("x&amp;y");
    // No double-escape regression sentinel.
    expect(xmlText).not.toContain("&amp;amp;");
    expect(xmlText).not.toContain("&amp;lt;");
  });

  it("addHeader / addFooter", () => {
    const doc = createDocx({ paragraphs: [] });
    addHeader(doc, NASTY);
    addFooter(doc, NASTY_AMP);
    const final = expectClean(doc);
    expect(headers(final)[0]?.text).toContain(NASTY);
    expect(footers(final)[0]?.text).toContain(NASTY_AMP);
  });

  it("addComment text + author + initials", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    addComment(doc, paragraphs(doc)[0]!, {
      author: NASTY,
      initials: NASTY_AMP,
      text: NASTY,
    });
    const final = expectClean(doc);
    const cp = commentsPart(final);
    expect(cp).toBeDefined();
    // author / initials live as attributes on <w:comment> — they must
    // be decoded back to the original on parse.
    const c = cp!.comments[0]!;
    expect(c.attrs.find((a) => a.name.local === "author")?.value).toBe(NASTY);
    expect(c.attrs.find((a) => a.name.local === "initials")?.value).toBe(NASTY_AMP);
  });

  it("addFootnote / addEndnote text", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    addFootnote(doc, paragraphs(doc)[0]!, NASTY);
    addEndnote(doc, paragraphs(doc)[0]!, NASTY_AMP);
    const final = expectClean(doc);
    expect(footnotesPart(final)).toBeDefined();
    expect(endnotesPart(final)).toBeDefined();
  });

  it("addBookmark name (sanity — names are restricted but assert no breakage)", () => {
    // bookmark names can't contain XML-unsafe chars per Word's UI, so
    // we use a plain underscore-laden name as a sanity baseline.
    const doc = createDocx({ paragraphs: ["body"] });
    addBookmark(doc, "ch_one", paragraphs(doc)[0]!);
    const final = expectClean(doc);
    expect(bookmarks(final).map((b) => b.name)).toContain("ch_one");
  });

  it("addHyperlink url with ampersand", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    addHyperlink(doc, "https://example.com/q?a=1&b=2", "link");
    const final = expectClean(doc);
    const relsBytes = final.opc.parts.get("/word/_rels/document.xml.rels")?.data;
    const relsText = new TextDecoder().decode(relsBytes ?? new Uint8Array());
    // Target attribute must escape the &
    expect(relsText).toContain("https://example.com/q?a=1&amp;b=2");
    expect(relsText).not.toContain("a=1&b=2"); // unescaped form would break parsers
  });

  it("setHyperlinkUrl rewrites with an ampersand-bearing URL", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    addHyperlink(doc, "https://stage.example.com/", "link");
    setHyperlinkUrl(doc, "https://stage.example.com/", "https://example.com/?x=1&y=2");
    const final = expectClean(doc);
    const relsText = new TextDecoder().decode(
      final.opc.parts.get("/word/_rels/document.xml.rels")?.data ?? new Uint8Array(),
    );
    expect(relsText).toContain("https://example.com/?x=1&amp;y=2");
  });

  it("setCoreProperties title / creator / description with nasty chars", () => {
    const doc = createDocx({ paragraphs: [] });
    setCoreProperties(doc, {
      title: NASTY,
      creator: NASTY_AMP,
      description: NASTY,
    });
    const final = expectClean(doc);
    const corePart = final.opc.parts.get("/docProps/core.xml");
    const xml = new TextDecoder().decode(corePart?.data ?? new Uint8Array());
    expect(xml).toContain("a&lt;b>&amp;");
    expect(xml).toContain("x&amp;y");
  });

  it("setTitle convenience also escapes", () => {
    const doc = createDocx({ paragraphs: [] });
    setTitle(doc, NASTY);
    const final = expectClean(doc);
    const corePart = final.opc.parts.get("/docProps/core.xml");
    const xml = new TextDecoder().decode(corePart?.data ?? new Uint8Array());
    expect(xml).toContain("a&lt;b>&amp;");
  });

  it("appendMergeField name is rejected for unsafe identifiers (defensive)", () => {
    const doc = createDocx({ paragraphs: [""] });
    const p = paragraphs(doc)[0]!;
    expect(() => appendMergeField(doc, p, "evil<name>")).toThrow();
  });

  it("appendField allows nasty display text safely", () => {
    const doc = createDocx({ paragraphs: [""] });
    appendField(doc, paragraphs(doc)[0]!, "PAGE", NASTY);
    const final = expectClean(doc);
    const xml = new TextDecoder().decode(
      final.opc.parts.get("/word/document.xml")?.data ?? new Uint8Array(),
    );
    expect(xml).toContain("a&lt;b>&amp;");
  });

  it("addStyle name with nasty chars", () => {
    const doc = createDocx({ paragraphs: [] });
    addStyle(doc, {
      type: "paragraph",
      styleId: "Plain",
      name: NASTY,
    });
    const final = expectClean(doc);
    const xml = new TextDecoder().decode(
      final.opc.parts.get("/word/styles.xml")?.data ?? new Uint8Array(),
    );
    // Attribute value — must escape `<`, `&`, `>`, and `"` (since the
    // attribute itself is enclosed in `"`). `'` may stay literal.
    expect(xml).toContain(`w:val="a&lt;b&gt;&amp;&quot;'c"`);
  });

  it("paragraphText decodes escaped entities back to the original", () => {
    const doc = createDocx({ paragraphs: [NASTY] });
    const final = expectClean(doc);
    expect(paragraphText(paragraphs(final)[0]!)).toContain(NASTY);
  });
});
