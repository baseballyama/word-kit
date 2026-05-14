import { removePart } from "@word-kit/opc";
import { describe, expect, it } from "vitest";
import {
  addComment,
  addEndnote,
  addFooter,
  addFootnote,
  addHeader,
  addImage,
  addTable,
  appendHeading,
  appendParagraph,
  createDocx,
  paragraphs,
  statistics,
  validate,
} from "./docx.js";

describe("Docx.validate", () => {
  it("a fresh docx has no issues", () => {
    const doc = createDocx();
    expect(validate(doc)).toEqual([]);
  });

  it("a kitchen-sink docx with comments and footnotes validates clean", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    const para = paragraphs(doc)[0];
    if (!para) return;
    addComment(doc, para, { author: "R", text: "hi" });
    addFootnote(doc, para, "fn");
    addEndnote(doc, para, "en");
    addHeader(doc, "hdr");
    addFooter(doc, "ftr");
    // MINIMAL_STYLES_XML now includes CommentReference /
    // FootnoteReference / Hyperlink character styles so no warnings
    // either.
    expect(validate(doc)).toEqual([]);
  });

  it("flags missing media when the part is deleted manually", () => {
    const doc = createDocx({ paragraphs: [] });
    addImage(doc, new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0, 0, 0, 0]), {
      widthEmu: 1000,
      heightEmu: 1000,
      contentType: "image/png",
    });
    removePart(doc.opc, "/word/media/image1.png");
    const issues = validate(doc);
    expect(issues.some((i) => i.code === "rel-target-missing" || i.code === "image-missing")).toBe(
      true,
    );
  });
});

describe("Docx.statistics", () => {
  it("counts paragraphs, words, characters, headings", () => {
    const doc = createDocx({ paragraphs: [] });
    appendHeading(doc, "Title", 1);
    appendParagraph(doc, "Hello world");
    appendParagraph(doc, "Second paragraph here.");
    const s = statistics(doc);
    expect(s.paragraphs).toBe(3);
    expect(s.headings).toBe(1);
    expect(s.words).toBe(6); // Title + Hello + world + Second + paragraph + here.
    expect(s.characters).toBeGreaterThan(0);
    expect(s.charactersNoSpaces).toBeLessThan(s.characters);
  });

  it("counts tables and comments", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    addTable(doc, [
      ["a", "b"],
      ["1", "2"],
    ]);
    const para = paragraphs(doc)[0];
    if (!para) return;
    addComment(doc, para, { author: "R", text: "x" });
    const s = statistics(doc);
    expect(s.tables).toBe(1);
    expect(s.comments).toBe(1);
  });

  it("subtracts the standard separator entries from footnote/endnote counts", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    const para = paragraphs(doc)[0];
    if (!para) return;
    addFootnote(doc, para, "one");
    addFootnote(doc, para, "two");
    expect(statistics(doc).footnotes).toBe(2);
  });
});
