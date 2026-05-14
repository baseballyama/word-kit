import { describe, expect, it } from "vitest";
import { Docx } from "./docx.js";

describe("Docx.validate", () => {
  it("a fresh docx has no issues", () => {
    const doc = Docx.create();
    expect(doc.validate()).toEqual([]);
  });

  it("a kitchen-sink docx with comments and footnotes validates clean", () => {
    const doc = Docx.create({ paragraphs: ["body"] });
    const para = doc.paragraphs[0];
    if (!para) return;
    doc.addComment(para, { author: "R", text: "hi" });
    doc.addFootnote(para, "fn");
    doc.addEndnote(para, "en");
    doc.addHeader("hdr");
    doc.addFooter("ftr");
    // MINIMAL_STYLES_XML now includes CommentReference /
    // FootnoteReference / Hyperlink character styles so no warnings
    // either.
    expect(doc.validate()).toEqual([]);
  });

  it("flags missing media when the part is deleted manually", () => {
    const doc = Docx.create({ paragraphs: [] });
    doc.addImage(new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0, 0, 0, 0]), {
      widthEmu: 1000,
      heightEmu: 1000,
      contentType: "image/png",
    });
    doc.opc.removePart("/word/media/image1.png");
    const issues = doc.validate();
    expect(issues.some((i) => i.code === "rel-target-missing" || i.code === "image-missing")).toBe(
      true,
    );
  });
});

describe("Docx.statistics", () => {
  it("counts paragraphs, words, characters, headings", () => {
    const doc = Docx.create({ paragraphs: [] });
    doc.appendHeading("Title", 1);
    doc.appendParagraph("Hello world");
    doc.appendParagraph("Second paragraph here.");
    const s = doc.statistics;
    expect(s.paragraphs).toBe(3);
    expect(s.headings).toBe(1);
    expect(s.words).toBe(6); // Title + Hello + world + Second + paragraph + here.
    expect(s.characters).toBeGreaterThan(0);
    expect(s.charactersNoSpaces).toBeLessThan(s.characters);
  });

  it("counts tables and comments", () => {
    const doc = Docx.create({ paragraphs: ["body"] });
    doc.addTable([
      ["a", "b"],
      ["1", "2"],
    ]);
    const para = doc.paragraphs[0];
    if (!para) return;
    doc.addComment(para, { author: "R", text: "x" });
    const s = doc.statistics;
    expect(s.tables).toBe(1);
    expect(s.comments).toBe(1);
  });

  it("subtracts the standard separator entries from footnote/endnote counts", () => {
    const doc = Docx.create({ paragraphs: ["body"] });
    const para = doc.paragraphs[0];
    if (!para) return;
    doc.addFootnote(para, "one");
    doc.addFootnote(para, "two");
    expect(doc.statistics.footnotes).toBe(2);
  });
});
