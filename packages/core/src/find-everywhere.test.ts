import { describe, expect, it } from "vitest";
import { Docx } from "./docx.js";

describe("Docx.findTextEverywhere", () => {
  it("returns hits in body, header, footer, and comments", () => {
    const doc = Docx.create({ paragraphs: ["body has {{name}}"] });
    const para = doc.paragraphs[0];
    if (!para) return;
    doc.addHeader("Header: {{name}}");
    doc.addFooter("Footer: {{name}}");
    doc.addComment(para, { author: "R", text: "Update {{name}}." });

    const results = doc.findTextEverywhere("{{name}}");
    const partNames = results.map((r) => r.partName);
    expect(partNames).toContain("/word/document.xml");
    expect(partNames).toContain("/word/header1.xml");
    expect(partNames).toContain("/word/footer1.xml");
    expect(partNames).toContain("/word/comments.xml");

    const total = results.reduce((acc, r) => acc + r.matches.length, 0);
    expect(total).toBe(4);
  });

  it("returns an empty list when nothing matches", () => {
    const doc = Docx.create({ paragraphs: ["hello world"] });
    doc.addHeader("header");
    expect(doc.findTextEverywhere("missing")).toEqual([]);
  });

  it("supports regex queries", () => {
    const doc = Docx.create({ paragraphs: ["x {{a}} y"] });
    doc.addFooter("footer {{b}}");
    const results = doc.findTextEverywhere(/\{\{(\w+)\}\}/g);
    const flat = results.flatMap((r) => r.matches.map((m) => m.captures[0]));
    expect(flat).toEqual(expect.arrayContaining(["a", "b"]));
  });
});

describe("Docx.fromBlob", () => {
  it("loads a docx from a Blob", async () => {
    const doc = Docx.create({ paragraphs: ["Hello"] });
    const blob = doc.toBlob();
    const loaded = await Docx.fromBlob(blob);
    expect(loaded.text).toBe("Hello");
  });
});
