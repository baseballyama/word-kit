import { describe, expect, it } from "vitest";
import {
  addComment,
  addFooter,
  addHeader,
  createDocx,
  findTextEverywhere,
  fromBlob,
  paragraphs,
  text,
  toBlob,
} from "./docx.js";

describe("Docx.findTextEverywhere", () => {
  it("returns hits in body, header, footer, and comments", () => {
    const doc = createDocx({ paragraphs: ["body has {{name}}"] });
    const para = paragraphs(doc)[0];
    if (!para) return;
    addHeader(doc, "Header: {{name}}");
    addFooter(doc, "Footer: {{name}}");
    addComment(doc, para, { author: "R", text: "Update {{name}}." });

    const results = findTextEverywhere(doc, "{{name}}");
    const partNames = results.map((r) => r.partName);
    expect(partNames).toContain("/word/document.xml");
    expect(partNames).toContain("/word/header1.xml");
    expect(partNames).toContain("/word/footer1.xml");
    expect(partNames).toContain("/word/comments.xml");

    const total = results.reduce((acc, r) => acc + r.matches.length, 0);
    expect(total).toBe(4);
  });

  it("returns an empty list when nothing matches", () => {
    const doc = createDocx({ paragraphs: ["hello world"] });
    addHeader(doc, "header");
    expect(findTextEverywhere(doc, "missing")).toEqual([]);
  });

  it("supports regex queries", () => {
    const doc = createDocx({ paragraphs: ["x {{a}} y"] });
    addFooter(doc, "footer {{b}}");
    const results = findTextEverywhere(doc, /\{\{(\w+)\}\}/g);
    const flat = results.flatMap((r) => r.matches.map((m) => m.captures[0]));
    expect(flat).toEqual(expect.arrayContaining(["a", "b"]));
  });
});

describe("Docx.fromBlob", () => {
  it("loads a docx from a Blob", async () => {
    const doc = createDocx({ paragraphs: ["Hello"] });
    const blob = toBlob(doc);
    const loaded = await fromBlob(blob);
    expect(text(loaded)).toBe("Hello");
  });
});
