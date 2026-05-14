import { getPart } from "@word-kit/opc";
import { describe, expect, it } from "vitest";
import { Docx } from "./docx.js";

describe("Docx.replaceTextEverywhere", () => {
  it("replaces text in headers and footers", () => {
    const doc = Docx.create({ paragraphs: ["body has {{name}}"] });
    doc.addHeader("Header: {{name}}");
    doc.addFooter("Footer: {{name}}");
    const count = doc.replaceTextEverywhere("{{name}}", "山田太郎");
    expect(count).toBe(3); // body + header + footer

    const bytes = doc.toUint8Array();
    const reopened = Docx.open(bytes);
    expect(reopened.text).toContain("body has 山田太郎");
    const headerXml = new TextDecoder().decode(
      getPart(reopened.opc, "/word/header1.xml")?.data ?? new Uint8Array(),
    );
    const footerXml = new TextDecoder().decode(
      getPart(reopened.opc, "/word/footer1.xml")?.data ?? new Uint8Array(),
    );
    expect(headerXml).toContain("Header: 山田太郎");
    expect(footerXml).toContain("Footer: 山田太郎");
  });

  it("replaces text inside comment bodies", () => {
    const doc = Docx.create({ paragraphs: ["body"] });
    const para = doc.paragraphs[0];
    if (!para) return;
    doc.addComment(para, { author: "R", text: "Please update {{name}}." });
    const count = doc.replaceTextEverywhere("{{name}}", "Alice");
    expect(count).toBeGreaterThan(0);
    const bytes = doc.toUint8Array();
    const reopened = Docx.open(bytes);
    const commentsXml = new TextDecoder().decode(
      getPart(reopened.opc, "/word/comments.xml")?.data ?? new Uint8Array(),
    );
    expect(commentsXml).toContain("Please update Alice.");
  });

  it("replaces text inside footnotes too", () => {
    const doc = Docx.create({ paragraphs: ["body"] });
    const para = doc.paragraphs[0];
    if (!para) return;
    doc.addFootnote(para, "Footnote with {{name}}");
    doc.replaceTextEverywhere("{{name}}", "Bob");
    const reopened = Docx.open(doc.toUint8Array());
    const fnXml = new TextDecoder().decode(
      getPart(reopened.opc, "/word/footnotes.xml")?.data ?? new Uint8Array(),
    );
    expect(fnXml).toContain("Footnote with Bob");
  });

  it("supports regex query with capture-based replacement", () => {
    const doc = Docx.create({ paragraphs: ["body {{a}}"] });
    doc.addHeader("Header {{b}}");
    const values: Record<string, string> = { a: "AAA", b: "BBB" };
    const count = doc.replaceTextEverywhere(
      /\{\{(\w+)\}\}/g,
      (m) => values[m.captures[0] ?? ""] ?? m.text,
    );
    expect(count).toBe(2);
    const reopened = Docx.open(doc.toUint8Array());
    expect(reopened.text).toContain("body AAA");
    const headerXml = new TextDecoder().decode(
      getPart(reopened.opc, "/word/header1.xml")?.data ?? new Uint8Array(),
    );
    expect(headerXml).toContain("Header BBB");
  });
});
