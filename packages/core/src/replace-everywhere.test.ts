import { getPart } from "@word-kit/opc";
import { describe, expect, it } from "vitest";
import {
  addComment,
  addFooter,
  addFootnote,
  addHeader,
  createDocx,
  openDocx,
  paragraphs,
  replaceTextEverywhere,
  text,
  toUint8Array,
} from "./docx.js";

describe("Docx.replaceTextEverywhere", () => {
  it("replaces text in headers and footers", () => {
    const doc = createDocx({ paragraphs: ["body has {{name}}"] });
    addHeader(doc, "Header: {{name}}");
    addFooter(doc, "Footer: {{name}}");
    const count = replaceTextEverywhere(doc, "{{name}}", "山田太郎");
    expect(count).toBe(3); // body + header + footer

    const bytes = toUint8Array(doc);
    const reopened = openDocx(bytes);
    expect(text(reopened)).toContain("body has 山田太郎");
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
    const doc = createDocx({ paragraphs: ["body"] });
    const para = paragraphs(doc)[0];
    if (!para) return;
    addComment(doc, para, { author: "R", text: "Please update {{name}}." });
    const count = replaceTextEverywhere(doc, "{{name}}", "Alice");
    expect(count).toBeGreaterThan(0);
    const bytes = toUint8Array(doc);
    const reopened = openDocx(bytes);
    const commentsXml = new TextDecoder().decode(
      getPart(reopened.opc, "/word/comments.xml")?.data ?? new Uint8Array(),
    );
    expect(commentsXml).toContain("Please update Alice.");
  });

  it("replaces text inside footnotes too", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    const para = paragraphs(doc)[0];
    if (!para) return;
    addFootnote(doc, para, "Footnote with {{name}}");
    replaceTextEverywhere(doc, "{{name}}", "Bob");
    const reopened = openDocx(toUint8Array(doc));
    const fnXml = new TextDecoder().decode(
      getPart(reopened.opc, "/word/footnotes.xml")?.data ?? new Uint8Array(),
    );
    expect(fnXml).toContain("Footnote with Bob");
  });

  it("supports regex query with capture-based replacement", () => {
    const doc = createDocx({ paragraphs: ["body {{a}}"] });
    addHeader(doc, "Header {{b}}");
    const values: Record<string, string> = { a: "AAA", b: "BBB" };
    const count = replaceTextEverywhere(
      doc,
      /\{\{(\w+)\}\}/g,
      (m) => values[m.captures[0] ?? ""] ?? m.text,
    );
    expect(count).toBe(2);
    const reopened = openDocx(toUint8Array(doc));
    expect(text(reopened)).toContain("body AAA");
    const headerXml = new TextDecoder().decode(
      getPart(reopened.opc, "/word/header1.xml")?.data ?? new Uint8Array(),
    );
    expect(headerXml).toContain("Header BBB");
  });
});
