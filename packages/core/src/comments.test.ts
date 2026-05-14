import { hasPart, partRelationships, relationshipsByType } from "@word-kit/opc";
import { describe, expect, it } from "vitest";
import { Docx } from "./docx.js";

describe("Docx.addComment", () => {
  it("creates comments.xml + comment entry and wraps the target paragraph", () => {
    const doc = Docx.create({ paragraphs: ["Hello world"] });
    const para = doc.paragraphs[0];
    expect(para).toBeDefined();
    if (!para) return;
    const id = doc.addComment(para, {
      author: "Reviewer",
      initials: "R",
      text: "Please double-check this.",
    });
    expect(id).toBe(0);
    expect(hasPart(doc.opc, "/word/comments.xml")).toBe(true);
    const cp = doc.commentsPart;
    expect(cp).toBeDefined();
    expect(cp?.comments).toHaveLength(1);
  });

  it("inserts commentRangeStart/End and commentReference around the paragraph", () => {
    const doc = Docx.create({ paragraphs: ["Hello world"] });
    const para = doc.paragraphs[0];
    if (!para) return;
    doc.addComment(para, { author: "R", text: "x" });
    const localNames = para.children.map((c) =>
      c.kind === "raw" ? c.node.name.local : c.kind === "run" ? "r" : c.kind,
    );
    expect(localNames[0]).toBe("commentRangeStart");
    expect(localNames).toContain("commentRangeEnd");
    expect(localNames[localNames.length - 1]).toBe("r"); // reference run last
  });

  it("registers a comments relationship from word/document.xml", () => {
    const doc = Docx.create({ paragraphs: ["Hello"] });
    const para = doc.paragraphs[0];
    if (!para) return;
    doc.addComment(para, { author: "R", text: "x" });
    const rels = partRelationships(doc.opc, "/word/document.xml");
    const commentRels = relationshipsByType(
      rels,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
    );
    expect(commentRels).toHaveLength(1);
  });

  it("save+reopen preserves the comment", () => {
    const doc = Docx.create({ paragraphs: ["body"] });
    const para = doc.paragraphs[0];
    if (!para) return;
    doc.addComment(para, { author: "Reviewer", text: "Looks good." });
    const bytes = doc.toUint8Array();
    const reopened = Docx.open(bytes);
    expect(reopened.commentsPart).toBeDefined();
    expect(reopened.commentsPart?.comments).toHaveLength(1);
  });

  it("assigns incremental ids when multiple comments are added", () => {
    const doc = Docx.create({ paragraphs: ["a", "b"] });
    const [p1, p2] = doc.paragraphs;
    if (!p1 || !p2) return;
    expect(doc.addComment(p1, { author: "R", text: "1" })).toBe(0);
    expect(doc.addComment(p2, { author: "R", text: "2" })).toBe(1);
  });
});
