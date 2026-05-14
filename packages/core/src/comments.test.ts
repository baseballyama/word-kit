import { hasPart, partRelationships, relationshipsByType } from "@word-kit/opc";
import { describe, expect, it } from "vitest";
import {
  addComment,
  commentsPart,
  createDocx,
  openDocx,
  paragraphs,
  toUint8Array,
} from "./docx.js";

describe("Docx.addComment", () => {
  it("creates comments.xml + comment entry and wraps the target paragraph", () => {
    const doc = createDocx({ paragraphs: ["Hello world"] });
    const para = paragraphs(doc)[0];
    expect(para).toBeDefined();
    if (!para) return;
    const id = addComment(doc, para, {
      author: "Reviewer",
      initials: "R",
      text: "Please double-check this.",
    });
    expect(id).toBe(0);
    expect(hasPart(doc.opc, "/word/comments.xml")).toBe(true);
    const cp = commentsPart(doc);
    expect(cp).toBeDefined();
    expect(cp?.comments).toHaveLength(1);
  });

  it("inserts commentRangeStart/End and commentReference around the paragraph", () => {
    const doc = createDocx({ paragraphs: ["Hello world"] });
    const para = paragraphs(doc)[0];
    if (!para) return;
    addComment(doc, para, { author: "R", text: "x" });
    const localNames = para.children.map((c) =>
      c.kind === "raw" ? c.node.name.local : c.kind === "run" ? "r" : c.kind,
    );
    expect(localNames[0]).toBe("commentRangeStart");
    expect(localNames).toContain("commentRangeEnd");
    expect(localNames[localNames.length - 1]).toBe("r"); // reference run last
  });

  it("registers a comments relationship from word/document.xml", () => {
    const doc = createDocx({ paragraphs: ["Hello"] });
    const para = paragraphs(doc)[0];
    if (!para) return;
    addComment(doc, para, { author: "R", text: "x" });
    const rels = partRelationships(doc.opc, "/word/document.xml");
    const commentRels = relationshipsByType(
      rels,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
    );
    expect(commentRels).toHaveLength(1);
  });

  it("save+reopen preserves the comment", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    const para = paragraphs(doc)[0];
    if (!para) return;
    addComment(doc, para, { author: "Reviewer", text: "Looks good." });
    const bytes = toUint8Array(doc);
    const reopened = openDocx(bytes);
    expect(commentsPart(reopened)).toBeDefined();
    expect(commentsPart(reopened)?.comments).toHaveLength(1);
  });

  it("assigns incremental ids when multiple comments are added", () => {
    const doc = createDocx({ paragraphs: ["a", "b"] });
    const [p1, p2] = paragraphs(doc);
    if (!p1 || !p2) return;
    expect(addComment(doc, p1, { author: "R", text: "1" })).toBe(0);
    expect(addComment(doc, p2, { author: "R", text: "2" })).toBe(1);
  });
});
