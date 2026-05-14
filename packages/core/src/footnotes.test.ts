import { hasPart, partRelationships, relationshipsByType } from "@word-kit/opc";
import { describe, expect, it } from "vitest";
import {
  addEndnote,
  addFootnote,
  createDocx,
  endnotesPart,
  footnotesPart,
  openDocx,
  paragraphs,
  toUint8Array,
} from "./docx.js";

describe("Docx.addFootnote", () => {
  it("creates footnotes.xml with separator entries on first use", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    const para = paragraphs(doc)[0];
    if (!para) return;
    expect(footnotesPart(doc)).toBeUndefined();
    const id = addFootnote(doc, para, "This is a footnote.");
    expect(id).toBe(1);
    const part = footnotesPart(doc);
    expect(part).toBeDefined();
    if (!part) return;
    // Standard separator (-1) + continuationSeparator (0) + the new footnote (1)
    expect(part.footnotes.length).toBe(3);
  });

  it("appends a footnoteReference run to the target paragraph", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    const para = paragraphs(doc)[0];
    if (!para) return;
    addFootnote(doc, para, "Note.");
    const lastChild = para.children.at(-1);
    expect(lastChild?.kind).toBe("raw");
    if (lastChild?.kind !== "raw") return;
    expect(lastChild.node.name.local).toBe("r");
    // Inside the run, look for footnoteReference
    const hasRef = lastChild.node.children.some(
      (c) => c.kind === "element" && c.name.local === "footnoteReference",
    );
    expect(hasRef).toBe(true);
  });

  it("survives save+reopen", () => {
    const doc = createDocx({ paragraphs: ["a"] });
    const para = paragraphs(doc)[0];
    if (!para) return;
    addFootnote(doc, para, "Note A");
    const reopened = openDocx(toUint8Array(doc));
    expect(footnotesPart(reopened)).toBeDefined();
    expect(footnotesPart(reopened)?.footnotes.length).toBeGreaterThanOrEqual(3);
    expect(hasPart(reopened.opc, "/word/footnotes.xml")).toBe(true);
  });

  it("registers footnotes relationship from document.xml", () => {
    const doc = createDocx({ paragraphs: ["a"] });
    const para = paragraphs(doc)[0];
    if (!para) return;
    addFootnote(doc, para, "Note");
    const rels = partRelationships(doc.opc, "/word/document.xml");
    const fr = relationshipsByType(
      rels,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
    );
    expect(fr).toHaveLength(1);
  });

  it("incremental ids when many footnotes are added", () => {
    const doc = createDocx({ paragraphs: ["a", "b"] });
    const [p1, p2] = paragraphs(doc);
    if (!p1 || !p2) return;
    expect(addFootnote(doc, p1, "one")).toBe(1);
    expect(addFootnote(doc, p2, "two")).toBe(2);
  });
});

describe("Docx.addEndnote", () => {
  it("creates endnotes.xml on first use and references from paragraph", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    const para = paragraphs(doc)[0];
    if (!para) return;
    addEndnote(doc, para, "End note");
    expect(endnotesPart(doc)).toBeDefined();
    const rels = partRelationships(doc.opc, "/word/document.xml");
    expect(
      relationshipsByType(
        rels,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
      ),
    ).toHaveLength(1);
  });
});
