import { hasPart, partRelationships, relationshipsByType } from "@word-kit/opc";
import { describe, expect, it } from "vitest";
import { Docx } from "./docx.js";

describe("Docx.addFootnote", () => {
  it("creates footnotes.xml with separator entries on first use", () => {
    const doc = Docx.create({ paragraphs: ["body"] });
    const para = doc.paragraphs[0];
    if (!para) return;
    expect(doc.footnotesPart).toBeUndefined();
    const id = doc.addFootnote(para, "This is a footnote.");
    expect(id).toBe(1);
    const part = doc.footnotesPart;
    expect(part).toBeDefined();
    if (!part) return;
    // Standard separator (-1) + continuationSeparator (0) + the new footnote (1)
    expect(part.footnotes.length).toBe(3);
  });

  it("appends a footnoteReference run to the target paragraph", () => {
    const doc = Docx.create({ paragraphs: ["body"] });
    const para = doc.paragraphs[0];
    if (!para) return;
    doc.addFootnote(para, "Note.");
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
    const doc = Docx.create({ paragraphs: ["a"] });
    const para = doc.paragraphs[0];
    if (!para) return;
    doc.addFootnote(para, "Note A");
    const reopened = Docx.open(doc.toUint8Array());
    expect(reopened.footnotesPart).toBeDefined();
    expect(reopened.footnotesPart?.footnotes.length).toBeGreaterThanOrEqual(3);
    expect(hasPart(reopened.opc, "/word/footnotes.xml")).toBe(true);
  });

  it("registers footnotes relationship from document.xml", () => {
    const doc = Docx.create({ paragraphs: ["a"] });
    const para = doc.paragraphs[0];
    if (!para) return;
    doc.addFootnote(para, "Note");
    const rels = partRelationships(doc.opc, "/word/document.xml");
    const fr = relationshipsByType(
      rels,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
    );
    expect(fr).toHaveLength(1);
  });

  it("incremental ids when many footnotes are added", () => {
    const doc = Docx.create({ paragraphs: ["a", "b"] });
    const [p1, p2] = doc.paragraphs;
    if (!p1 || !p2) return;
    expect(doc.addFootnote(p1, "one")).toBe(1);
    expect(doc.addFootnote(p2, "two")).toBe(2);
  });
});

describe("Docx.addEndnote", () => {
  it("creates endnotes.xml on first use and references from paragraph", () => {
    const doc = Docx.create({ paragraphs: ["body"] });
    const para = doc.paragraphs[0];
    if (!para) return;
    doc.addEndnote(para, "End note");
    expect(doc.endnotesPart).toBeDefined();
    const rels = partRelationships(doc.opc, "/word/document.xml");
    expect(
      relationshipsByType(
        rels,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
      ),
    ).toHaveLength(1);
  });
});
