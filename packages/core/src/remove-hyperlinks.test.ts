import { allRelationships, partRelationships } from "@word-kit/opc";
import { describe, expect, it } from "vitest";
import {
  addHyperlink,
  addInternalHyperlink,
  appendParagraph,
  createDocx,
  openDocx,
  removeAllHyperlinks,
  text,
  toUint8Array,
} from "./docx.js";

describe("removeAllHyperlinks", () => {
  it("unwraps external links: text remains, rel is dropped", () => {
    const doc = createDocx({ paragraphs: ["intro"] });
    addHyperlink(doc, "https://example.com/", "the report");
    appendParagraph(doc, "outro");

    // Sanity: text contains the link text before removal.
    expect(text(doc)).toContain("the report");

    const before = allRelationships(partRelationships(doc.opc, doc.partName)).filter((r) =>
      r.type.endsWith("/hyperlink"),
    ).length;
    expect(before).toBe(1);

    const unwrapped = removeAllHyperlinks(doc);
    expect(unwrapped).toBe(1);
    // Text content preserved.
    expect(text(doc)).toContain("the report");

    const after = allRelationships(partRelationships(doc.opc, doc.partName)).filter((r) =>
      r.type.endsWith("/hyperlink"),
    ).length;
    expect(after).toBe(0);
  });

  it("survives a save+reopen with link text intact and no hyperlink elements", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    addHyperlink(doc, "https://example.com/", "click me");
    removeAllHyperlinks(doc);

    const bytes = toUint8Array(doc);
    const reopened = openDocx(bytes);
    const xmlText = new TextDecoder().decode(
      reopened.opc.parts.get("/word/document.xml")?.data ?? new Uint8Array(),
    );
    expect(xmlText).not.toContain("<w:hyperlink");
    expect(xmlText).toContain("click me");
    // Hyperlink rels should not be in document.xml.rels either.
    const relsPart = reopened.opc.parts.get("/word/_rels/document.xml.rels");
    const relsText = new TextDecoder().decode(relsPart?.data ?? new Uint8Array());
    expect(relsText).not.toContain("https://example.com/");
  });

  it("unwraps internal anchor hyperlinks without affecting other rels", () => {
    const doc = createDocx({ paragraphs: ["target"] });
    addInternalHyperlink(doc, "ch1", "jump");

    const unwrapped = removeAllHyperlinks(doc);
    expect(unwrapped).toBe(1);
    // No external hyperlink relationships were created, so nothing to remove.
    const hyperRels = allRelationships(partRelationships(doc.opc, doc.partName)).filter((r) =>
      r.type.endsWith("/hyperlink"),
    );
    expect(hyperRels).toHaveLength(0);
    // The link text survives.
    expect(text(doc)).toContain("jump");
  });

  it("returns 0 when no hyperlinks are present", () => {
    const doc = createDocx({ paragraphs: ["plain"] });
    expect(removeAllHyperlinks(doc)).toBe(0);
  });
});
