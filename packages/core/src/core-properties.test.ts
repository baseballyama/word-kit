import { hasPart, packageRelationships, relationshipsByType } from "@word-kit/opc";
import { describe, expect, it } from "vitest";
import { coreProperties, createDocx, openDocx, setCoreProperties, toUint8Array } from "./docx.js";

describe("Docx core properties", () => {
  it("returns an empty record when the package has no docProps/core.xml", () => {
    const doc = createDocx();
    expect(coreProperties(doc)).toEqual({});
  });

  it("setCoreProperties creates the part on first use", () => {
    const doc = createDocx();
    setCoreProperties(doc, {
      title: "Quarterly Report",
      creator: "Yamada",
      subject: "Sales",
      description: "Q3 2026 numbers",
      keywords: "sales, q3, 2026",
      lastModifiedBy: "Reviewer",
      created: "2026-05-14T00:00:00Z",
      modified: "2026-05-15T00:00:00Z",
    });
    expect(hasPart(doc.opc, "/docProps/core.xml")).toBe(true);
    const pkgRels = relationshipsByType(
      packageRelationships(doc.opc),
      "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
    );
    expect(pkgRels).toHaveLength(1);
  });

  it("round-trips properties through save+reopen", () => {
    const doc = createDocx();
    setCoreProperties(doc, { title: "T", creator: "C", description: "D" });
    const reopened = openDocx(toUint8Array(doc));
    expect(coreProperties(reopened)).toMatchObject({
      title: "T",
      creator: "C",
      description: "D",
    });
  });

  it("merges new properties with existing ones", () => {
    const doc = createDocx();
    setCoreProperties(doc, { title: "First" });
    setCoreProperties(doc, { creator: "Author" });
    expect(coreProperties(doc)).toMatchObject({ title: "First", creator: "Author" });
  });
});
