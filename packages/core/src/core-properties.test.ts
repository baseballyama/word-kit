import { describe, expect, it } from "vitest";
import { Docx } from "./docx.js";

describe("Docx core properties", () => {
  it("returns an empty record when the package has no docProps/core.xml", () => {
    const doc = Docx.create();
    expect(doc.coreProperties).toEqual({});
  });

  it("setCoreProperties creates the part on first use", () => {
    const doc = Docx.create();
    doc.setCoreProperties({
      title: "Quarterly Report",
      creator: "Yamada",
      subject: "Sales",
      description: "Q3 2026 numbers",
      keywords: "sales, q3, 2026",
      lastModifiedBy: "Reviewer",
      created: "2026-05-14T00:00:00Z",
      modified: "2026-05-15T00:00:00Z",
    });
    expect(doc.opc.hasPart("/docProps/core.xml")).toBe(true);
    const pkgRels = doc.opc.packageRelationships.byType(
      "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
    );
    expect(pkgRels).toHaveLength(1);
  });

  it("round-trips properties through save+reopen", () => {
    const doc = Docx.create();
    doc.setCoreProperties({ title: "T", creator: "C", description: "D" });
    const reopened = Docx.open(doc.toUint8Array());
    expect(reopened.coreProperties).toMatchObject({
      title: "T",
      creator: "C",
      description: "D",
    });
  });

  it("merges new properties with existing ones", () => {
    const doc = Docx.create();
    doc.setCoreProperties({ title: "First" });
    doc.setCoreProperties({ creator: "Author" });
    expect(doc.coreProperties).toMatchObject({ title: "First", creator: "Author" });
  });
});
