import { describe, expect, it } from "vitest";
import { Docx } from "./docx.js";

describe("Docx app properties", () => {
  it("returns an empty record when no app.xml is present", () => {
    const doc = Docx.create();
    expect(doc.appProperties).toEqual({});
  });

  it("setAppProperties creates docProps/app.xml + rel on first use", () => {
    const doc = Docx.create();
    doc.setAppProperties({
      application: "word-kit",
      appVersion: "0.1.0",
      pages: 5,
      words: 1200,
      characters: 7800,
      company: "Acme",
    });
    expect(doc.opc.hasPart("/docProps/app.xml")).toBe(true);
    const rels = doc.opc.packageRelationships.byType(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
    );
    expect(rels).toHaveLength(1);
  });

  it("round-trips through save+reopen", () => {
    const doc = Docx.create();
    doc.setAppProperties({
      application: "word-kit",
      appVersion: "0.1.0",
      words: 42,
      manager: "Yamada",
    });
    const reopened = Docx.open(doc.toUint8Array());
    expect(reopened.appProperties).toMatchObject({
      application: "word-kit",
      appVersion: "0.1.0",
      words: 42,
      manager: "Yamada",
    });
  });

  it("merges new properties with existing ones", () => {
    const doc = Docx.create();
    doc.setAppProperties({ application: "first" });
    doc.setAppProperties({ pages: 3 });
    expect(doc.appProperties).toMatchObject({ application: "first", pages: 3 });
  });
});
