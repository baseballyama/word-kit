import { hasPart, packageRelationships, relationshipsByType } from "@word-kit/opc";
import { describe, expect, it } from "vitest";
import { appProperties, createDocx, openDocx, setAppProperties, toUint8Array } from "./docx.js";

describe("Docx app properties", () => {
  it("returns an empty record when no app.xml is present", () => {
    const doc = createDocx();
    expect(appProperties(doc)).toEqual({});
  });

  it("setAppProperties creates docProps/app.xml + rel on first use", () => {
    const doc = createDocx();
    setAppProperties(doc, {
      application: "word-kit",
      appVersion: "0.1.0",
      pages: 5,
      words: 1200,
      characters: 7800,
      company: "Acme",
    });
    expect(hasPart(doc.opc, "/docProps/app.xml")).toBe(true);
    const rels = relationshipsByType(
      packageRelationships(doc.opc),
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
    );
    expect(rels).toHaveLength(1);
  });

  it("round-trips through save+reopen", () => {
    const doc = createDocx();
    setAppProperties(doc, {
      application: "word-kit",
      appVersion: "0.1.0",
      words: 42,
      manager: "Yamada",
    });
    const reopened = openDocx(toUint8Array(doc));
    expect(appProperties(reopened)).toMatchObject({
      application: "word-kit",
      appVersion: "0.1.0",
      words: 42,
      manager: "Yamada",
    });
  });

  it("merges new properties with existing ones", () => {
    const doc = createDocx();
    setAppProperties(doc, { application: "first" });
    setAppProperties(doc, { pages: 3 });
    expect(appProperties(doc)).toMatchObject({ application: "first", pages: 3 });
  });
});
