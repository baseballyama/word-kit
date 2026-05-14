import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { Docx } from "./docx.js";

const REPO_ROOT = resolve(import.meta.dirname, "../../..");
const FIXTURE_DIR = resolve(REPO_ROOT, "references/mammoth.js/test/test-data");

function loadFixture(name: string): Uint8Array {
  return readFileSync(resolve(FIXTURE_DIR, name));
}

describe("real-world docx fixtures (mammoth.js test-data)", () => {
  it("opens single-paragraph.docx and reads its visible text", () => {
    const doc = Docx.open(loadFixture("single-paragraph.docx"));
    expect(doc.paragraphs.length).toBeGreaterThan(0);
    expect(doc.text).toContain("Walking on imported air");
  });

  it("opens empty.docx without throwing and reads its body", () => {
    const doc = Docx.open(loadFixture("empty.docx"));
    expect(doc.paragraphs.length).toBeGreaterThan(0);
  });

  it("opens tables.docx and recognizes the table structure", () => {
    const doc = Docx.open(loadFixture("tables.docx"));
    expect(doc.tables.length).toBeGreaterThan(0);
    const first = doc.tables[0];
    expect(first?.rows.length).toBeGreaterThan(0);
  });

  it("round-trips single-paragraph.docx through write+read", () => {
    const original = Docx.open(loadFixture("single-paragraph.docx"));
    const textBefore = original.text;
    const bytes = original.toUint8Array();
    const reopened = Docx.open(bytes);
    expect(reopened.text).toBe(textBefore);
  });

  it("round-trips tables.docx (preserves table structure and content)", () => {
    const original = Docx.open(loadFixture("tables.docx"));
    const rowsBefore = original.tables[0]?.rows.length ?? 0;
    const cellsBefore = original.tables[0]?.rows[0]?.cells.length ?? 0;
    const bytes = original.toUint8Array();
    const reopened = Docx.open(bytes);
    expect(reopened.tables[0]?.rows.length).toBe(rowsBefore);
    expect(reopened.tables[0]?.rows[0]?.cells.length).toBe(cellsBefore);
  });

  it("replaces text in a real Word-style docx and survives save+reopen", () => {
    const doc = Docx.open(loadFixture("single-paragraph.docx"));
    // The fixture's body is just "Walking on imported air" — exercise the replace path.
    expect(doc.replaceText("Walking on imported air", "Greetings.")).toBe(1);
    const reopened = Docx.open(doc.toUint8Array());
    expect(reopened.text).toContain("Greetings.");
  });
});
