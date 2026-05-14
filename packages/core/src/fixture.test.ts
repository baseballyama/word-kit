import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { openDocx, paragraphs, replaceText, tables, text, toUint8Array } from "./docx.js";

const REPO_ROOT = resolve(import.meta.dirname, "../../..");
const FIXTURE_DIR = resolve(REPO_ROOT, "references/mammoth.js/test/test-data");

function loadFixture(name: string): Uint8Array {
  return readFileSync(resolve(FIXTURE_DIR, name));
}

describe("real-world docx fixtures (mammoth.js test-data)", () => {
  it("opens single-paragraph.docx and reads its visible text", () => {
    const doc = openDocx(loadFixture("single-paragraph.docx"));
    expect(paragraphs(doc).length).toBeGreaterThan(0);
    expect(text(doc)).toContain("Walking on imported air");
  });

  it("opens empty.docx without throwing and reads its body", () => {
    const doc = openDocx(loadFixture("empty.docx"));
    expect(paragraphs(doc).length).toBeGreaterThan(0);
  });

  it("opens tables.docx and recognizes the table structure", () => {
    const doc = openDocx(loadFixture("tables.docx"));
    expect(tables(doc).length).toBeGreaterThan(0);
    const first = tables(doc)[0];
    expect(first?.rows.length).toBeGreaterThan(0);
  });

  it("round-trips single-paragraph.docx through write+read", () => {
    const original = openDocx(loadFixture("single-paragraph.docx"));
    const textBefore = text(original);
    const bytes = toUint8Array(original);
    const reopened = openDocx(bytes);
    expect(text(reopened)).toBe(textBefore);
  });

  it("round-trips tables.docx (preserves table structure and content)", () => {
    const original = openDocx(loadFixture("tables.docx"));
    const rowsBefore = tables(original)[0]?.rows.length ?? 0;
    const cellsBefore = tables(original)[0]?.rows[0]?.cells.length ?? 0;
    const bytes = toUint8Array(original);
    const reopened = openDocx(bytes);
    expect(tables(reopened)[0]?.rows.length).toBe(rowsBefore);
    expect(tables(reopened)[0]?.rows[0]?.cells.length).toBe(cellsBefore);
  });

  it("replaces text in a real Word-style docx and survives save+reopen", () => {
    const doc = openDocx(loadFixture("single-paragraph.docx"));
    // The fixture's body is just "Walking on imported air" — exercise the replace path.
    expect(replaceText(doc, "Walking on imported air", "Greetings.")).toBe(1);
    const reopened = openDocx(toUint8Array(doc));
    expect(text(reopened)).toContain("Greetings.");
  });
});
