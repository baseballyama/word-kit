import { readdirSync, readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { Docx } from "./docx.js";

const REPO_ROOT = resolve(import.meta.dirname, "../../..");
const FIXTURE_DIR = resolve(REPO_ROOT, "references/python-docx/tests/test_files");

describe("python-docx fixture corpus: open + round-trip parity", () => {
  const all = readdirSync(FIXTURE_DIR).filter((n) => n.endsWith(".docx"));

  for (const name of all) {
    it(`opens and re-saves ${name} losslessly at the paragraph level`, () => {
      const bytes = readFileSync(resolve(FIXTURE_DIR, name));
      const doc = Docx.open(bytes);
      const paragraphsBefore = doc.paragraphs.length;
      const textBefore = doc.text;
      const reopened = Docx.open(doc.toUint8Array());
      expect(reopened.paragraphs.length).toBe(paragraphsBefore);
      expect(reopened.text).toBe(textBefore);
    });
  }

  it("test.docx has at least one paragraph", () => {
    const bytes = readFileSync(resolve(FIXTURE_DIR, "test.docx"));
    const doc = Docx.open(bytes);
    expect(doc.paragraphs.length).toBeGreaterThan(0);
  });

  it("having-images.docx exposes its media parts", () => {
    const bytes = readFileSync(resolve(FIXTURE_DIR, "having-images.docx"));
    const doc = Docx.open(bytes);
    expect(doc.images.length).toBeGreaterThan(0);
  });
});
