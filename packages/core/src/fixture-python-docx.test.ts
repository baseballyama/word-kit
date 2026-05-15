import { existsSync, readdirSync, readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { images, openDocx, paragraphs, text, toUint8Array } from "./docx.js";

const REPO_ROOT = resolve(import.meta.dirname, "../../..");
const FIXTURE_DIR = resolve(REPO_ROOT, "references/python-docx/tests/test_files");

describe("python-docx fixture corpus: open + round-trip parity", () => {
  if (!existsSync(FIXTURE_DIR)) {
    it.skip("references/python-docx submodule not initialised — run `git submodule update --init --recursive`", () => {
      // intentionally empty
    });
    return;
  }

  const all = readdirSync(FIXTURE_DIR).filter((n) => n.endsWith(".docx"));

  for (const name of all) {
    it(`opens and re-saves ${name} losslessly at the paragraph level`, () => {
      const bytes = readFileSync(resolve(FIXTURE_DIR, name));
      const doc = openDocx(bytes);
      const paragraphsBefore = paragraphs(doc).length;
      const textBefore = text(doc);
      const reopened = openDocx(toUint8Array(doc));
      expect(paragraphs(reopened).length).toBe(paragraphsBefore);
      expect(text(reopened)).toBe(textBefore);
    });
  }

  it("test.docx has at least one paragraph", () => {
    const bytes = readFileSync(resolve(FIXTURE_DIR, "test.docx"));
    const doc = openDocx(bytes);
    expect(paragraphs(doc).length).toBeGreaterThan(0);
  });

  it("having-images.docx exposes its media parts", () => {
    const bytes = readFileSync(resolve(FIXTURE_DIR, "having-images.docx"));
    const doc = openDocx(bytes);
    expect(images(doc).length).toBeGreaterThan(0);
  });
});
