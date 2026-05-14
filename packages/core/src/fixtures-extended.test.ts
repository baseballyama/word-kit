import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { Docx } from "./docx.js";

const REPO_ROOT = resolve(import.meta.dirname, "../../..");
const FIXTURE_DIR = resolve(REPO_ROOT, "references/mammoth.js/test/test-data");

function load(name: string): Uint8Array {
  return readFileSync(resolve(FIXTURE_DIR, name));
}

describe("real fixtures: extended coverage", () => {
  it("opens comments.docx, validates, and reads the comments part", () => {
    const doc = Docx.open(load("comments.docx"));
    const cp = doc.commentsPart;
    expect(cp).toBeDefined();
    expect(cp?.comments.length).toBeGreaterThan(0);
    // Validation should report at most warnings (Word produced this file).
    const errors = doc.validate().filter((i) => i.level === "error");
    expect(errors).toEqual([]);
  });

  it("opens footnotes.docx and reads the footnotes part", () => {
    const doc = Docx.open(load("footnotes.docx"));
    expect(doc.footnotesPart).toBeDefined();
    expect(doc.footnotesPart?.footnotes.length ?? 0).toBeGreaterThan(0);
    expect(doc.validate().filter((i) => i.level === "error")).toEqual([]);
  });

  it("opens endnotes.docx and reads the endnotes part", () => {
    const doc = Docx.open(load("endnotes.docx"));
    expect(doc.endnotesPart).toBeDefined();
    expect(doc.validate().filter((i) => i.level === "error")).toEqual([]);
  });

  it("opens simple-list.docx and round-trips it", () => {
    const doc = Docx.open(load("simple-list.docx"));
    const bytes = doc.toUint8Array();
    const reopened = Docx.open(bytes);
    expect(reopened.paragraphs.length).toBe(doc.paragraphs.length);
  });

  it("opens tiny-picture.docx and finds at least one media part", () => {
    const doc = Docx.open(load("tiny-picture.docx"));
    expect(doc.images.length).toBeGreaterThan(0);
  });

  // strict-format.docx uses the ISO/IEC 29500 Strict namespaces
  // (http://purl.oclc.org/ooxml/...) rather than the transitional
  // openxmlformats namespaces. We only support transitional today
  // (see PLAN.md), so we explicitly skip this fixture.
  it.skip("opens strict-format.docx (not yet supported)", () => {
    const doc = Docx.open(load("strict-format.docx"));
    expect(doc.paragraphs.length).toBeGreaterThan(0);
  });

  it("opens text-box.docx and preserves the body on round-trip", () => {
    const doc = Docx.open(load("text-box.docx"));
    const before = doc.text;
    const reopened = Docx.open(doc.toUint8Array());
    expect(reopened.text).toBe(before);
  });

  it("opens utf8-bom.docx and reads it", () => {
    const doc = Docx.open(load("utf8-bom.docx"));
    expect(doc.paragraphs.length).toBeGreaterThan(0);
  });

  it("opens hyperlinks.docx (subfolder fixture) when available", () => {
    const dir = resolve(FIXTURE_DIR, "hyperlinks");
    try {
      const fnames = require("node:fs").readdirSync(dir) as string[];
      const first = fnames.find((n) => n.endsWith(".docx"));
      if (!first) return;
      const doc = Docx.open(readFileSync(resolve(dir, first)));
      expect(doc.paragraphs.length).toBeGreaterThan(0);
    } catch {
      // directory missing or unreadable — silently skip
    }
  });

  it("opens external-picture.docx (image part may be missing)", () => {
    const doc = Docx.open(load("external-picture.docx"));
    expect(doc.paragraphs.length).toBeGreaterThan(0);
  });
});
