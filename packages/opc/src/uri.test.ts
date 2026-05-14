import { describe, expect, it } from "vitest";
import {
  normalizePartName,
  partBaseName,
  partFolder,
  partNameToZipEntry,
  relativizeTarget,
  relsPartNameFor,
  resolveInternalTarget,
  zipEntryToPartName,
} from "./uri.js";

describe("normalizePartName", () => {
  it("adds a leading slash if missing", () => {
    expect(normalizePartName("word/document.xml")).toBe("/word/document.xml");
  });

  it("collapses redundant slashes", () => {
    expect(normalizePartName("//word//document.xml")).toBe("/word/document.xml");
  });

  it("resolves dot segments", () => {
    expect(normalizePartName("/word/./media/../document.xml")).toBe("/word/document.xml");
  });

  it("rejects escapes above the root", () => {
    expect(() => normalizePartName("/word/../../etc/passwd")).toThrow();
  });

  it("rejects a name that resolves to root", () => {
    expect(() => normalizePartName("/")).toThrow();
  });
});

describe("partNameToZipEntry / zipEntryToPartName", () => {
  it("round-trips", () => {
    expect(partNameToZipEntry("/word/document.xml")).toBe("word/document.xml");
    expect(zipEntryToPartName("word/document.xml")).toBe("/word/document.xml");
  });
});

describe("partFolder / partBaseName", () => {
  it("handles nested paths", () => {
    expect(partFolder("/word/_rels/document.xml.rels")).toBe("/word/_rels/");
    expect(partBaseName("/word/_rels/document.xml.rels")).toBe("document.xml.rels");
  });
  it("handles root-level parts", () => {
    expect(partFolder("/[Content_Types].xml")).toBe("/");
    expect(partBaseName("/[Content_Types].xml")).toBe("[Content_Types].xml");
  });
});

describe("relsPartNameFor", () => {
  it("returns the package rels path for 'package'", () => {
    expect(relsPartNameFor("package")).toBe("/_rels/.rels");
  });
  it("returns the part rels path for a part", () => {
    expect(relsPartNameFor("/word/document.xml")).toBe("/word/_rels/document.xml.rels");
  });
});

describe("resolveInternalTarget", () => {
  it("resolves a relative target against the source folder", () => {
    expect(resolveInternalTarget("/word/document.xml", "media/image1.png")).toBe(
      "/word/media/image1.png",
    );
  });
  it("treats an absolute target as-is", () => {
    expect(resolveInternalTarget("/word/document.xml", "/word/media/image1.png")).toBe(
      "/word/media/image1.png",
    );
  });
  it("handles ../ in the target", () => {
    expect(resolveInternalTarget("/word/document.xml", "../docProps/core.xml")).toBe(
      "/docProps/core.xml",
    );
  });
});

describe("relativizeTarget", () => {
  it("returns sibling-style relative paths when possible", () => {
    expect(relativizeTarget("/word/document.xml", "/word/styles.xml")).toBe("styles.xml");
    expect(relativizeTarget("/word/document.xml", "/word/media/image1.png")).toBe(
      "media/image1.png",
    );
  });
  it("returns ../-prefixed paths when going up", () => {
    expect(relativizeTarget("/word/document.xml", "/docProps/core.xml")).toBe(
      "../docProps/core.xml",
    );
  });
});
