import { describe, expect, it } from "vitest";
import { ContentTypesIndex } from "./content-types.js";

const SAMPLE_XML = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
  '<Default Extension="xml" ContentType="application/xml"/>',
  '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
  "</Types>",
].join("");

describe("ContentTypesIndex", () => {
  it("parses Default and Override entries in order", () => {
    const idx = ContentTypesIndex.fromXml(SAMPLE_XML);
    expect(idx.allDefaults.map((d) => d.extension)).toEqual(["rels", "xml"]);
    expect(idx.allOverrides.map((o) => o.partName)).toEqual(["/word/document.xml"]);
  });

  it("resolves Overrides before Defaults", () => {
    const idx = ContentTypesIndex.fromXml(SAMPLE_XML);
    expect(idx.resolve("/word/document.xml")).toContain("wordprocessingml.document.main+xml");
    expect(idx.resolve("/foo.xml")).toBe("application/xml");
    expect(idx.resolve("/no-such-extension.unknown")).toBeUndefined();
  });

  it("treats Override part names case-insensitively", () => {
    const idx = ContentTypesIndex.fromXml(SAMPLE_XML);
    expect(idx.resolve("/WORD/DOCUMENT.XML")).toContain("wordprocessingml");
  });

  it("round-trips through XML", () => {
    const idx = ContentTypesIndex.fromXml(SAMPLE_XML);
    const out = idx.toXml();
    const idx2 = ContentTypesIndex.fromXml(out);
    expect(idx2.allDefaults).toEqual(idx.allDefaults);
    expect(idx2.allOverrides).toEqual(idx.allOverrides);
  });

  it("supports setDefault / setOverride / remove*", () => {
    const idx = ContentTypesIndex.empty();
    idx.setDefault("xml", "application/xml");
    idx.setOverride("/word/document.xml", "application/x-test+xml");
    expect(idx.resolve("/foo.xml")).toBe("application/xml");
    expect(idx.resolve("/word/document.xml")).toBe("application/x-test+xml");
    expect(idx.removeOverride("/word/document.xml")).toBe(true);
    expect(idx.resolve("/word/document.xml")).toBe("application/xml");
    expect(idx.removeDefault("xml")).toBe(true);
    expect(idx.resolve("/foo.xml")).toBeUndefined();
  });
});
