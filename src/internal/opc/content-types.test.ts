import { describe, expect, it } from "vitest";
import {
  emptyContentTypes,
  parseContentTypesXml,
  removeContentTypeDefault,
  removeContentTypeOverride,
  resolveContentType,
  serializeContentTypesXml,
  setContentTypeDefault,
  setContentTypeOverride,
} from "./content-types.js";

const SAMPLE_XML = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
  '<Default Extension="xml" ContentType="application/xml"/>',
  '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
  "</Types>",
].join("");

describe("content types index", () => {
  it("parses Default and Override entries in order", () => {
    const idx = parseContentTypesXml(SAMPLE_XML);
    expect(idx.defaults.map((d) => d.extension)).toEqual(["rels", "xml"]);
    expect(idx.overrides.map((o) => o.partName)).toEqual(["/word/document.xml"]);
  });

  it("resolves Overrides before Defaults", () => {
    const idx = parseContentTypesXml(SAMPLE_XML);
    expect(resolveContentType(idx, "/word/document.xml")).toContain(
      "wordprocessingml.document.main+xml",
    );
    expect(resolveContentType(idx, "/foo.xml")).toBe("application/xml");
    expect(resolveContentType(idx, "/no-such-extension.unknown")).toBeUndefined();
  });

  it("treats Override part names case-insensitively", () => {
    const idx = parseContentTypesXml(SAMPLE_XML);
    expect(resolveContentType(idx, "/WORD/DOCUMENT.XML")).toContain("wordprocessingml");
  });

  it("round-trips through XML", () => {
    const idx = parseContentTypesXml(SAMPLE_XML);
    const idx2 = parseContentTypesXml(serializeContentTypesXml(idx));
    expect(idx2.defaults).toEqual(idx.defaults);
    expect(idx2.overrides).toEqual(idx.overrides);
  });

  it("supports setDefault / setOverride / remove*", () => {
    const idx = emptyContentTypes();
    setContentTypeDefault(idx, "xml", "application/xml");
    setContentTypeOverride(idx, "/word/document.xml", "application/x-test+xml");
    expect(resolveContentType(idx, "/foo.xml")).toBe("application/xml");
    expect(resolveContentType(idx, "/word/document.xml")).toBe("application/x-test+xml");
    expect(removeContentTypeOverride(idx, "/word/document.xml")).toBe(true);
    expect(resolveContentType(idx, "/word/document.xml")).toBe("application/xml");
    expect(removeContentTypeDefault(idx, "xml")).toBe(true);
    expect(resolveContentType(idx, "/foo.xml")).toBeUndefined();
  });
});
