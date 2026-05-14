import { describe, expect, it } from "vitest";
import { buildMinimalDocx } from "./minimal-docx.js";
import { OpcPackage } from "./package.js";

describe("OpcPackage (minimal docx)", () => {
  it("builds a package with all expected parts", () => {
    const pkg = buildMinimalDocx();
    expect(pkg.hasPart("/word/document.xml")).toBe(true);
    const document = pkg.getPart("/word/document.xml");
    expect(document?.contentType).toContain("wordprocessingml.document.main");
  });

  it("round-trips through write+read", () => {
    const pkg = buildMinimalDocx();
    const bytes = pkg.write();
    expect(bytes.length).toBeGreaterThan(0);
    const reloaded = OpcPackage.read(bytes);

    expect(reloaded.hasPart("/word/document.xml")).toBe(true);
    const document = reloaded.getPart("/word/document.xml");
    expect(document).toBeDefined();
    const text = new TextDecoder().decode(document?.data ?? new Uint8Array());
    expect(text).toContain("<w:body>");

    const rels = reloaded.packageRelationships.all;
    expect(rels).toHaveLength(1);
    expect(rels[0]?.target).toBe("word/document.xml");
  });

  it("preserves added parts and re-emits them on a second round-trip", () => {
    const pkg = buildMinimalDocx();
    pkg.addPart({
      name: "/word/styles.xml",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
      data: new TextEncoder().encode(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
      ),
    });
    pkg.packageRelationships.add({
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
      target: "word/styles.xml",
    });

    const bytes = pkg.write();
    const reloaded = OpcPackage.read(bytes);
    expect(reloaded.hasPart("/word/styles.xml")).toBe(true);
    expect(reloaded.packageRelationships.all).toHaveLength(2);
    expect(
      reloaded.contentTypes.resolve("/word/styles.xml") ??
        reloaded.contentTypes.resolve("/word/foo.xml"),
    ).toContain("styles+xml");
  });

  it("listParts() preserves insertion order", () => {
    const pkg = buildMinimalDocx();
    const namesBefore = pkg.listParts().map((p) => p.name);
    expect(namesBefore).toContain("/word/document.xml");

    const bytes = pkg.write();
    const reloaded = OpcPackage.read(bytes);
    const namesAfter = reloaded.listParts().map((p) => p.name);
    expect(namesAfter[0]).toBe("/[Content_Types].xml");
    expect(namesAfter).toContain("/word/document.xml");
  });

  it("rejects re-adding an existing part", () => {
    const pkg = buildMinimalDocx();
    expect(() =>
      pkg.addPart({
        name: "/word/document.xml",
        contentType: "application/xml",
        data: new Uint8Array(),
      }),
    ).toThrow(/already exists/);
  });

  it("rejects manipulating the content types part directly", () => {
    const pkg = buildMinimalDocx();
    expect(() =>
      pkg.addPart({
        name: "/[Content_Types].xml",
        contentType: "x",
        data: new Uint8Array(),
      }),
    ).toThrow();
    expect(() => pkg.removePart("/[Content_Types].xml")).toThrow();
  });
});
