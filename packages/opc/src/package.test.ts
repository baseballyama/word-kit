import { describe, expect, it } from "vitest";
import { resolveContentType } from "./content-types.js";
import { buildMinimalDocx } from "./minimal-docx.js";
import {
  addPart,
  getPart,
  hasPart,
  listParts,
  packageRelationships,
  readOpcPackage,
  removePart,
  writeOpcPackage,
} from "./package.js";
import { addRelationship, allRelationships } from "./relationships.js";

describe("OPC package (minimal docx)", () => {
  it("builds a package with all expected parts", () => {
    const pkg = buildMinimalDocx();
    expect(hasPart(pkg, "/word/document.xml")).toBe(true);
    const document = getPart(pkg, "/word/document.xml");
    expect(document?.contentType).toContain("wordprocessingml.document.main");
  });

  it("round-trips through write+read", () => {
    const pkg = buildMinimalDocx();
    const bytes = writeOpcPackage(pkg);
    expect(bytes.length).toBeGreaterThan(0);
    const reloaded = readOpcPackage(bytes);

    expect(hasPart(reloaded, "/word/document.xml")).toBe(true);
    const document = getPart(reloaded, "/word/document.xml");
    expect(document).toBeDefined();
    const text = new TextDecoder().decode(document?.data ?? new Uint8Array());
    expect(text).toContain("<w:body>");

    const rels = allRelationships(packageRelationships(reloaded));
    expect(rels).toHaveLength(1);
    expect(rels[0]?.target).toBe("word/document.xml");
  });

  it("preserves added parts and re-emits them on a second round-trip", () => {
    const pkg = buildMinimalDocx();
    addPart(pkg, {
      name: "/word/styles.xml",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
      data: new TextEncoder().encode(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
      ),
    });
    addRelationship(packageRelationships(pkg), {
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
      target: "word/styles.xml",
    });

    const bytes = writeOpcPackage(pkg);
    const reloaded = readOpcPackage(bytes);
    expect(hasPart(reloaded, "/word/styles.xml")).toBe(true);
    expect(allRelationships(packageRelationships(reloaded))).toHaveLength(2);
    expect(
      resolveContentType(reloaded.contentTypes, "/word/styles.xml") ??
        resolveContentType(reloaded.contentTypes, "/word/foo.xml"),
    ).toContain("styles+xml");
  });

  it("listParts() preserves insertion order", () => {
    const pkg = buildMinimalDocx();
    const namesBefore = listParts(pkg).map((p) => p.name);
    expect(namesBefore).toContain("/word/document.xml");

    const bytes = writeOpcPackage(pkg);
    const reloaded = readOpcPackage(bytes);
    const namesAfter = listParts(reloaded).map((p) => p.name);
    expect(namesAfter[0]).toBe("/[Content_Types].xml");
    expect(namesAfter).toContain("/word/document.xml");
  });

  it("rejects re-adding an existing part", () => {
    const pkg = buildMinimalDocx();
    expect(() =>
      addPart(pkg, {
        name: "/word/document.xml",
        contentType: "application/xml",
        data: new Uint8Array(),
      }),
    ).toThrow(/already exists/);
  });

  it("rejects manipulating the content types part directly", () => {
    const pkg = buildMinimalDocx();
    expect(() =>
      addPart(pkg, {
        name: "/[Content_Types].xml",
        contentType: "x",
        data: new Uint8Array(),
      }),
    ).toThrow();
    expect(() => removePart(pkg, "/[Content_Types].xml")).toThrow();
  });
});
