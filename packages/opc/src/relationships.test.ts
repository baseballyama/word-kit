import { describe, expect, it } from "vitest";
import {
  addRelationship,
  allRelationships,
  parseRelationshipsXml,
  relationshipById,
  relationshipsByType,
  removeRelationship,
  serializeRelationshipsXml,
} from "./relationships.js";

const SAMPLE_XML = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>',
  '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>',
  '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/>',
  "</Relationships>",
].join("");

describe("relationship set", () => {
  it("parses Id/Type/Target/TargetMode preserving order", () => {
    const set = parseRelationshipsXml(SAMPLE_XML);
    expect(allRelationships(set).map((r) => r.id)).toEqual(["rId1", "rId2", "rId3"]);
    expect(relationshipById(set, "rId3")?.targetMode).toBe("External");
    expect(relationshipById(set, "rId1")?.targetMode).toBe("Internal");
  });

  it("filters by type", () => {
    const set = parseRelationshipsXml(SAMPLE_XML);
    expect(
      relationshipsByType(
        set,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
      ),
    ).toHaveLength(1);
  });

  it("round-trips through XML", () => {
    const set = parseRelationshipsXml(SAMPLE_XML);
    const reparsed = parseRelationshipsXml(serializeRelationshipsXml(set));
    expect(allRelationships(reparsed)).toEqual(allRelationships(set));
  });

  it("allocates the next free rId on add", () => {
    const set = parseRelationshipsXml(SAMPLE_XML);
    const r = addRelationship(set, {
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
      target: "styles.xml",
    });
    expect(r.id).toBe("rId4");
    expect(r.targetMode).toBe("Internal");
  });

  it("rejects duplicate ids", () => {
    const set = parseRelationshipsXml(SAMPLE_XML);
    expect(() => addRelationship(set, { id: "rId1", type: "x", target: "y" })).toThrow(
      /already in use/,
    );
  });

  it("removeRelationship returns true and updates membership", () => {
    const set = parseRelationshipsXml(SAMPLE_XML);
    expect(removeRelationship(set, "rId2")).toBe(true);
    expect(relationshipById(set, "rId2")).toBeUndefined();
    expect(allRelationships(set)).toHaveLength(2);
  });
});
