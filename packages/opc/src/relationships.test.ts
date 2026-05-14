import { describe, expect, it } from "vitest";
import { RelationshipSet } from "./relationships.js";

const SAMPLE_XML = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>',
  '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>',
  '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/>',
  "</Relationships>",
].join("");

describe("RelationshipSet", () => {
  it("parses Id/Type/Target/TargetMode preserving order", () => {
    const set = RelationshipSet.fromXml(SAMPLE_XML);
    expect(set.all.map((r) => r.id)).toEqual(["rId1", "rId2", "rId3"]);
    expect(set.byId("rId3")?.targetMode).toBe("External");
    expect(set.byId("rId1")?.targetMode).toBe("Internal");
  });

  it("filters by type", () => {
    const set = RelationshipSet.fromXml(SAMPLE_XML);
    expect(
      set.byType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"),
    ).toHaveLength(1);
  });

  it("round-trips through XML", () => {
    const set = RelationshipSet.fromXml(SAMPLE_XML);
    const reparsed = RelationshipSet.fromXml(set.toXml());
    expect(reparsed.all).toEqual(set.all);
  });

  it("allocates the next free rId on add", () => {
    const set = RelationshipSet.fromXml(SAMPLE_XML);
    const r = set.add({
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
      target: "styles.xml",
    });
    expect(r.id).toBe("rId4");
    expect(r.targetMode).toBe("Internal");
  });

  it("rejects duplicate ids", () => {
    const set = RelationshipSet.fromXml(SAMPLE_XML);
    expect(() => set.add({ id: "rId1", type: "x", target: "y" })).toThrow(/already in use/);
  });

  it("remove() returns true and updates membership", () => {
    const set = RelationshipSet.fromXml(SAMPLE_XML);
    expect(set.remove("rId2")).toBe(true);
    expect(set.byId("rId2")).toBeUndefined();
    expect(set.all).toHaveLength(2);
  });
});
