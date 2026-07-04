import { describe, expect, it } from "vitest";
import { parseMiniXml, serializeMiniXml } from "./mini-xml.js";

const CT_XML = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
  '<Default Extension="xml" ContentType="application/xml"/>',
  '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
  "</Types>",
].join("");

describe("parseMiniXml", () => {
  it("captures the root element name and its attributes", () => {
    const doc = parseMiniXml(CT_XML);
    expect(doc.root).toBe("Types");
    expect(doc.rootAttrs).toEqual([
      ["xmlns", "http://schemas.openxmlformats.org/package/2006/content-types"],
    ]);
  });

  it("captures children in order with their attributes", () => {
    const doc = parseMiniXml(CT_XML);
    expect(doc.children.map((c) => c.name)).toEqual(["Default", "Default", "Override"]);
    expect(doc.children[0]?.attrs).toEqual([
      ["Extension", "rels"],
      ["ContentType", "application/vnd.openxmlformats-package.relationships+xml"],
    ]);
    expect(doc.children[2]?.attrs[0]).toEqual(["PartName", "/word/document.xml"]);
  });

  it("captures the standalone declaration", () => {
    const doc = parseMiniXml(CT_XML);
    expect(doc.standalone).toBe("yes");
  });

  it("decodes the five standard entities and numeric references", () => {
    const doc = parseMiniXml(
      '<Root attr="a&amp;b&lt;c&gt;d&quot;e&apos;f&#48;&#x41;"><Child/></Root>',
    );
    expect(doc.rootAttrs[0]?.[1]).toBe("a&b<c>d\"e'f0A");
  });
});

describe("serializeMiniXml", () => {
  it("emits a well-formed declaration and root", () => {
    const xml = serializeMiniXml({
      root: "Types",
      rootAttrs: [["xmlns", "http://example/"]],
      children: [],
      standalone: "yes",
    });
    expect(xml.startsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')).toBe(true);
    expect(xml).toContain('<Types xmlns="http://example/"/>');
  });

  it("escapes attribute values", () => {
    const xml = serializeMiniXml({
      root: "Root",
      rootAttrs: [["attr", `a"b&c<d>e`]],
      children: [],
    });
    expect(xml).toContain('attr="a&quot;b&amp;c&lt;d&gt;e"');
  });

  it("round-trips a Content Types document", () => {
    const original = parseMiniXml(CT_XML);
    const reserialized = serializeMiniXml(original);
    const reparsed = parseMiniXml(reserialized);
    expect(reparsed.root).toBe(original.root);
    expect(reparsed.children.length).toBe(original.children.length);
    expect(reparsed.children.map((c) => c.name)).toEqual(original.children.map((c) => c.name));
  });
});
