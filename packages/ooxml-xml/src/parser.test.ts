import { describe, expect, it } from "vitest";
import { parseXml, XmlParseError } from "./parser.js";

describe("parseXml", () => {
  it("captures the XML declaration", () => {
    const doc = parseXml('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><root/>');
    expect(doc.declaration).toEqual({
      version: "1.0",
      encoding: "UTF-8",
      standalone: "yes",
    });
  });

  it("returns no declaration if absent", () => {
    const doc = parseXml("<root/>");
    expect(doc.declaration).toBeUndefined();
  });

  it("preserves element name prefix", () => {
    const doc = parseXml(
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
    );
    expect(doc.root.name.prefix).toBe("w");
    expect(doc.root.name.local).toBe("document");
    expect(doc.root.name.uri).toBe("http://schemas.openxmlformats.org/wordprocessingml/2006/main");
  });

  it("preserves attribute order", () => {
    const doc = parseXml('<root xmlns:w="urn:test" w:a="1" w:b="2" w:c="3"><child/></root>');
    const attrs = doc.root.attrs.map((a) => `${a.name.prefix}:${a.name.local}=${a.value}`);
    expect(attrs).toEqual(["xmlns:w=urn:test", "w:a=1", "w:b=2", "w:c=3"]);
  });

  it("decodes standard entities and character references in attributes and text", () => {
    const doc = parseXml('<root attr="a&amp;b">x&lt;y&#65;</root>');
    expect(doc.root.attrs[0]?.value).toBe("a&b");
    expect(doc.root.children[0]).toMatchObject({ kind: "text", value: "x<yA" });
  });

  it("treats unprefixed attributes as having no namespace", () => {
    const doc = parseXml('<w:p xmlns:w="urn:test" id="1" w:rsidR="abc"/>');
    const idAttr = doc.root.attrs.find((a) => a.name.local === "id");
    expect(idAttr?.name.uri).toBe("");
    expect(idAttr?.name.prefix).toBe("");
    const rsidAttr = doc.root.attrs.find((a) => a.name.local === "rsidR");
    expect(rsidAttr?.name.uri).toBe("urn:test");
    expect(rsidAttr?.name.prefix).toBe("w");
  });

  it('honors xml:space="preserve"', () => {
    const doc = parseXml('<w:t xmlns:w="urn:test" xml:space="preserve">  hello  </w:t>');
    expect(doc.root.xmlSpace).toBe("preserve");
    expect(doc.root.children).toHaveLength(1);
    expect(doc.root.children[0]).toMatchObject({ kind: "text", value: "  hello  " });
  });

  it("drops whitespace between elements by default", () => {
    const doc = parseXml("<root>\n  <a/>\n  <b/>\n</root>");
    const kinds = doc.root.children.map((c) => c.kind);
    expect(kinds).toEqual(["element", "element"]);
  });

  it("keeps interior whitespace when mixed with text", () => {
    const doc = parseXml("<root>hello <b/> world</root>");
    const kinds = doc.root.children.map((c) => c.kind);
    expect(kinds).toEqual(["text", "element", "text"]);
  });

  it("captures CDATA sections verbatim", () => {
    const doc = parseXml("<root><![CDATA[<not> & escaped]]></root>");
    expect(doc.root.children[0]).toEqual({ kind: "cdata", value: "<not> & escaped" });
  });

  it("captures comments", () => {
    const doc = parseXml("<root><!-- hi --></root>");
    expect(doc.root.children[0]).toEqual({ kind: "comment", value: " hi " });
  });

  it("captures processing instructions", () => {
    const doc = parseXml('<?mso-something foo="bar"?><root/>');
    expect(doc.prologue[0]).toMatchObject({ kind: "pi", target: "mso-something" });
  });

  it("rejects mismatched close tags", () => {
    expect(() => parseXml("<a><b></a></b>")).toThrow(XmlParseError);
  });

  it("rejects unbound prefixes", () => {
    expect(() => parseXml("<w:root/>")).toThrow(/Unbound namespace prefix/);
  });
});
