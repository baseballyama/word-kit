import { describe, expect, it } from "vitest";
import { parseXml } from "./parser.js";
import { serializeXml } from "./serializer.js";

function roundTrip(xml: string): string {
  return serializeXml(parseXml(xml));
}

describe("serializeXml", () => {
  it("re-emits the XML declaration with original version/encoding/standalone", () => {
    const out = roundTrip('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><root/>');
    expect(out.startsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n')).toBe(
      true,
    );
  });

  it("preserves prefixes on element names", () => {
    const out = roundTrip(
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
    );
    expect(out).toContain("<w:document");
    expect(out).toContain("xmlns:w=");
  });

  it("preserves attribute order", () => {
    const xml = '<root xmlns:w="urn:test" w:a="1" w:b="2" w:c="3"/>';
    expect(roundTrip(xml).replace(/<\?xml[^?]*\?>\r?\n?/, "")).toBe(xml);
  });

  it("preserves CDATA, comments, and PIs", () => {
    const xml = "<root><!-- hi --><![CDATA[<x>]]><?mso-x foo?></root>";
    const out = roundTrip(xml);
    expect(out).toContain("<!-- hi -->");
    expect(out).toContain("<![CDATA[<x>]]>");
    expect(out).toContain("<?mso-x foo?>");
  });

  it("re-emits self-closing form when the source used it", () => {
    const out = roundTrip('<root xmlns:w="urn:test"><w:p/></root>');
    expect(out).toContain("<w:p/>");
  });

  it("encodes attribute values that contain reserved characters", () => {
    const xml = '<root attr="a&amp;b&lt;c"/>';
    const out = roundTrip(xml);
    expect(out).toContain('attr="a&amp;b&lt;c"');
  });

  it('preserves xml:space="preserve" whitespace inside text nodes', () => {
    const xml = '<w:t xmlns:w="urn:test" xml:space="preserve">  hello  </w:t>';
    const out = roundTrip(xml);
    expect(out).toContain(">  hello  <");
  });

  it("round-trips a sample Word document.xml fragment to AST-equivalence", () => {
    const xml = [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
      "<w:body>",
      '<w:p><w:r><w:t xml:space="preserve">Hello </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>world</w:t></w:r></w:p>',
      "</w:body>",
      "</w:document>",
    ].join("");
    const first = parseXml(xml);
    const second = parseXml(serializeXml(first));
    expect(second).toEqual(first);
  });
});
