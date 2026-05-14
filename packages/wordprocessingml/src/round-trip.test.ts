import { parseXml, serializeXml } from "@word-kit/ooxml-xml";
import { describe, expect, it } from "vitest";
import { parseWmlDocument } from "./parser.js";
import { writeWmlDocument } from "./writer.js";

function fullCycle(xml: string): string {
  const parsed = parseXml(xml);
  const wml = parseWmlDocument(parsed);
  const xmlOut = writeWmlDocument(wml);
  return serializeXml(xmlOut);
}

describe("WML round-trip", () => {
  it("round-trips an empty document", () => {
    const xml = [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
      "<w:body/>",
      "</w:document>",
    ].join("");
    const out = fullCycle(xml);
    expect(out).toContain("<w:body/>");
    expect(out).toContain("xmlns:w=");
  });

  it("round-trips a single empty paragraph", () => {
    const xml = [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
      "<w:body><w:p/></w:body>",
      "</w:document>",
    ].join("");
    const out = fullCycle(xml);
    expect(out).toContain("<w:p/>");
  });

  it("round-trips a paragraph with a single text run", () => {
    const xml = [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
      "<w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body>",
      "</w:document>",
    ].join("");
    const out = fullCycle(xml);
    expect(out).toContain("<w:t>Hello</w:t>");
  });

  it('preserves xml:space="preserve" on text runs', () => {
    const xml = [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
      '<w:body><w:p><w:r><w:t xml:space="preserve">  hello  </w:t></w:r></w:p></w:body>',
      "</w:document>",
    ].join("");
    const out = fullCycle(xml);
    expect(out).toContain('xml:space="preserve"');
    expect(out).toContain(">  hello  <");
  });

  it("preserves runs with multiple text/tab/break pieces", () => {
    const xml = [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
      "<w:body><w:p><w:r>",
      "<w:t>A</w:t><w:tab/><w:t>B</w:t><w:br/><w:t>C</w:t>",
      "</w:r></w:p></w:body>",
      "</w:document>",
    ].join("");
    const out = fullCycle(xml);
    expect(out).toContain("<w:t>A</w:t><w:tab/><w:t>B</w:t><w:br/><w:t>C</w:t>");
  });

  it("preserves rPr verbatim", () => {
    const xml = [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
      "<w:body><w:p><w:r>",
      '<w:rPr><w:b/><w:sz w:val="24"/></w:rPr>',
      "<w:t>Hi</w:t>",
      "</w:r></w:p></w:body>",
      "</w:document>",
    ].join("");
    const out = fullCycle(xml);
    expect(out).toContain('<w:rPr><w:b/><w:sz w:val="24"/></w:rPr>');
  });

  it("preserves pPr verbatim and a trailing sectPr", () => {
    const xml = [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
      "<w:body>",
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>',
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>',
      "</w:body>",
      "</w:document>",
    ].join("");
    const out = fullCycle(xml);
    expect(out).toContain('<w:pStyle w:val="Heading1"/>');
    expect(out).toContain('<w:pgSz w:w="12240" w:h="15840"/>');
    expect(out.indexOf("<w:sectPr>")).toBeGreaterThan(out.indexOf("<w:p>"));
  });

  it("preserves unrecognized blocks (e.g. tables) as raw passthrough", () => {
    const xml = [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
      "<w:body>",
      "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>X</w:t></w:r></w:p></w:tc></w:tr></w:tbl>",
      "</w:body>",
      "</w:document>",
    ].join("");
    const out = fullCycle(xml);
    expect(out).toContain("<w:tbl>");
    expect(out).toContain("<w:t>X</w:t>");
  });

  it("returns AST-equivalent output (re-parse equals first parse)", () => {
    const xml = [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
      "<w:body>",
      "<w:p>",
      '<w:pPr><w:jc w:val="center"/></w:pPr>',
      '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Bold </w:t></w:r>',
      "<w:r><w:t>tail</w:t></w:r>",
      "</w:p>",
      "</w:body>",
      "</w:document>",
    ].join("");
    const first = parseXml(xml);
    const second = parseXml(serializeXml(writeWmlDocument(parseWmlDocument(first))));
    expect(second).toEqual(first);
  });
});
