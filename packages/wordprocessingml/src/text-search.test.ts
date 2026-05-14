import { parseXml, serializeXml } from "@word-kit/ooxml-xml";
import { describe, expect, it } from "vitest";
import { parseWmlDocument } from "./parser.js";
import { documentText, findText, paragraphText, replaceText } from "./text-search.js";
import type { WmlDocument } from "./types.js";
import { writeWmlDocument } from "./writer.js";

function doc(bodyInner: string): WmlDocument {
  const xml = [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
    `<w:body>${bodyInner}</w:body>`,
    "</w:document>",
  ].join("");
  return parseWmlDocument(parseXml(xml));
}

function render(d: WmlDocument): string {
  return serializeXml(writeWmlDocument(d));
}

describe("text-search: paragraphText / documentText", () => {
  it("flattens text across runs", () => {
    const d = doc("<w:p><w:r><w:t>{{na</w:t></w:r><w:r><w:t>me}}</w:t></w:r></w:p>");
    expect(paragraphText(d.body.blocks[0] as never)).toBe("{{name}}");
    expect(documentText(d)).toBe("{{name}}");
  });

  it("ignores tab and break pieces in the flat text", () => {
    const d = doc("<w:p><w:r><w:t>A</w:t><w:tab/><w:t>B</w:t><w:br/><w:t>C</w:t></w:r></w:p>");
    expect(paragraphText(d.body.blocks[0] as never)).toBe("ABC");
  });
});

describe("text-search: findText", () => {
  it("finds a literal across runs", () => {
    const d = doc("<w:p><w:r><w:t>{{na</w:t></w:r><w:r><w:t>me}}</w:t></w:r></w:p>");
    const matches = findText(d, "{{name}}");
    expect(matches).toHaveLength(1);
    expect(matches[0]).toMatchObject({ text: "{{name}}", start: 0, end: 8 });
  });

  it("finds all occurrences (string mode)", () => {
    const d = doc("<w:p><w:r><w:t>x{{a}}y{{a}}z</w:t></w:r></w:p>");
    const matches = findText(d, "{{a}}");
    expect(matches).toHaveLength(2);
  });

  it("supports regular expressions with capture groups", () => {
    const d = doc(
      "<w:p><w:r><w:t>Hello </w:t></w:r><w:r><w:t>{{ first }} and {{ second }}!</w:t></w:r></w:p>",
    );
    const matches = findText(d, /\{\{\s*(\w+)\s*\}\}/g);
    expect(matches).toHaveLength(2);
    expect(matches[0]?.captures[0]).toBe("first");
    expect(matches[1]?.captures[0]).toBe("second");
  });
});

describe("text-search: replaceText", () => {
  it("replaces a literal contained within a single run", () => {
    const d = doc("<w:p><w:r><w:t>Hello {{name}}!</w:t></w:r></w:p>");
    expect(replaceText(d, "{{name}}", "Alice")).toBe(1);
    expect(paragraphText(d.body.blocks[0] as never)).toBe("Hello Alice!");
  });

  it("replaces a literal split across two runs", () => {
    const d = doc("<w:p><w:r><w:t>{{na</w:t></w:r><w:r><w:t>me}} ok</w:t></w:r></w:p>");
    expect(replaceText(d, "{{name}}", "Bob")).toBe(1);
    expect(paragraphText(d.body.blocks[0] as never)).toBe("Bob ok");
  });

  it("replaces a literal split across many runs", () => {
    const d = doc(
      "<w:p>" +
        "<w:r><w:t>{{</w:t></w:r>" +
        "<w:r><w:t>na</w:t></w:r>" +
        "<w:r><w:t>me</w:t></w:r>" +
        "<w:r><w:t>}}</w:t></w:r>" +
        "</w:p>",
    );
    expect(replaceText(d, "{{name}}", "山田太郎")).toBe(1);
    expect(paragraphText(d.body.blocks[0] as never)).toBe("山田太郎");
  });

  it("preserves rPr of the first containing run", () => {
    const d = doc(
      "<w:p>" +
        "<w:r><w:rPr><w:b/></w:rPr><w:t>{{na</w:t></w:r>" +
        "<w:r><w:rPr><w:i/></w:rPr><w:t>me}}</w:t></w:r>" +
        "</w:p>",
    );
    expect(replaceText(d, "{{name}}", "Alice")).toBe(1);
    const out = render(d);
    expect(out).toContain("<w:rPr><w:b/></w:rPr><w:t>Alice</w:t>");
    expect(out).toContain("<w:rPr><w:i/></w:rPr>");
  });

  it("preserves trailing text in the last touched run", () => {
    const d = doc(
      "<w:p>" + "<w:r><w:t>Hello {{na</w:t></w:r>" + "<w:r><w:t>me}} world</w:t></w:r>" + "</w:p>",
    );
    expect(replaceText(d, "{{name}}", "X")).toBe(1);
    expect(paragraphText(d.body.blocks[0] as never)).toBe("Hello X world");
  });

  it("replaces using a callback for each match", () => {
    const d = doc("<w:p><w:r><w:t>{{name}}, {{age}}, {{name}}</w:t></w:r></w:p>");
    const values: Record<string, string> = { name: "Yuki", age: "42" };
    const count = replaceText(d, /\{\{\s*(\w+)\s*\}\}/g, (m) => values[m.captures[0] ?? ""] ?? "");
    expect(count).toBe(3);
    expect(paragraphText(d.body.blocks[0] as never)).toBe("Yuki, 42, Yuki");
  });

  it("returns 0 when nothing matches", () => {
    const d = doc("<w:p><w:r><w:t>hello</w:t></w:r></w:p>");
    expect(replaceText(d, "missing", "x")).toBe(0);
  });

  it('forces xml:space="preserve" when replacement introduces edge whitespace', () => {
    const d = doc("<w:p><w:r><w:t>X</w:t></w:r></w:p>");
    expect(replaceText(d, "X", " spaced ")).toBe(1);
    const out = render(d);
    expect(out).toContain('xml:space="preserve"');
    expect(out).toContain("> spaced <");
  });

  it("survives serialization to bytes and back", () => {
    const d = doc("<w:p><w:r><w:t>{{na</w:t></w:r><w:r><w:t>me}}</w:t></w:r></w:p>");
    replaceText(d, "{{name}}", "Final");
    const xmlOut = render(d);
    const reparsed = parseWmlDocument(parseXml(xmlOut));
    expect(paragraphText(reparsed.body.blocks[0] as never)).toBe("Final");
  });
});
