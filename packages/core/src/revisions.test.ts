import { getPart, writeOpcPackage } from "@word-kit/opc";
import { describe, expect, it } from "vitest";
import { Docx } from "./docx.js";

function buildDocWithRevisions(): Docx {
  // Construct a document with both <w:ins> and <w:del> revisions.
  const xml = [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    "<w:body>",
    "<w:p>",
    "<w:r><w:t>Before </w:t></w:r>",
    '<w:ins w:id="1" w:author="X" w:date="2026-05-14T00:00:00Z">',
    "<w:r><w:t>INSERTED</w:t></w:r>",
    "</w:ins>",
    "<w:r><w:t> middle </w:t></w:r>",
    '<w:del w:id="2" w:author="X" w:date="2026-05-14T00:00:00Z">',
    "<w:r><w:delText>DELETED</w:delText></w:r>",
    "</w:del>",
    "<w:r><w:t> after</w:t></w:r>",
    "</w:p>",
    "</w:body>",
    "</w:document>",
  ].join("");
  // Embed into a real package by overwriting Docx.create's document.xml
  // with our crafted XML, then re-opening.
  const seed = Docx.create({ paragraphs: [] });
  const docPart = getPart(seed.opc, "/word/document.xml");
  if (!docPart) throw new Error("no document part");
  docPart.data = new TextEncoder().encode(xml);
  const bytes = writeOpcPackage(seed.opc);
  return Docx.open(bytes);
}

describe("Docx.acceptAllRevisions", () => {
  it("keeps inserted text and drops deleted text", () => {
    const doc = buildDocWithRevisions();
    const n = doc.acceptAllRevisions();
    expect(n).toBeGreaterThan(0);
    expect(doc.text).toBe("Before INSERTED middle  after");
  });

  it("survives save+reopen with accepted state intact", () => {
    const doc = buildDocWithRevisions();
    doc.acceptAllRevisions();
    const reopened = Docx.open(doc.toUint8Array());
    expect(reopened.text).toBe("Before INSERTED middle  after");
  });
});

describe("Docx.rejectAllRevisions", () => {
  it("drops inserted text and keeps deleted text", () => {
    const doc = buildDocWithRevisions();
    const n = doc.rejectAllRevisions();
    expect(n).toBeGreaterThan(0);
    expect(doc.text).toBe("Before  middle DELETED after");
  });

  it("survives save+reopen with rejected state intact", () => {
    const doc = buildDocWithRevisions();
    doc.rejectAllRevisions();
    const reopened = Docx.open(doc.toUint8Array());
    expect(reopened.text).toBe("Before  middle DELETED after");
  });
});

describe("revisions: empty document is a no-op", () => {
  it("returns 0 when nothing to accept", () => {
    const doc = Docx.create({ paragraphs: ["plain text"] });
    expect(doc.acceptAllRevisions()).toBe(0);
    expect(doc.rejectAllRevisions()).toBe(0);
  });
});
