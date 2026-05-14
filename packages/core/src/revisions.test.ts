import { getPart, writeOpcPackage } from "@word-kit/opc";
import { describe, expect, it } from "vitest";
import {
  acceptAllRevisions,
  createDocx,
  openDocx,
  rejectAllRevisions,
  text,
  toUint8Array,
  type Docx,
} from "./docx.js";

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
  const seed = createDocx({ paragraphs: [] });
  const docPart = getPart(seed.opc, "/word/document.xml");
  if (!docPart) throw new Error("no document part");
  docPart.data = new TextEncoder().encode(xml);
  const bytes = writeOpcPackage(seed.opc);
  return openDocx(bytes);
}

describe("Docx.acceptAllRevisions", () => {
  it("keeps inserted text and drops deleted text", () => {
    const doc = buildDocWithRevisions();
    const n = acceptAllRevisions(doc);
    expect(n).toBeGreaterThan(0);
    expect(text(doc)).toBe("Before INSERTED middle  after");
  });

  it("survives save+reopen with accepted state intact", () => {
    const doc = buildDocWithRevisions();
    acceptAllRevisions(doc);
    const reopened = openDocx(toUint8Array(doc));
    expect(text(reopened)).toBe("Before INSERTED middle  after");
  });
});

describe("Docx.rejectAllRevisions", () => {
  it("drops inserted text and keeps deleted text", () => {
    const doc = buildDocWithRevisions();
    const n = rejectAllRevisions(doc);
    expect(n).toBeGreaterThan(0);
    expect(text(doc)).toBe("Before  middle DELETED after");
  });

  it("survives save+reopen with rejected state intact", () => {
    const doc = buildDocWithRevisions();
    rejectAllRevisions(doc);
    const reopened = openDocx(toUint8Array(doc));
    expect(text(reopened)).toBe("Before  middle DELETED after");
  });
});

describe("revisions: empty document is a no-op", () => {
  it("returns 0 when nothing to accept", () => {
    const doc = createDocx({ paragraphs: ["plain text"] });
    expect(acceptAllRevisions(doc)).toBe(0);
    expect(rejectAllRevisions(doc)).toBe(0);
  });
});
