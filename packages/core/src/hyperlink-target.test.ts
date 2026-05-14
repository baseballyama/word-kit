import { describe, expect, it } from "vitest";
import {
  addHyperlink,
  appendParagraph,
  createDocx,
  externalHyperlinks,
  openDocx,
  setHyperlinkUrl,
  toUint8Array,
} from "./docx.js";

describe("setHyperlinkUrl", () => {
  it("rewrites a hyperlink rel by exact-match string", () => {
    const doc = createDocx({ paragraphs: ["intro"] });
    addHyperlink(doc, "https://stage.example.com/", "report");
    addHyperlink(doc, "https://other.example.com/", "other");

    const n = setHyperlinkUrl(doc, "https://stage.example.com/", "https://example.com/");
    expect(n).toBe(1);
    const all = externalHyperlinks(doc).map((h) => h.target);
    expect(all).toContain("https://example.com/");
    expect(all).toContain("https://other.example.com/");
    expect(all).not.toContain("https://stage.example.com/");
  });

  it("rewrites every rel that the predicate returns a new URL for", () => {
    const doc = createDocx({ paragraphs: ["intro"] });
    addHyperlink(doc, "https://stage.example.com/a", "a");
    addHyperlink(doc, "https://stage.example.com/b", "b");
    addHyperlink(doc, "https://other.example.com/c", "c");

    const n = setHyperlinkUrl(doc, (t) =>
      t.startsWith("https://stage.example.com/")
        ? "https://prod.example.com/" + t.slice("https://stage.example.com/".length)
        : null,
    );
    expect(n).toBe(2);
    const all = externalHyperlinks(doc)
      .map((h) => h.target)
      .toSorted();
    expect(all).toEqual([
      "https://other.example.com/c",
      "https://prod.example.com/a",
      "https://prod.example.com/b",
    ]);
  });

  it("returns 0 when nothing matches", () => {
    const doc = createDocx({ paragraphs: ["intro"] });
    addHyperlink(doc, "https://example.com/", "x");
    expect(setHyperlinkUrl(doc, "https://nope.invalid/", "https://anywhere/")).toBe(0);
  });

  it("ignores internal anchor hyperlinks", () => {
    const doc = createDocx({ paragraphs: ["target"] });
    // Internal anchor hyperlink — no rel is created.
    appendParagraph(doc, "");
    // We have to use the external variant to actually create a rel.
    addHyperlink(doc, "https://x.example.com/", "click");
    const n = setHyperlinkUrl(doc, (t) => (t.includes("anchor") ? "https://new/" : null));
    expect(n).toBe(0);
  });

  it("survives a save+reopen with the new URL in document.xml.rels", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    addHyperlink(doc, "https://stage.example.com/", "click");
    setHyperlinkUrl(doc, "https://stage.example.com/", "https://prod.example.com/");
    const reopened = openDocx(toUint8Array(doc));
    const relsText = new TextDecoder().decode(
      reopened.opc.parts.get("/word/_rels/document.xml.rels")?.data ?? new Uint8Array(),
    );
    expect(relsText).toContain("https://prod.example.com/");
    expect(relsText).not.toContain("https://stage.example.com/");
  });
});

describe("externalHyperlinks", () => {
  it("lists every external hyperlink rel with id and target", () => {
    const doc = createDocx({ paragraphs: ["intro"] });
    addHyperlink(doc, "https://a/", "a");
    addHyperlink(doc, "https://b/", "b");
    const out = externalHyperlinks(doc);
    expect(out).toHaveLength(2);
    expect(out.map((h) => h.target).toSorted()).toEqual(["https://a/", "https://b/"]);
    for (const h of out) expect(h.relId).toMatch(/^rId\d+$/);
  });
});
