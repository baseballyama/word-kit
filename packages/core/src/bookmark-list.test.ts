import { describe, expect, it } from "vitest";
import { Docx } from "./docx.js";

describe("Docx.bookmarks", () => {
  it("lists every bookmark with name, id, and source paragraph", () => {
    const doc = Docx.create({ paragraphs: ["A", "B", "C"] });
    const [p1, p2, p3] = doc.paragraphs;
    if (!p1 || !p2 || !p3) return;
    doc.addBookmark("alpha", p1);
    doc.addBookmark("beta", p2);
    doc.addBookmark("gamma", p3);

    const list = doc.bookmarks;
    expect(list).toHaveLength(3);
    expect(list.map((b) => b.name)).toEqual(["alpha", "beta", "gamma"]);
    expect(list.map((b) => b.id)).toEqual([0, 1, 2]);
  });

  it("removeBookmark drops the named bookmark's start and end", () => {
    const doc = Docx.create({ paragraphs: ["A", "B"] });
    const [p1, p2] = doc.paragraphs;
    if (!p1 || !p2) return;
    doc.addBookmark("alpha", p1);
    doc.addBookmark("beta", p2);
    expect(doc.removeBookmark("alpha")).toBe(true);
    expect(doc.bookmarks.map((b) => b.name)).toEqual(["beta"]);
    expect(doc.removeBookmark("alpha")).toBe(false);
  });

  it("survives save+reopen", () => {
    const doc = Docx.create({ paragraphs: ["A"] });
    const p = doc.paragraphs[0];
    if (!p) return;
    doc.addBookmark("ch1", p);
    const reopened = Docx.open(doc.toUint8Array());
    expect(reopened.bookmarks.map((b) => b.name)).toEqual(["ch1"]);
  });
});
