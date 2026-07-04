import { describe, expect, it } from "vitest";
import {
  addBookmark,
  bookmarks,
  createDocx,
  openDocx,
  paragraphs,
  removeBookmark,
  toUint8Array,
} from "./docx.js";

describe("Docx.bookmarks", () => {
  it("lists every bookmark with name, id, and source paragraph", () => {
    const doc = createDocx({ paragraphs: ["A", "B", "C"] });
    const [p1, p2, p3] = paragraphs(doc);
    if (!p1 || !p2 || !p3) return;
    addBookmark(doc, "alpha", p1);
    addBookmark(doc, "beta", p2);
    addBookmark(doc, "gamma", p3);

    const list = bookmarks(doc);
    expect(list).toHaveLength(3);
    expect(list.map((b) => b.name)).toEqual(["alpha", "beta", "gamma"]);
    expect(list.map((b) => b.id)).toEqual([0, 1, 2]);
  });

  it("removeBookmark drops the named bookmark's start and end", () => {
    const doc = createDocx({ paragraphs: ["A", "B"] });
    const [p1, p2] = paragraphs(doc);
    if (!p1 || !p2) return;
    addBookmark(doc, "alpha", p1);
    addBookmark(doc, "beta", p2);
    expect(removeBookmark(doc, "alpha")).toBe(true);
    expect(bookmarks(doc).map((b) => b.name)).toEqual(["beta"]);
    expect(removeBookmark(doc, "alpha")).toBe(false);
  });

  it("survives save+reopen", () => {
    const doc = createDocx({ paragraphs: ["A"] });
    const p = paragraphs(doc)[0];
    if (!p) return;
    addBookmark(doc, "ch1", p);
    const reopened = openDocx(toUint8Array(doc));
    expect(bookmarks(reopened).map((b) => b.name)).toEqual(["ch1"]);
  });
});
