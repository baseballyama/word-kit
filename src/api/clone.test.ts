import { describe, expect, it } from "vitest";
import {
  addComment,
  addStyle,
  clone,
  commentsPart,
  createDocx,
  paragraphs,
  replaceText,
  stylesPart,
  text,
} from "./docx.js";

describe("Docx.clone", () => {
  it("returns an independent copy that can be edited separately", () => {
    const original = createDocx({ paragraphs: ["template {{name}}"] });
    const a = clone(original);
    const b = clone(original);
    replaceText(a, "{{name}}", "Alice");
    replaceText(b, "{{name}}", "Bob");
    expect(text(a)).toContain("Alice");
    expect(text(b)).toContain("Bob");
    expect(text(original)).toContain("{{name}}");
  });

  it("copies side parts (styles, comments, footnotes)", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    const p = paragraphs(doc)[0];
    if (!p) return;
    addComment(doc, p, { author: "R", text: "hi" });
    addStyle(doc, { type: "paragraph", styleId: "MyHeading", name: "My Heading" });
    const cloned = clone(doc);
    expect(commentsPart(cloned)?.comments).toHaveLength(1);
    expect(
      stylesPart(cloned)?.styles.some((s) =>
        s.attrs.some((a) => a.name.local === "styleId" && a.value === "MyHeading"),
      ),
    ).toBe(true);
  });
});
