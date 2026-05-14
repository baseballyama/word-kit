import { describe, expect, it } from "vitest";
import { Docx } from "./docx.js";

describe("Docx.clone", () => {
  it("returns an independent copy that can be edited separately", () => {
    const original = Docx.create({ paragraphs: ["template {{name}}"] });
    const a = original.clone();
    const b = original.clone();
    a.replaceText("{{name}}", "Alice");
    b.replaceText("{{name}}", "Bob");
    expect(a.text).toContain("Alice");
    expect(b.text).toContain("Bob");
    expect(original.text).toContain("{{name}}");
  });

  it("copies side parts (styles, comments, footnotes)", () => {
    const doc = Docx.create({ paragraphs: ["body"] });
    const p = doc.paragraphs[0];
    if (!p) return;
    doc.addComment(p, { author: "R", text: "hi" });
    doc.addStyle({ type: "paragraph", styleId: "MyHeading", name: "My Heading" });
    const cloned = doc.clone();
    expect(cloned.commentsPart?.comments).toHaveLength(1);
    expect(
      cloned.stylesPart?.styles.some((s) =>
        s.attrs.some((a) => a.name.local === "styleId" && a.value === "MyHeading"),
      ),
    ).toBe(true);
  });
});
