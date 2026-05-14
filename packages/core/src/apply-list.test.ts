import { describe, expect, it } from "vitest";
import { Docx } from "./docx.js";

describe("Docx.applyListToParagraph", () => {
  it("attaches numbering to an existing paragraph", () => {
    const doc = Docx.create({ paragraphs: ["plain", "second"] });
    const para = doc.paragraphs[0];
    if (!para) return;
    // First create a bullet numId via addBulletList so numbering.xml exists.
    const seeded = doc.addBulletList(["seed"]);
    const numId = (() => {
      // The seeded paragraph has the new bullet numId in its numPr.
      const p = seeded[0];
      const numPr = p?.pPr?.children.find((c) => c.kind === "element" && c.name.local === "numPr");
      if (numPr?.kind !== "element") return undefined;
      const numIdEl = numPr.children.find((c) => c.kind === "element" && c.name.local === "numId");
      if (numIdEl?.kind !== "element") return undefined;
      const v = numIdEl.attrs.find((a) => a.name.local === "val")?.value;
      return v ? Number.parseInt(v, 10) : undefined;
    })();
    expect(numId).toBeDefined();
    if (numId === undefined) return;
    doc.applyListToParagraph(para, numId);
    // numPr is now attached to the originally-plain paragraph.
    const has = para.pPr?.children.some((c) => c.kind === "element" && c.name.local === "numPr");
    expect(has).toBe(true);
  });

  it("replaces an existing numPr rather than duplicating it", () => {
    const doc = Docx.create({ paragraphs: ["body"] });
    const para = doc.paragraphs[0];
    if (!para) return;
    const seeded = doc.addBulletList(["x"]);
    const p = seeded[0];
    if (!p) return;
    const numPr = p.pPr?.children.find((c) => c.kind === "element" && c.name.local === "numPr");
    if (numPr?.kind !== "element") return;
    const numId = Number.parseInt(
      numPr.children.find((c) => c.kind === "element" && c.name.local === "numId")?.kind ===
        "element"
        ? "1"
        : "1",
      10,
    );
    doc.applyListToParagraph(para, numId);
    doc.applyListToParagraph(para, numId);
    const numPrs =
      para.pPr?.children.filter((c) => c.kind === "element" && c.name.local === "numPr").length ??
      0;
    expect(numPrs).toBe(1);
  });
});
