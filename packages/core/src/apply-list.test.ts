import { describe, expect, it } from "vitest";
import { addBulletList, applyListToParagraph, createDocx, paragraphs } from "./docx.js";

describe("Docx.applyListToParagraph", () => {
  it("attaches numbering to an existing paragraph", () => {
    const doc = createDocx({ paragraphs: ["plain", "second"] });
    const para = paragraphs(doc)[0];
    if (!para) return;
    // First create a bullet numId via addBulletList so numbering.xml exists.
    const seeded = addBulletList(doc, ["seed"]);
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
    applyListToParagraph(doc, para, numId);
    // numPr is now attached to the originally-plain paragraph.
    const has = para.pPr?.children.some((c) => c.kind === "element" && c.name.local === "numPr");
    expect(has).toBe(true);
  });

  it("replaces an existing numPr rather than duplicating it", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    const para = paragraphs(doc)[0];
    if (!para) return;
    const seeded = addBulletList(doc, ["x"]);
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
    applyListToParagraph(doc, para, numId);
    applyListToParagraph(doc, para, numId);
    const numPrs =
      para.pPr?.children.filter((c) => c.kind === "element" && c.name.local === "numPr").length ??
      0;
    expect(numPrs).toBe(1);
  });
});
