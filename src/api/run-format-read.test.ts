import { describe, expect, it } from "vitest";
import {
  appendParagraph,
  appendTextRun,
  createDocx,
  getRunFormat,
  paragraphs,
  paragraphText,
  setRunFormat,
} from "./index.js";

describe("getRunFormat", () => {
  it("returns an empty object for a run without an rPr", () => {
    const doc = createDocx({ paragraphs: ["plain"] });
    const para = paragraphs(doc)[0]!;
    const run = para.children.find((c) => c.kind === "run");
    if (!run || run.kind !== "run") throw new Error("expected a run");
    expect(getRunFormat(run)).toEqual({});
  });

  it("reads back bold/italic/underline/color/font", () => {
    const doc = createDocx({ paragraphs: [""] });
    appendParagraph(doc, "");
    const target = paragraphs(doc)[1]!;
    const run = appendTextRun(target, "hello", {
      bold: true,
      italic: true,
      underline: "single",
      color: "1F497D",
      font: "Helvetica",
      fontSizeHalfPoints: 24,
    });
    const fmt = getRunFormat(run);
    expect(fmt.bold).toBe(true);
    expect(fmt.italic).toBe(true);
    expect(fmt.underline).toBe("single");
    expect(fmt.color).toBe("1F497D");
    expect(fmt.font).toBe("Helvetica");
    expect(fmt.fontSizeHalfPoints).toBe(24);
  });

  it("round-trips through setRunFormat", () => {
    const doc = createDocx({ paragraphs: [""] });
    appendParagraph(doc, "");
    const a = paragraphs(doc)[1]!;
    const runA = appendTextRun(a, "src", {
      bold: true,
      color: "112233",
      fontSizeHalfPoints: 32,
    });
    appendParagraph(doc, "");
    const b = paragraphs(doc)[2]!;
    const runB = appendTextRun(b, "dst");
    setRunFormat(runB, getRunFormat(runA));
    expect(getRunFormat(runB)).toEqual(getRunFormat(runA));
  });

  it("skips unknown underline values rather than reporting them", () => {
    const doc = createDocx({ paragraphs: [""] });
    appendParagraph(doc, "");
    const target = paragraphs(doc)[1]!;
    const run = appendTextRun(target, "x", {
      bold: true,
      underline: "single",
    });
    // Inject an underline value the type doesn't permit; ensure we ignore it.
    const u = (run.rPr?.children ?? []).find((c) => c.kind === "element" && c.name.local === "u");
    if (u && u.kind === "element") {
      (u.attrs as { name: { local: string }; value: string }[]).find(
        (a) => a.name.local === "val",
      )!.value = "dashDotted";
    }
    const fmt = getRunFormat(run);
    expect(fmt.bold).toBe(true);
    expect(fmt.underline).toBeUndefined();
  });
});

describe("paragraphText re-export", () => {
  it("returns the visible text of a paragraph", () => {
    const doc = createDocx({ paragraphs: ["hello world"] });
    const p = paragraphs(doc)[0]!;
    expect(paragraphText(p)).toBe("hello world");
  });
});
