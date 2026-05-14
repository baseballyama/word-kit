import { describe, expect, it } from "vitest";
import {
  appendParagraph,
  createDocx,
  openDocx,
  paragraphs,
  replaceText,
  text,
  toUint8Array,
} from "./docx.js";

describe("Docx.create", () => {
  it("returns a docx with one empty paragraph by default", () => {
    const doc = createDocx();
    expect(paragraphs(doc).length).toBe(1);
    expect(text(doc)).toBe("");
  });

  it("seeds the body when paragraphs are supplied", () => {
    const doc = createDocx({ paragraphs: ["Hello", "World"] });
    expect(paragraphs(doc).length).toBe(2);
    expect(text(doc)).toBe("Hello\nWorld");
  });
});

describe("Docx.open round-trip", () => {
  it("opens what create + toUint8Array produced", () => {
    const original = createDocx({ paragraphs: ["First", "Second"] });
    const bytes = toUint8Array(original);
    expect(bytes.length).toBeGreaterThan(0);
    const reopened = openDocx(bytes);
    expect(paragraphs(reopened).length).toBe(2);
    expect(text(reopened)).toBe("First\nSecond");
  });
});

describe("Docx.replaceText", () => {
  it("rewrites template placeholders and survives a save+reopen", () => {
    const doc = createDocx({ paragraphs: ["Hello {{name}}!"] });
    expect(replaceText(doc, "{{name}}", "Alice")).toBe(1);
    expect(text(doc)).toBe("Hello Alice!");
    const reopened = openDocx(toUint8Array(doc));
    expect(text(reopened)).toBe("Hello Alice!");
  });

  it("handles a placeholder that gets split across runs after the user manually edits in Word", () => {
    // Simulate the canonical split-run scenario by constructing the
    // document.xml directly and re-reading it as a Docx.
    const fresh = createDocx({ paragraphs: ["Hello placeholder"] });
    const bytes = toUint8Array(fresh);
    const reopened = openDocx(bytes);

    // First, plant a split placeholder in the AST by editing pieces
    // directly. This is the worst case for replaceText.
    const para = paragraphs(reopened)[0];
    expect(para).toBeDefined();
    if (!para) return;
    para.children = [
      {
        kind: "run",
        pieces: [{ kind: "text", value: "Hello ", preserveSpace: true }],
        extras: [],
      },
      {
        kind: "run",
        pieces: [{ kind: "text", value: "{{na", preserveSpace: false }],
        extras: [],
      },
      {
        kind: "run",
        pieces: [{ kind: "text", value: "me}}", preserveSpace: false }],
        extras: [],
      },
    ];

    expect(replaceText(reopened, "{{name}}", "山田太郎")).toBe(1);
    expect(text(reopened)).toBe("Hello 山田太郎");
  });
});

describe("Docx.appendParagraph", () => {
  it("appends text and keeps it through a save round-trip", () => {
    const doc = createDocx();
    appendParagraph(doc, "Added paragraph");
    expect(text(doc)).toBe("\nAdded paragraph");
    const reopened = openDocx(toUint8Array(doc));
    expect(paragraphs(reopened).at(-1)).toBeDefined();
    expect(text(reopened).endsWith("Added paragraph")).toBe(true);
  });

  it("supports style and bold options", () => {
    const doc = createDocx({ paragraphs: [] });
    appendParagraph(doc, "Title", { style: "Heading1", bold: true });
    const bytes = toUint8Array(doc);
    const reopened = openDocx(bytes);
    const para = paragraphs(reopened)[0];
    expect(para).toBeDefined();
    expect(para?.pPr).toBeDefined();
    expect(para?.children[0]).toMatchObject({ kind: "run" });
    expect(text(reopened)).toBe("Title");
  });
});
