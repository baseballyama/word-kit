import { describe, expect, it } from "vitest";
import { Docx } from "./docx.js";

describe("Docx.create", () => {
  it("returns a docx with one empty paragraph by default", () => {
    const doc = Docx.create();
    expect(doc.paragraphs.length).toBe(1);
    expect(doc.text).toBe("");
  });

  it("seeds the body when paragraphs are supplied", () => {
    const doc = Docx.create({ paragraphs: ["Hello", "World"] });
    expect(doc.paragraphs.length).toBe(2);
    expect(doc.text).toBe("Hello\nWorld");
  });
});

describe("Docx.open round-trip", () => {
  it("opens what create + toUint8Array produced", () => {
    const original = Docx.create({ paragraphs: ["First", "Second"] });
    const bytes = original.toUint8Array();
    expect(bytes.length).toBeGreaterThan(0);
    const reopened = Docx.open(bytes);
    expect(reopened.paragraphs.length).toBe(2);
    expect(reopened.text).toBe("First\nSecond");
  });
});

describe("Docx.replaceText", () => {
  it("rewrites template placeholders and survives a save+reopen", () => {
    const doc = Docx.create({ paragraphs: ["Hello {{name}}!"] });
    expect(doc.replaceText("{{name}}", "Alice")).toBe(1);
    expect(doc.text).toBe("Hello Alice!");
    const reopened = Docx.open(doc.toUint8Array());
    expect(reopened.text).toBe("Hello Alice!");
  });

  it("handles a placeholder that gets split across runs after the user manually edits in Word", () => {
    // Simulate the canonical split-run scenario by constructing the
    // document.xml directly and re-reading it as a Docx.
    const fresh = Docx.create({ paragraphs: ["Hello placeholder"] });
    const bytes = fresh.toUint8Array();
    const reopened = Docx.open(bytes);

    // First, plant a split placeholder in the AST by editing pieces
    // directly. This is the worst case for replaceText.
    const para = reopened.paragraphs[0];
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

    expect(reopened.replaceText("{{name}}", "山田太郎")).toBe(1);
    expect(reopened.text).toBe("Hello 山田太郎");
  });
});

describe("Docx.appendParagraph", () => {
  it("appends text and keeps it through a save round-trip", () => {
    const doc = Docx.create();
    doc.appendParagraph("Added paragraph");
    expect(doc.text).toBe("\nAdded paragraph");
    const reopened = Docx.open(doc.toUint8Array());
    expect(reopened.paragraphs.at(-1)).toBeDefined();
    expect(reopened.text.endsWith("Added paragraph")).toBe(true);
  });

  it("supports style and bold options", () => {
    const doc = Docx.create({ paragraphs: [] });
    doc.appendParagraph("Title", { style: "Heading1", bold: true });
    const bytes = doc.toUint8Array();
    const reopened = Docx.open(bytes);
    const para = reopened.paragraphs[0];
    expect(para).toBeDefined();
    expect(para?.pPr).toBeDefined();
    expect(para?.children[0]).toMatchObject({ kind: "run" });
    expect(reopened.text).toBe("Title");
  });
});
