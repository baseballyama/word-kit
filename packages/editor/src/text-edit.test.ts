/**
 * Unit tests for the text-sync helpers that back canvas typing.
 */

import { createDocx, type WmlRun } from "@office-kit/docx";
import { describe, expect, it } from "vitest";
import { editorFor } from "./index.js";
import { runAtPath, setSimpleRunText } from "./text-edit.js";

describe("runAtPath", () => {
  it("resolves the run at a block/inline position", () => {
    const model = editorFor(createDocx({ paragraphs: ["Alpha", "Beta"] }));
    const run = runAtPath(model.doc, { block: 1, inline: 0 });
    expect(run?.kind).toBe("run");
    const text = run?.pieces.map((p) => (p.kind === "text" ? p.value : "")).join("");
    expect(text).toBe("Beta");
  });

  it("returns undefined for out-of-range positions", () => {
    const model = editorFor(createDocx({ paragraphs: ["only"] }));
    expect(runAtPath(model.doc, { block: 5, inline: 0 })).toBeUndefined();
  });
});

describe("setSimpleRunText", () => {
  it("replaces a text-only run's content", () => {
    const model = editorFor(createDocx({ paragraphs: ["old"] }));
    const run = runAtPath(model.doc, { block: 0, inline: 0 })!;
    expect(setSimpleRunText(run, "new text")).toBe(true);
    expect(run.pieces).toEqual([{ kind: "text", value: "new text", preserveSpace: false }]);
  });

  it("preserves leading/trailing space with xml:space", () => {
    const model = editorFor(createDocx({ paragraphs: ["x"] }));
    const run = runAtPath(model.doc, { block: 0, inline: 0 })!;
    setSimpleRunText(run, " padded ");
    expect(run.pieces[0]).toMatchObject({ kind: "text", preserveSpace: true });
  });

  it("refuses to overwrite runs carrying non-text pieces", () => {
    const run: WmlRun = {
      kind: "run",
      pieces: [{ kind: "tab" }, { kind: "text", value: "a", preserveSpace: false }],
      extras: [],
    };
    expect(setSimpleRunText(run, "b")).toBe(false);
    // Original pieces untouched.
    expect(run.pieces.length).toBe(2);
  });
});
