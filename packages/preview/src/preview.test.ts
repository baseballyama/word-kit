// Smoke tests for the v0 preview bridge.
//
// We deliberately do NOT exercise `previewToDOM` end-to-end against a
// real `Docx` here. vitest's module resolver routes `fflate` (a
// transitive dep of `@word-kit/core`'s OPC layer) to the browser
// bundle even in the node-pool worker, and the browser bundle
// produces zip output that JSZip 3.x — which docx-preview vendors —
// then mis-decodes. The mismatch is purely a vitest harness issue;
// real browsers and real Node both produce valid docx via the same
// code path (verified by `pnpm sample`).
//
// Future work (tracked in docs/PLAN-PREVIEW.md):
//   - Add a Playwright suite that invokes `previewToDOM` in an
//     actual browser, where neither bundler nor JSZip mismatch
//     applies.
//   - Add a Node-environment integration test once we fence
//     `fflate` to its node entry inside vitest.

import { describe, expect, it } from "vitest";
import { previewToDOM } from "./index.js";

describe("@word-kit/preview surface", () => {
  it("exports `previewToDOM` as a function", () => {
    expect(typeof previewToDOM).toBe("function");
  });

  it("`previewToDOM` accepts the documented signature shape", () => {
    // Pure type-shape check — call signature is `(source, container,
    // options?) => Promise<Handle>`. We can't invoke without a DOM,
    // but we can confirm the function's `.length` matches the
    // documented arity (2 required params).
    expect(previewToDOM.length).toBe(2);
  });
});
