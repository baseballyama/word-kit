// Performance smoke test. The library hasn't been profiled; this file
// exists to catch accidental O(n²) / quadratic-string-concat regressions
// rather than to assert a specific throughput. Bounds are deliberately
// loose so CI on a slow runner doesn't false-fail; they will, however,
// fire if a future change makes a common operation an order of magnitude
// slower.
//
// Numbers measured locally (Apple M-series, Node 24): each block well
// under 1 s. CI bound is ~10× that, with explicit headroom for shared
// runners.

import { describe, expect, it } from "vitest";
import {
  addBulletList,
  addHeader,
  addPageNumberFooter,
  addTable,
  appendHeading,
  appendParagraph,
  createDocx,
  MARGINS_NORMAL,
  openDocx,
  PAGE_SIZE_A4,
  paragraphs,
  replaceText,
  setCoreProperties,
  setPageMargins,
  setPageSize,
  text,
  toUint8Array,
  validate,
} from "./index.js";

function timeit(label: string, fn: () => void): number {
  const start = performance.now();
  fn();
  const elapsed = performance.now() - start;
  // eslint-disable-next-line no-console
  console.log(`  ⏱ ${label}: ${elapsed.toFixed(0)} ms`);
  return elapsed;
}

const BUDGET_MS = 10_000;

describe("perf smoke (regression bounds)", () => {
  it("appends 10 000 paragraphs without going quadratic", () => {
    const doc = createDocx({ paragraphs: [] });
    const elapsed = timeit("append 10k paragraphs", () => {
      for (let i = 0; i < 10_000; i++) {
        appendParagraph(doc, `body line ${i}`);
      }
    });
    expect(paragraphs(doc).length).toBe(10_000);
    expect(elapsed).toBeLessThan(BUDGET_MS);
  });

  it("serializes a 10k-paragraph doc within budget and round-trips back", () => {
    const doc = createDocx({ paragraphs: [] });
    for (let i = 0; i < 10_000; i++) appendParagraph(doc, `line ${i}`);
    let bytes!: Uint8Array;
    const writeMs = timeit("toUint8Array on 10k", () => {
      bytes = toUint8Array(doc);
    });
    expect(writeMs).toBeLessThan(BUDGET_MS);
    expect(bytes.byteLength).toBeGreaterThan(0);
    let reopened!: ReturnType<typeof openDocx>;
    const readMs = timeit("openDocx on the same bytes", () => {
      reopened = openDocx(bytes);
    });
    expect(readMs).toBeLessThan(BUDGET_MS);
    expect(paragraphs(reopened).length).toBe(10_000);
  });

  it("100 tables × 10 cells rebuild + serialize stays bounded", () => {
    const doc = createDocx({ paragraphs: [] });
    const elapsed = timeit("build 100 tables", () => {
      for (let i = 0; i < 100; i++) {
        const rows: string[][] = [];
        for (let r = 0; r < 5; r++) {
          rows.push(Array.from({ length: 2 }, (_, c) => `t${i}-r${r}-c${c}`));
        }
        addTable(doc, rows);
      }
    });
    expect(elapsed).toBeLessThan(BUDGET_MS);
    const writeMs = timeit("toUint8Array on 100 tables", () => {
      toUint8Array(doc);
    });
    expect(writeMs).toBeLessThan(BUDGET_MS);
  });

  it("replaceText across a large body finds & rewrites in linear time", () => {
    const doc = createDocx({ paragraphs: [] });
    for (let i = 0; i < 5_000; i++) {
      appendParagraph(doc, i % 7 === 0 ? `line ${i} {{name}}` : `line ${i}`);
    }
    let count = 0;
    const elapsed = timeit("replaceText {{name}} ×5k", () => {
      count = replaceText(doc, /\{\{name\}\}/g, () => "山田太郎");
    });
    expect(count).toBeGreaterThan(0);
    expect(elapsed).toBeLessThan(BUDGET_MS);
    const t = text(doc);
    expect(t).toContain("山田太郎");
  });

  it("kitchen-sink doc validates without quadratic blowup", () => {
    const doc = createDocx({ paragraphs: [] });
    setPageSize(doc, PAGE_SIZE_A4);
    setPageMargins(doc, MARGINS_NORMAL);
    setCoreProperties(doc, { title: "Perf smoke", creator: "word-kit" });
    addHeader(doc, "Perf");
    addPageNumberFooter(doc, "Page ", "");
    appendHeading(doc, "Top", 1);
    for (let i = 0; i < 1_000; i++) appendParagraph(doc, `body ${i}`);
    addBulletList(
      doc,
      Array.from({ length: 200 }, (_, i) => `item ${i}`),
    );
    addTable(
      doc,
      Array.from({ length: 100 }, (_, r) => [`r${r}c0`, `r${r}c1`, `r${r}c2`]),
    );
    let validateMs = 0;
    const writeMs = timeit("toUint8Array (kitchen sink)", () => {
      const bytes = toUint8Array(doc);
      validateMs = timeit("validate(reopen)", () => {
        const reopened = openDocx(bytes);
        const issues = validate(reopened);
        expect(issues.filter((i) => i.level === "error")).toHaveLength(0);
      });
    });
    expect(writeMs).toBeLessThan(BUDGET_MS);
    expect(validateMs).toBeLessThan(BUDGET_MS);
  });
});
