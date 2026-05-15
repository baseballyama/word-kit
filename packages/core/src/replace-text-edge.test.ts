// Edge-case audit for replaceText / replaceTextEverywhere — the
// most-used public mutation. The original happy-path tests cover the
// common templating scenarios; this file pokes at corners that tend to
// hide bugs:
//
//   - capture groups in the regex
//   - empty match alternatives
//   - replacement that itself matches the regex (cycle hazard)
//   - replacement returning empty string
//   - run-spanning matches across formatting boundaries
//   - multiple matches in the same paragraph
//   - case-insensitive flag
//   - replaceTextEverywhere across header / footer / comment

import { describe, expect, it } from "vitest";
import {
  addComment,
  addFooter,
  addHeader,
  appendParagraph,
  appendTextRun,
  createDocx,
  findText,
  findTextEverywhere,
  openDocx,
  paragraphs,
  paragraphText,
  replaceText,
  replaceTextEverywhere,
  text,
  toUint8Array,
} from "./index.js";

function reopen(doc: ReturnType<typeof createDocx>): ReturnType<typeof createDocx> {
  return openDocx(toUint8Array(doc));
}

describe("replaceText edge cases", () => {
  it("captures are forwarded to the replacement callback", () => {
    const doc = createDocx({ paragraphs: ["Hello, world. Hello, friend."] });
    const n = replaceText(doc, /Hello, (\w+)/g, (m) => `Hi ${m.captures[0]}`);
    expect(n).toBe(2);
    const final = reopen(doc);
    expect(text(final)).toContain("Hi world");
    expect(text(final)).toContain("Hi friend");
  });

  it("string replacement is treated literally (no $1 expansion)", () => {
    const doc = createDocx({ paragraphs: ["Hello, world"] });
    replaceText(doc, /Hello, (\w+)/g, "$1!");
    const final = reopen(doc);
    expect(text(final)).toContain("$1!");
    expect(text(final)).not.toContain("world");
  });

  it("replacement that returns empty string deletes the match", () => {
    const doc = createDocx({ paragraphs: ["[remove this]middle[remove this]end"] });
    replaceText(doc, /\[remove this\]/g, () => "");
    const final = reopen(doc);
    expect(text(final)).toBe("middleend");
  });

  it("non-global regex still replaces every occurrence (documented Word-like 'Replace All' semantics)", () => {
    // String.prototype.replace would do first-only here; replaceText
    // intentionally diverges so callers get Word's "Replace All" behaviour
    // without having to remember the /g flag. Documented on the function.
    const doc = createDocx({ paragraphs: ["aaa bbb aaa"] });
    const n = replaceText(doc, /aaa/, () => "X");
    expect(n).toBe(2);
    expect(text(reopen(doc))).toBe("X bbb X");
  });

  it("case-insensitive flag works", () => {
    const doc = createDocx({ paragraphs: ["FOO foo Foo"] });
    const n = replaceText(doc, /foo/gi, () => "X");
    expect(n).toBe(3);
    expect(text(reopen(doc))).toBe("X X X");
  });

  it("returns 0 when nothing matches and the doc isn't dirty", () => {
    const doc = createDocx({ paragraphs: ["plain"] });
    expect(replaceText(doc, /nope/g, () => "X")).toBe(0);
  });

  it("match spanning multiple runs gets resolved correctly", () => {
    const doc = createDocx({ paragraphs: [""] });
    appendParagraph(doc, "");
    const target = paragraphs(doc).at(-1)!;
    appendTextRun(target, "{{na");
    appendTextRun(target, "me}}");
    const n = replaceText(doc, /\{\{name\}\}/g, () => "Yamada");
    expect(n).toBe(1);
    expect(paragraphText(target)).toBe("Yamada");
  });

  it("multiple matches in the same paragraph all get replaced", () => {
    const doc = createDocx({ paragraphs: ["x{{a}}y{{b}}z{{c}}"] });
    const seen: string[] = [];
    replaceText(doc, /\{\{(\w+)\}\}/g, (m) => {
      seen.push(m.captures[0] ?? "");
      return m.captures[0]?.toUpperCase() ?? "";
    });
    expect(seen).toEqual(["a", "b", "c"]);
    expect(text(reopen(doc))).toBe("xAyBzC");
  });

  it("replacement that itself contains the regex pattern does not loop forever", () => {
    const doc = createDocx({ paragraphs: ["{{name}}"] });
    const n = replaceText(doc, /\{\{(\w+)\}\}/g, (m) => `{{${m.captures[0]}}}!`);
    // The new tokens contain `{{...}}` again, but replaceText must not
    // re-match its own output — a single-pass operation.
    expect(n).toBe(1);
    expect(text(reopen(doc))).toBe("{{name}}!");
  });
});

describe("findText edge cases", () => {
  it("returns the exact substring per match including captures", () => {
    const doc = createDocx({ paragraphs: ["Order #1001 + #1002"] });
    const matches = findText(doc, /#(\d+)/g);
    expect(matches.map((m) => m.text)).toEqual(["#1001", "#1002"]);
    expect(matches.map((m) => m.captures[0])).toEqual(["1001", "1002"]);
  });

  it("non-global regex returns every match (documented Word-like 'Find All' semantics)", () => {
    // Mirrors replaceText: callers don't have to remember /g.
    const doc = createDocx({ paragraphs: ["Order #1001 + #1002"] });
    const matches = findText(doc, /#(\d+)/);
    expect(matches).toHaveLength(2);
    expect(matches.map((m) => m.text)).toEqual(["#1001", "#1002"]);
  });

  it("returns an empty array when nothing matches", () => {
    const doc = createDocx({ paragraphs: ["plain"] });
    expect(findText(doc, /nope/g)).toEqual([]);
  });
});

describe("replaceTextEverywhere edge cases", () => {
  it("replaces in body + header + footer + comment (one pass)", () => {
    const doc = createDocx({ paragraphs: ["body has {{x}}"] });
    addHeader(doc, "header has {{x}}");
    addFooter(doc, "footer has {{x}}");
    addComment(doc, paragraphs(doc)[0]!, {
      author: "R",
      initials: "R",
      text: "comment has {{x}}",
    });
    const n = replaceTextEverywhere(doc, /\{\{x\}\}/g, () => "Y");
    expect(n).toBeGreaterThanOrEqual(4);
    const final = reopen(doc);
    expect(text(final)).toContain("body has Y");
    // Header / footer / comment XML byte-check.
    const xml = (path: string) =>
      new TextDecoder().decode(final.opc.parts.get(path)?.data ?? new Uint8Array());
    expect(xml("/word/header1.xml")).toContain("Y");
    expect(xml("/word/footer1.xml")).toContain("Y");
    expect(xml("/word/comments.xml")).toContain("Y");
  });

  it("findTextEverywhere returns one entry per part with its matches", () => {
    const doc = createDocx({ paragraphs: ["body has X"] });
    addHeader(doc, "header has X");
    const groups = findTextEverywhere(doc, /X/g);
    const partNames = groups.map((g) => g.partName);
    expect(partNames).toContain("/word/document.xml");
    expect(partNames.some((n) => n.startsWith("/word/header"))).toBe(true);
  });
});
