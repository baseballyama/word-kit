/**
 * The completeness-guarantee test.
 *
 * This is the enforcement half of the capability ledger: it fails the build if
 * any WordprocessingML element slips through unclassified, if an `edit` element
 * points at a command that no longer exists, or if the element universe drifts
 * from the committed snapshot without being regenerated. It also (re)writes
 * COVERAGE.md so the scoreboard stays current.
 */

import { existsSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { COMMAND_IDS } from "../commands/registry.js";
import universe from "./element-universe.json" with { type: "json" };
import { classify, coverageSummary, LEDGER } from "./ledger.js";

const here = dirname(fileURLToPath(import.meta.url));

interface UniverseElement {
  element: string;
  name: string;
  prefix: string;
  domain: string;
}
const ELEMENTS = (universe as { elements: UniverseElement[]; totalElements: number }).elements;

describe("capability ledger completeness", () => {
  it("classifies every element in the schema universe", () => {
    const missing: string[] = [];
    for (const e of ELEMENTS) {
      const cap = classify(e.element, e.domain);
      if (cap.notes?.startsWith("UNCLASSIFIED")) missing.push(e.element);
    }
    expect(missing, `unclassified elements:\n${missing.join("\n")}`).toEqual([]);
  });

  it("has zero unclassified elements in the summary", () => {
    expect(coverageSummary().unclassified).toBe(0);
  });

  it("covers exactly the committed universe (snapshot not drifted)", () => {
    expect(LEDGER.length).toBe(ELEMENTS.length);
    expect(LEDGER.length).toBe((universe as { totalElements: number }).totalElements);
  });

  it("every `edit` element names a command that exists in the registry", () => {
    const s = coverageSummary();
    expect(
      s.danglingCommands,
      `edit elements referencing missing commands:\n${s.danglingCommands.join("\n")}`,
    ).toEqual([]);
  });

  it("every command id referenced by the ledger is unique-resolvable", () => {
    for (const cap of LEDGER) {
      if (cap.disposition === "edit" && cap.command) {
        expect(COMMAND_IDS.has(cap.command)).toBe(true);
      }
    }
  });

  it("has meaningful edit coverage of the WordprocessingML core", () => {
    // Guardrail: the wml domain must retain a real editable surface so a
    // regression that silently strips commands is caught.
    const wml = coverageSummary().byDomain.wml;
    expect(wml).toBeDefined();
    expect(wml!.edit).toBeGreaterThanOrEqual(50);
  });

  it("writes COVERAGE.md", async () => {
    const { renderCoverageMarkdown } = await import("./report.js");
    const md = renderCoverageMarkdown();
    const out = join(here, "..", "..", "COVERAGE.md");
    // Only rewrite when changed to keep the file stable in git diffs.
    const prev = existsSync(out) ? readFileSync(out, "utf8") : "";
    if (prev !== md) writeFileSync(out, md);
    expect(md).toContain("# Editor capability coverage");
  });
});
