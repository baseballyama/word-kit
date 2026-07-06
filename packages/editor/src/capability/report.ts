/**
 * Render the capability ledger as a Markdown coverage report (COVERAGE.md).
 * The report is the human-facing scoreboard: how much of the WordprocessingML
 * universe the editor can edit / render / preserve, per schema domain.
 */

import { coverageSummary, LEDGER } from "./ledger.js";

function pct(n: number, total: number): string {
  return total === 0 ? "0%" : `${((n / total) * 100).toFixed(1)}%`;
}

/** Build the full COVERAGE.md contents. */
export function renderCoverageMarkdown(): string {
  const s = coverageSummary();
  const lines: string[] = [];
  lines.push("# Editor capability coverage");
  lines.push("");
  lines.push("> Generated from the capability ledger (`src/capability/ledger.ts`) against the");
  lines.push(
    "> ECMA-376 schema universe. Regenerate with the coverage test (it writes this file).",
  );
  lines.push("");
  lines.push(`- **Total elements**: ${s.total}`);
  lines.push(`- **Editable** (\`edit\`): ${s.edit} (${pct(s.edit, s.total)})`);
  lines.push(`- **Rendered** (\`render\`): ${s.render} (${pct(s.render, s.total)})`);
  lines.push(`- **Preserved** (\`preserve\`): ${s.preserve} (${pct(s.preserve, s.total)})`);
  lines.push(`- **Unclassified**: ${s.unclassified} (must be 0)`);
  lines.push("");
  lines.push("Every element is round-tripped losslessly regardless of disposition;");
  lines.push('`preserve` means "not yet surfaced in the UI", not "dropped".');
  lines.push("");
  lines.push("## By schema domain");
  lines.push("");
  lines.push("| Domain | Editable | Rendered | Preserved | Total |");
  lines.push("| ------ | -------- | -------- | --------- | ----- |");
  for (const [domain, d] of Object.entries(s.byDomain).toSorted()) {
    const total = d.edit + d.render + d.preserve;
    lines.push(`| ${domain} | ${d.edit} | ${d.render} | ${d.preserve} | ${total} |`);
  }
  lines.push("");
  lines.push("## Editable elements");
  lines.push("");
  lines.push("| Element | Command |");
  lines.push("| ------- | ------- |");
  for (const cap of LEDGER.filter((c) => c.disposition === "edit").toSorted((a, b) =>
    a.element.localeCompare(b.element),
  )) {
    lines.push(`| \`${cap.element}\` | \`${cap.command}\` |`);
  }
  lines.push("");
  return lines.join("\n");
}
