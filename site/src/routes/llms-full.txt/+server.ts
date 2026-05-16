// /llms-full.txt — every doc page concatenated into a single Markdown
// document. Companion of /llms.txt (which is just an index): this is the
// payload, intended for LLM ingestion in one fetch.
//
// Sibling of xlsx-kit's /llms-full.txt. xlsx-kit assembles its body by
// glob-loading mdsvex `.svx` files; word-kit's docs pages are
// hand-written `.svelte`, so we render the same content from two
// already-existing sources of truth instead:
//
//   * `$lib/examples` — the type-checked example files used by the home
//     page and the recipes page.
//   * `$lib/api-groups` — the curated grouped-export list used by the
//     `/api` page and the `check:api-page` CI gate.
//
// That keeps the LLM payload aligned with what the site actually
// renders, with no second copy of the content to drift.

import { examples } from "$lib/examples";
import { apiGroups, apiTotalCount } from "$lib/api-groups";
import type { RequestHandler } from "./$types";

export const prerender = true;

const PREAMBLE = `# word-kit — full documentation

This file is the concatenation of every page on the word-kit docs site, intended for LLM ingestion in a single fetch. Page boundaries are marked with H1 headings prefixed by the source path. The companion index at \`/llms.txt\` lists the same pages with one-line descriptions.

Source repo: https://github.com/baseballyama/word-kit
`;

const PROJECT_OVERVIEW = `OOXML-compliant (ECMA-376 Part 1 — WordprocessingML) Word \`.docx\` library for Node 22+ and modern browsers. Function-first API: every operation is a standalone, tree-shakeable export, never a method on a class. Lossless round-trip for every element the library does not yet model, so hand-designed templates survive intact.

Two packages on npm:

- \`@word-kit/core\` — the public authoring API.
- \`@word-kit/preview\` — optional companion that mounts a read-only preview of any \`Docx\` (or raw bytes) into a DOM container, by wrapping the OSS \`docx-preview\` renderer.

Bundle budgets (CI-enforced): minimal \`createDocx + appendParagraph + toUint8Array\` slice ~42 KB minified; full surface ~131 KB.

Test gate: 512 vitest tests on Node 22 and 24 every change. Real-Word fixture corpora (mammoth.js + python-docx) round-trip-verified.
`;

const GETTING_STARTED = `Install both packages:

\`\`\`sh
pnpm add @word-kit/core @word-kit/preview
\`\`\`

Both ship as ESM with bundled \`.d.ts\` types. Neither has Node-only dependencies.

## Build a document from scratch

The hello-world. \`createDocx\` returns a plain \`Docx\` object that the rest of the API treats as a value.

\`\`\`ts title="${examples.fromScratch.path}"
${examples.fromScratch.source.trimEnd()}
\`\`\`

## Open a template, fill placeholders

Existing \`.docx\` files are opened with \`openDocx\` and edited in place. word-kit preserves every XML element it does not yet model as a pass-through node, so re-saving an unmodified template doesn't trip Word's "needs repair" prompt.

\`replaceTextEverywhere\` walks every story — body, headers, footers, footnotes, endnotes, comments, textboxes — not just the main document. Run-spanning matches like \`{{name}}\` split across multiple runs are joined before the regex sees them.

\`\`\`ts title="${examples.templateFill.path}"
${examples.templateFill.source.trimEnd()}
\`\`\`

## Render in the browser

The companion package \`@word-kit/preview\` mounts a read-only preview of any \`Docx\` into a DOM container. It wraps the OSS \`docx-preview\` renderer behind a stable function-API entry point.

\`\`\`ts title="${examples.previewEmbed.path}"
${examples.previewEmbed.source.trimEnd()}
\`\`\`
`;

function buildRecipesMarkdown(): string {
  const intro = `Type-checked snippets for common scenarios. Every snippet below lives under \`site/src/lib/examples/\` and is type-checked by \`svelte-check\` against the live \`@word-kit/core\` / \`@word-kit/preview\` surface — an API rename breaks the docs build before anything ships.

Pointers under each snippet name the matching sample file produced by \`pnpm sample\` (or the integration test that exercises the same path).
`;

  const recipeKeys = [
    "recipeMailMerge",
    "recipeStyledBase",
    "recipeTrackedChanges",
    "previewEmbed",
  ] as const;

  const pointers: Record<string, string> = {
    recipeMailMerge: "samples/20-mailmerge-text-replace-*.docx (run `pnpm sample`)",
    recipeStyledBase: "samples/30-styled-base-*.docx",
    recipeTrackedChanges: "samples/09-tracked-changes.docx",
    previewEmbed: "see live in the /playground page",
  };

  const sections = recipeKeys.map((key, i) => {
    const ex = examples[key];
    const pointer = pointers[key] ?? "—";
    return [
      `### ${String(i + 1).padStart(2, "0")} · ${ex.title}`,
      "",
      ex.description,
      "",
      `*Where:* \`${pointer}\``,
      "",
      `\`\`\`ts title="${ex.path}"`,
      ex.source.trimEnd(),
      "```",
    ].join("\n");
  });

  return `${intro}\n${sections.join("\n\n")}\n`;
}

function buildApiMarkdown(): string {
  const intro = `\`@word-kit/core\` exposes ${apiTotalCount} standalone functions and constants, grouped by area below. \`previewToDOM\` lives in the companion \`@word-kit/preview\` package; everything else lives in \`@word-kit/core\`. For full signatures, read \`packages/core/src/index.ts\` in the repo.
`;

  const sections = apiGroups.map((g) => {
    const items = g.entries.map((e) => {
      const sig = e.sig ? `: \`${e.sig}\`` : "";
      return `- \`${e.name}\`${sig}`;
    });
    return `### ${g.num} · ${g.title}\n\n${items.join("\n")}`;
  });

  return `${intro}\n${sections.join("\n\n")}\n`;
}

const PLAYGROUND = `An interactive page at \`/playground\`. Drop a \`.docx\` onto the canvas (or click "Generate sample" to build one in the browser via \`@word-kit/core\`) and see it rendered live by \`@word-kit/preview\`. The bytes never leave the browser.

The "Download current bytes" button hands back whatever document is mounted right now — useful for shipping an edited workbook back out of the page after a transform.
`;

const SCOPE = `## In scope

WordprocessingML (\`.docx\`) — read, edit, write. OPC packaging (ECMA-376 Part 2) and DrawingML are part of the OOXML stack and live in their own packages, but exist to serve docx. Browser preview lives in the companion \`@word-kit/preview\`.

## Out of scope (for now)

\`.pptx\` (PresentationML) and \`.xlsx\` (SpreadsheetML). For spreadsheets, see the sibling project [xlsx-kit](https://github.com/baseballyama/xlsx-kit).

## Out of scope (permanent)

Rendering to PDF, headless Word automation, binary \`.doc\` (pre-2007 Word).
`;

function buildBody(): string {
  const parts = [
    PREAMBLE,
    "\n---\n\n<!-- Page: / -->\n# word-kit — overview\n\n" + PROJECT_OVERVIEW,
    "\n---\n\n<!-- Page: /docs/getting-started -->\n# Getting started\n\n" + GETTING_STARTED,
    "\n---\n\n<!-- Page: /docs/recipes -->\n# Recipes\n\n" + buildRecipesMarkdown(),
    "\n---\n\n<!-- Page: /api -->\n# API reference\n\n" + buildApiMarkdown(),
    "\n---\n\n<!-- Page: /playground -->\n# Playground\n\n" + PLAYGROUND,
    "\n---\n\n<!-- Page: scope -->\n# Scope\n\n" + SCOPE,
  ];
  return parts.join("\n");
}

export const GET: RequestHandler = () => {
  return new Response(buildBody(), {
    headers: {
      "Content-Type": "text/plain; charset=utf-8",
      "Cache-Control": "public, max-age=300",
    },
  });
};
