// /llms.txt — a short, machine-readable index of word-kit's documentation
// for LLM consumers. Format follows the llmstxt.org proposal: H1 title,
// blockquote summary, then sections with linked bullets.
//
// Sibling of xlsx-kit's /llms.txt; URLs are relative so the file works
// under any base path (custom domain, user page, or project page
// /word-kit/).

import type { RequestHandler } from "./$types";

export const prerender = true;

const BODY = `# word-kit

> OOXML-compliant (ECMA-376) Word \`.docx\` library. Reads, edits, and writes WordprocessingML in both Node 20+ and modern browsers. Lossless round-trip for every element the library doesn't yet model (so hand-designed templates survive intact). Function-first API: every operation is a standalone, tree-shakeable export. Optional companion \`@word-kit/preview\` mounts a read-only render of any \`Docx\` into a DOM container.

## Packages

- [\`@word-kit/core\`](./api) — the public authoring API. \`Docx\` is a plain interface; every operation is a standalone function. Minimal slice (\`createDocx + appendParagraph + toUint8Array\`) bundles to ~42 KB minified; full surface ~131 KB.
- [\`@word-kit/preview\`](./api) — browser-side read-only preview. Single function entry \`previewToDOM(source, container, options?)\`. Wraps the OSS \`docx-preview\` renderer; wrap is intentional and final.

## Docs

- [Getting started](./docs/getting-started) — install, build a doc from scratch, open and fill a template, embed the preview.
- [Recipes](./docs/recipes) — type-checked snippets for mail-merge, PowerPoint-style designed bases, image replacement, fields/TOC, tracked changes, and the browser embed.
- [API reference](./api) — every public export grouped by area: lifecycle, paragraphs & blocks, inline & text, styles & numbering, tables, images, headers/footers/sections, comments/notes/hyperlinks/bookmarks, fields & tracked changes, document properties, diagnostics, browser preview.
- [Playground](./playground) — drop a \`.docx\` (or generate a built-in sample) and see it rendered live by \`@word-kit/preview\`. The bytes never leave the page.

## Scope

In scope: WordprocessingML read/edit/write, browser preview via the companion package, OPC packaging and DrawingML as underlying layers.

Out of scope (for now): \`.pptx\` (PresentationML) and \`.xlsx\` (SpreadsheetML) — for spreadsheets see the sibling project [xlsx-kit](https://github.com/baseballyama/xlsx-kit).

Out of scope (permanent): rendering to PDF, headless Word automation, binary \`.doc\` (pre-2007 Word).

## Status

Pre-1.0. Core (\`openDocx\` / \`createDocx\` / \`toUint8Array\`) is stable and lossless-round-trip-verified against the mammoth.js and python-docx fixture corpora. The 512-test suite runs in Node 20 / 22 / 24 on every change; tree-shake budget gate prevents accidental bundle bloat.

## Source

- [GitHub repository](https://github.com/baseballyama/word-kit)
- [Sibling project: xlsx-kit](https://github.com/baseballyama/xlsx-kit)
`;

export const GET: RequestHandler = () => {
  return new Response(BODY, {
    headers: {
      "Content-Type": "text/plain; charset=utf-8",
      "Cache-Control": "public, max-age=300",
    },
  });
};
