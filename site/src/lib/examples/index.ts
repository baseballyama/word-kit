// Registry of every example file. The ?raw imports give us the on-disk
// source verbatim. The .ts files themselves are in the project's tsconfig
// include path, so svelte-check type-checks them against the live
// @word-kit/core / @word-kit/preview surface — an API rename breaks the
// docs build before anything ships, even though the modules are never
// evaluated at runtime.

import fromScratch from "./from-scratch.ts?raw";
import templateFill from "./template-fill.ts?raw";
import previewEmbed from "./preview-embed.ts?raw";
import recipeMailMerge from "./recipe-mail-merge.ts?raw";
import recipeStyledBase from "./recipe-styled-base.ts?raw";
import recipeTrackedChanges from "./recipe-tracked-changes.ts?raw";

export type Example = {
  /** Human title for the snippet (shown above the code block). */
  title: string;
  /** Repo-relative path, also used for the file-tab caption. */
  path: string;
  /** Verbatim source text. */
  source: string;
  /** Short description used in docs. */
  description: string;
};

export const examples = {
  fromScratch: {
    title: "Build a document from scratch",
    path: "site/src/lib/examples/from-scratch.ts",
    description:
      "Heading, body paragraph, bullet list, table, A4 page size — the canonical authoring slice.",
    source: fromScratch,
  },
  templateFill: {
    title: "Open a template, fill placeholders",
    path: "site/src/lib/examples/template-fill.ts",
    description:
      "replaceTextEverywhere walks every story (body, headers, footers, footnotes, textboxes) — not just the main document.",
    source: templateFill,
  },
  previewEmbed: {
    title: "Render in the browser with @word-kit/preview",
    path: "site/src/lib/examples/preview-embed.ts",
    description:
      "previewToDOM mounts a read-only render of any Docx into a DOM container. Returns an idempotent dispose handle.",
    source: previewEmbed,
  },
  recipeMailMerge: {
    title: "Mail-merge across every story in the package",
    path: "site/src/lib/examples/recipe-mail-merge.ts",
    description:
      "replaceTextEverywhere walks body, headers, footers, footnotes, comments, and textboxes — so a placeholder in a header is filled the same way one in the body is.",
    source: recipeMailMerge,
  },
  recipeStyledBase: {
    title: "PowerPoint-style designed base",
    path: "site/src/lib/examples/recipe-styled-base.ts",
    description:
      "Open a hand-designed .docx, lift its style table into your authoring graph with mergeStylesFromTemplate, then keep appending. Custom styles applied by name via setParagraphStyle.",
    source: recipeStyledBase,
  },
  recipeTrackedChanges: {
    title: "Accept or reject tracked changes in bulk",
    path: "site/src/lib/examples/recipe-tracked-changes.ts",
    description:
      "acceptAllRevisions inlines suggested inserts and bakes in deletions. rejectAllRevisions does the inverse — useful for producing both 'final' and 'original' branches from one reviewed doc.",
    source: recipeTrackedChanges,
  },
} as const satisfies Record<string, Example>;

export type ExampleKey = keyof typeof examples;
