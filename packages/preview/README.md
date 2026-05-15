# @word-kit/preview

Browser-side read-only preview for `.docx` documents produced (or
opened) by [`@word-kit/core`](../core).

## Status

**Stable.** Wraps the OSS
[`docx-preview`](https://github.com/VolodymyrBaydalka/docxjs)
(Apache-2.0) renderer behind `word-kit`'s function-API surface. The
wrap is intentional and final — see
[`docs/PLAN-PREVIEW.md`](../../docs/PLAN-PREVIEW.md) for why we are
not re-implementing the renderer in-house.

If you hit a fidelity bug, please file it against
[`docx-preview`](https://github.com/VolodymyrBaydalka/docxjs/issues)
upstream first; we follow that project's releases and bump the
peer-dep promptly.

## Install

```bash
npm install @word-kit/core @word-kit/preview
# or: pnpm add @word-kit/core @word-kit/preview
```

`@word-kit/preview` declares `@word-kit/core` and `docx-preview` as
runtime dependencies. Both are ESM-only.

## Usage

```ts
import { openDocx } from "@word-kit/core";
import { previewToDOM } from "@word-kit/preview";

const bytes = /* Uint8Array | Blob | ArrayBuffer */;
const doc = openDocx(bytes);

const handle = await previewToDOM(
  doc,
  document.getElementById("preview")!,
  {
    classPrefix: "wk-",          // CSS namespace (default "wk-")
    inWrapper: true,             // wrap each page in a scrolling block
    breakPages: true,            // honour source-declared page breaks
    renderFonts: true,           // inject @font-face from the docx
    experimentalComments: false, // upstream-experimental
    experimentalChanges: false,  // upstream-experimental
  },
);

// …later, when you no longer need the preview:
handle.dispose();
```

`previewToDOM` accepts:

- a `Docx` value returned by `createDocx` / `openDocx` /
  `clone` / `fromBlob` from `@word-kit/core`, **or**
- raw bytes — `Uint8Array`, `Blob`, or `ArrayBuffer`.

It returns a `Handle` with a single `dispose()` method that detaches
the rendered DOM and releases internal references. `dispose()` is
idempotent.

## What renders well

- Body text + paragraph styling (bold, italic, underline, color,
  font, size).
- Custom paragraph / character / table styles defined in
  `styles.xml`. Custom styles your code sets through
  `addStyle(doc, …)` work the same as styles imported from a hand-
  designed Word template.
- Lists (bullet + numbered).
- Tables — borders, shading, header rows.
- Inline images.
- Hyperlinks.
- Headers / footers — drawn at the top / bottom of each page.
- Footnotes / endnotes — flat tail block.

## Known gaps (inherited from `docx-preview` v0)

- **No live re-pagination.** Pages are drawn at source-declared
  breaks (`<w:br w:type="page"/>`, `<w:lastRenderedPageBreak/>`).
  Edits to the doc don't reflow page boundaries.
- **No field evaluation.** TOC, PAGE, NUMPAGES etc. show their
  cached display values when present, otherwise the field
  instruction.
- Tab stops / list edge cases / empty-paragraph height — known
  upstream issues.

The v1 native renderer will address the long tail one feature area
at a time; see `docs/PLAN-PREVIEW.md` for the order.

## Out of scope

- **Editing.** Read-only by design.
- **Pixel-perfect Word rendering.** HTML/CSS cannot replicate every
  WML page semantic.
- **PDF rendering.** Document the LibreOffice headless → `pdf.js`
  recipe externally; we don't ship it.

## License

[MIT](../../LICENSE)
