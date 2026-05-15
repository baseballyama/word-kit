# Changelog

The format is loosely based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/)
and the project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).
Each user-visible change is also tracked via [Changesets](https://github.com/changesets/changesets);
this file is a hand-curated overview.

## Unreleased

### Changed

- **No-classes API.** Every package now exposes plain data types and
  standalone functions instead of classes. The motivation is
  tree-shaking: classes carry every method along with the prototype,
  so a bundler can't drop unused operations once an instance escapes.
  - `Docx` is an `interface` (just `{ opc, document, partName, … }`).
  - `Docx.create(…)` → `createDocx(…)`. `Docx.open(bytes)` → `openDocx(bytes)`.
    `Docx.fromBlob(blob)` → `fromBlob(blob)`.
  - `doc.appendParagraph(text)` → `appendParagraph(doc, text)`. Same
    pattern for every previous method.
  - Previous getters become functions: `doc.paragraphs` → `paragraphs(doc)`,
    `doc.statistics` → `statistics(doc)`, etc.
  - Previous setters become `setX` functions: `doc.title = x` →
    `setTitle(doc, x)`. `doc.opc`, `doc.document`, `doc.partName` stay
    as direct property access.
  - `OpcPackage`, `ContentTypesIndex`, `RelationshipSet` are likewise
    plain interfaces backed by standalone `addPart`, `getPart`,
    `allRelationships`, `packageRelationships`, … functions.
  - `XmlParseError` is no longer a class. Use `XmlParseError.is(e)` for
    narrowing (in place of `e instanceof XmlParseError`).

  A minimal `createDocx + appendParagraph + toUint8Array` slice bundles
  to ~40 KB minified; the full surface is ~115 KB. CI enforces both the
  byte budget and that no feature-specific string literals from unused
  branches leak into the minimal bundle (see `scripts/check-tree-shake.mjs`).

### Added

- `@word-kit/opc` — Open Packaging Conventions reader/writer with byte-stable
  round-trip for untouched parts. Built on `fflate`. Plain-data API
  (`OpcPackage` interface + `readOpcPackage`, `writeOpcPackage`,
  `addPart`, `getPart`, `allRelationships`, `packageRelationships`, …).
- `@word-kit/ooxml-xml` — namespace-aware XML parser/serializer that preserves
  attribute order, prefixes, `xml:space="preserve"`, CDATA, comments, and PIs.
- `@word-kit/wml` — WordprocessingML AST plus parsers, writers, and builders
  for paragraphs, runs, tables, styles, numbering, headers/footers, sections,
  comments, footnotes/endnotes, hyperlinks, fields, and document properties.
- `@word-kit/core` — the public `Docx` interface and standalone-function API
  wrapping the lower packages.
- `@word-kit/preview` — browser-side read-only preview. Single function entry,
  `previewToDOM(source, container, options?) → Promise<Handle>`. v0
  implementation wraps the OSS `docx-preview` (Apache-2.0). The wrap is
  intentional and final; see `docs/PLAN-PREVIEW.md` for the rationale.

#### Authoring (function API on `@word-kit/core`)

- Lifecycle: `createDocx({ paragraphs? })`, `openDocx(bytes)`,
  `fromBlob(blob)`, `toUint8Array(doc)`, `toBlob(doc)`, `clone(doc)`.
- Paragraphs: `appendParagraph`, `insertParagraphAt`, `removeParagraph`,
  `appendHeading`, `appendPageBreak`, `appendLineBreak`,
  `appendSectionBreak`, `clearBody`.
- Inline / text: `replaceText`, `replaceTextEverywhere`, `findText`,
  `findTextEverywhere`, `appendTextRun`, `setParagraphText`,
  `paragraphText`, `setRunFormat`, `clearRunFormat`, `getRunFormat`,
  `setParagraphAlignment`, `setParagraphIndent`, `setParagraphSpacing`.
- Styles + numbering: `addStyle`, `removeStyle`, `listStyles`,
  `ensureHeadingStyles`, `addBulletList`, `addNumberedList`,
  `applyListToParagraph`.
- Tables: `addTable`, `tables`, `removeTable`, `removeAllTables`,
  `unwrapTable`.
- Images: `addImage`, `addImageRun`, `insertImageInto`, `images`,
  `replaceImage`, `removeAllImages`.
- Headers / footers / sections: `addHeader`, `addFooter`,
  `addPageNumberFooter`, `setPageSize`, `setPageMargins`,
  `setPageOrientation`, `headers`, `footers`,
  `removeAllHeaders`, `removeAllFooters`.
- Comments / footnotes / endnotes: `addComment`, `addFootnote`,
  `addEndnote`, `removeAllComments`, `removeAllFootnotes`,
  `removeAllEndnotes`.
- Hyperlinks + bookmarks: `addHyperlink`, `addInternalHyperlink`,
  `externalHyperlinks`, `setHyperlinkUrl`, `removeAllHyperlinks`,
  `addBookmark`, `removeBookmark`, `removeAllBookmarks`, `bookmarks`.
- Fields: `appendField`, `addTableOfContents`, `appendMergeField`,
  `WORD_FIELDS`, plus the page-number footer helper above.
- Tracked changes: `acceptAllRevisions`, `rejectAllRevisions`.
- Core / app properties: `coreProperties`, `setCoreProperties`,
  `appProperties`, `setAppProperties`, `title`, `author`,
  `setTitle`, `setAuthor`.
- Templates (PowerPoint-style "open a designed base, append content"):
  `mergeStylesFromTemplate`, `findStyleIdByName`,
  `setParagraphStyle`, `imageReferences`, `replaceImageByAltText`.
- Diagnostics: `validate(doc)`, `statistics(doc)`, `outline(doc)`,
  `fields(doc)`.

#### Browser preview (function API on `@word-kit/preview`)

- `previewToDOM(source, container, options?)` renders a `Docx`,
  `Uint8Array`, `Blob`, or `ArrayBuffer` into a DOM container.
  Returns an idempotent `Handle.dispose()` for teardown.
- Options: `classPrefix` (default `"wk-"`), `inWrapper`,
  `breakPages`, `renderFonts`, `experimentalComments`,
  `experimentalChanges`. All overridable.

### Build + tooling

- `tsdown` (rolldown-based) replaces `tsup` for the bundling step.
  Output extensions changed from `.js`/`.d.ts` to `.mjs`/`.d.mts`.
- `pnpm test` now runs `pnpm build` first via the npm-standard
  `pretest` hook so cross-package imports always resolve to fresh
  dist.
- New CI gate: `pnpm check:tree-shake` budgets a minimal
  `createDocx + appendParagraph + toUint8Array` bundle (~42 KB
  minified) against the full surface (~131 KB).
- `pnpm sample` writes 32 demonstration `.docx` files into
  `./samples/` for manual verification in Microsoft Word.
- New perf-smoke test catches accidental quadratic regressions:
  10k-paragraph round-trip in 220 ms locally; budget is 10 s per
  block to leave room for slow CI.

### Implementation notes

- Lossless round-trip: every XML element the library does not yet structure
  is preserved as a `WmlRawBlock`/`WmlRawInline` pass-through with its
  original child position, so re-saving an unmodified template leaves
  Word's "needs repair" prompt out of the picture.
- Verified against the mammoth.js fixture corpus (comments, footnotes,
  endnotes, tables, images, hyperlinks, text boxes, UTF-8 BOM, lists)
  and the python-docx test corpus. The ISO/IEC 29500 Strict variant is
  explicitly out of scope today and remains pass-through only.
- 311 tests, all running in vitest under Node ≥ 20; the public surface
  also runs in modern browsers (no Node-only dependencies in the published
  bundles).
- CI gate runs typecheck, lint, format check, tests across Node 20/22/24,
  AND the tree-shake budget check.
