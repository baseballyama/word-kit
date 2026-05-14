# Changelog

The format is loosely based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/)
and the project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).
Each user-visible change is also tracked via [Changesets](https://github.com/changesets/changesets);
this file is a hand-curated overview.

## Unreleased

### Added

- `@word-kit/opc` — Open Packaging Conventions reader/writer with byte-stable
  round-trip for untouched parts. Built on `fflate`.
- `@word-kit/ooxml-xml` — namespace-aware XML parser/serializer that preserves
  attribute order, prefixes, `xml:space="preserve"`, CDATA, comments, and PIs.
- `@word-kit/wml` — WordprocessingML AST plus parsers, writers, and builders
  for paragraphs, runs, tables, styles, numbering, headers/footers, sections,
  comments, footnotes/endnotes, hyperlinks, fields, and document properties.
- `@word-kit/core` — the public `Docx` class wrapping the lower packages.

#### Authoring (`Docx`)

- `Docx.create({ paragraphs? })`, `Docx.open(bytes)`, `Docx.fromBlob(blob)`,
  `toUint8Array()`, `toBlob()`, `clone()`.
- Paragraphs: `appendParagraph`, `insertParagraphAt`, `removeParagraph`,
  `appendHeading`, `appendPageBreak`, `appendSectionBreak`, `clearBody`.
- Inline / text: `replaceText`, `replaceTextEverywhere`, `findText`,
  `findTextEverywhere`, `appendTextRun`, `setParagraphText`,
  `setParagraphAlignment`, `setParagraphIndent`, `setParagraphSpacing`.
- Styles + numbering: `addStyle`, `ensureHeadingStyles`, `addBulletList`,
  `addNumberedList`.
- Tables: `addTable`, `tables`.
- Images: `addImage`, `addImageRun`, `images`, `replaceImage`.
- Headers / footers / sections: `addHeader`, `addFooter`,
  `addPageNumberFooter`, `setPageSize`, `setPageMargins`,
  `setPageOrientation`, `headers`, `footers`.
- Comments / footnotes / endnotes: `addComment`, `addFootnote`,
  `addEndnote`, `removeAllComments`.
- Hyperlinks + bookmarks: `addHyperlink`, `addInternalHyperlink`,
  `addBookmark`, `removeBookmark`, `bookmarks`.
- Fields: `appendField`, `WORD_FIELDS`, plus the page-number footer
  helper above.
- Tracked changes: `acceptAllRevisions`, `rejectAllRevisions`.
- Core properties: `coreProperties`, `setCoreProperties`, `title`, `author`.
- Diagnostics: `validate()`, `statistics`, `outline()`, `fields` (enumeration).

### Implementation notes

- Lossless round-trip: every XML element the library does not yet structure
  is preserved as a `WmlRawBlock`/`WmlRawInline` pass-through with its
  original child position, so re-saving an unmodified template leaves
  Word's "needs repair" prompt out of the picture.
- Verified against the mammoth.js fixture corpus (comments, footnotes,
  endnotes, tables, images, hyperlinks, text boxes, UTF-8 BOM, lists).
  The ISO/IEC 29500 Strict variant is explicitly out of scope today and
  remains pass-through only.
- 200+ tests, all running in vitest under Node ≥ 20; the public surface
  also runs in modern browsers (no Node-only dependencies in the published
  bundles).
