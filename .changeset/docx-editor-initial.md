---
"@office-kit/docx-editor": minor
"@office-kit/docx": patch
---

feat: add `@office-kit/docx-editor` — an MS Office-like WYSIWYG editing core

New package that turns `@office-kit/docx` into a Word-like editor: an
`EditorModel` with undo/redo, a command layer (text/paragraph/structure/list/
table/image/style/section/header-footer/references/review) where every mutation
routes through the `@office-kit/docx` public API, an AST→HTML canvas renderer,
and DOM-selection mapping. A SvelteKit ribbon UI ships at the site's `/editor`
route.

The editing surface is interactive: Enter splits the paragraph at the caret and
Backspace/Delete merge across paragraph boundaries (`splitParagraphCommand` /
`mergeBackCommand`, backed by the library's `splitParagraphAt` /
`mergeParagraphIntoPrevious`); the caret is preserved across structural
re-renders; Ctrl+Z/Y drive undo/redo; paste inserts plain text (multi-line →
paragraphs) keeping the AST in sync; and the UI adds live word/character count,
find & replace, zoom, caret-synced font/size/style, and heading-styled rendering
so the document reads with real visual hierarchy.

Character formatting is now selection-precise: bolding (or any run format on) a
partial selection splits runs at the selection boundaries and formats only the
selected characters instead of the whole line — backed by the library's
`isolateParagraphRunRange` / `runTextLength`. The ribbon UI is localized into
six languages (English, 日本語, Español, Français, Deutsch, 中文) with a
language switcher.

Completeness against the WordprocessingML spec is enforced by a capability
ledger: every element in the ECMA-376 schema universe is classified as
`edit` / `render` / `preserve`, and a test fails the build if any element is
unclassified or an editable element's command goes missing.

The library gains generic property setters so the editor can reach the long
tail of WordprocessingML formatting without a bespoke function per element:

- run / paragraph: `setRunOnOff` / `setRunValProp` / `getRunProp` and the
  `setParagraphOnOff` / `setParagraphValProp` / `getParagraphProp` equivalents;
- any properties container: `setElementOnOff` / `setElementValProp` /
  `getElementProp` / `makePropsElement` (for `<w:tcPr>` / `<w:trPr>` /
  `<w:tblPr>` / `<w:sectPr>`);
- document settings: `setDocumentSettingOnOff` / `setDocumentSettingVal` /
  `getDocumentSetting` (creates `word/settings.xml` on demand);
- style definitions: `setStyleOnOff` / `setStyleValProp` / `getStyleProp`;
- numbering levels: `setNumberingLevelVal` / `setNumberingLevelOnOff` /
  `getNumberingLevelProp`;
- images: `imageDrawings` / `setImageSizeEmu` / `setImageAltText` /
  `getImageInfo` (resize and alt-text existing pictures);
- raw XML nodes: `setElementAttr` / `getElementAttr` / `childElementsOf` /
  `appendChildElement` — the universal escape hatch that lets the editor edit
  any element (DrawingML shape geometry, OMML math, legacy VML) by attribute or
  child, since these live inline in `document.xml` and flush through the
  document AST on save;
- arbitrary XML parts: `xmlPartNames` / `getRawPartRoot` / `markRawPartDirty` —
  open, edit, and re-serialize any XML part (fontTable, settings, styles,
  numbering, comments, foot-/endnotes, headers, footers, webSettings, docProps),
  so every OOXML element in the whole package is reachable and editable;
- caret-level structure: `splitParagraphAt` / `mergeParagraphIntoPrevious` —
  split a paragraph at a run/offset and merge a paragraph into the previous one
  (the primitives behind Enter and Backspace-at-start);
- selection formatting: `isolateParagraphRunRange` / `runTextLength` — isolate a
  character range into its own run(s) so formatting applies to exactly the
  selected text.

Also newly exported from `@office-kit/docx` (they are part of the public
surface as parameter/return/AST types): `BuildStyleOptions`, `BuildTableOptions`,
`DocumentCoreProperties`, `DocumentAppProperties`, and the raw XML AST types
`XmlElement` / `XmlNode` / `XmlAttr` / `QName`.
