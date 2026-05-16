// Single source of truth for the public API listing.
//
// Consumed by:
//   - `/api` — renders the groups as cards on the docs site.
//   - `/llms-full.txt` — flattens the same groups into the API section
//     of the LLM-readable concatenation.
//   - `scripts/check-api-page.mjs` — diffs the entries here against the
//     live `@word-kit/core` exports (plus a `previewToDOM` whitelist for
//     `@word-kit/preview`). CI fails if a new export isn't added here.

export type ApiEntry = { name: string; sig?: string };
export type ApiGroup = { num: string; title: string; entries: ApiEntry[] };

export const apiGroups: ApiGroup[] = [
  {
    num: "01",
    title: "Lifecycle",
    entries: [
      { name: "createDocx", sig: "({ paragraphs? }?) => Docx" },
      { name: "openDocx", sig: "(bytes: Uint8Array) => Docx" },
      { name: "fromBlob", sig: "(blob: Blob) => Promise<Docx>" },
      { name: "toUint8Array", sig: "(doc: Docx) => Uint8Array" },
      { name: "toBlob", sig: "(doc: Docx) => Blob" },
      { name: "clone", sig: "(doc: Docx) => Docx" },
    ],
  },
  {
    num: "02",
    title: "Paragraphs & blocks",
    entries: [
      { name: "appendParagraph" },
      { name: "insertParagraphAt" },
      { name: "removeParagraph" },
      { name: "appendHeading" },
      { name: "appendPageBreak" },
      { name: "appendLineBreak" },
      { name: "appendSectionBreak" },
      { name: "clearBody" },
      { name: "paragraphs" },
    ],
  },
  {
    num: "03",
    title: "Inline & text",
    entries: [
      { name: "replaceText" },
      { name: "replaceTextEverywhere" },
      { name: "findText" },
      { name: "findTextEverywhere" },
      { name: "appendTextRun" },
      { name: "setParagraphText" },
      { name: "paragraphText" },
      { name: "setRunFormat" },
      { name: "clearRunFormat" },
      { name: "getRunFormat" },
      { name: "setParagraphAlignment" },
      { name: "getParagraphAlignment" },
      { name: "setParagraphIndent" },
      { name: "setParagraphSpacing" },
      { name: "setParagraphBorders" },
      { name: "setParagraphShading" },
      { name: "getParagraphStyle" },
      { name: "getParagraphNumbering" },
      { name: "mergeAdjacentRuns" },
      { name: "mergeAdjacentRunsInBody" },
    ],
  },
  {
    num: "04",
    title: "Styles & numbering",
    entries: [
      { name: "addStyle" },
      { name: "removeStyle" },
      { name: "listStyles" },
      { name: "ensureHeadingStyles" },
      { name: "findStyleIdByName" },
      { name: "setParagraphStyle" },
      { name: "addBulletList" },
      { name: "addNumberedList" },
      { name: "applyListToParagraph" },
      { name: "mergeStylesFromTemplate" },
    ],
  },
  {
    num: "05",
    title: "Tables",
    entries: [
      { name: "addTable" },
      { name: "tables" },
      { name: "removeTable" },
      { name: "removeAllTables" },
      { name: "unwrapTable" },
      { name: "appendTableRow" },
      { name: "removeTableRow" },
      { name: "setTableRowAsHeader" },
      { name: "setTableRowHeight" },
      { name: "setTableBorders" },
      { name: "setTableCellText" },
      { name: "getTableCellText" },
      { name: "setTableCellShading" },
      { name: "setTableCellVerticalAlign" },
    ],
  },
  {
    num: "06",
    title: "Images",
    entries: [
      { name: "addImage" },
      { name: "addImageRun" },
      { name: "insertImageInto" },
      { name: "images" },
      { name: "imageReferences" },
      { name: "replaceImage" },
      { name: "replaceImageByAltText" },
      { name: "removeAllImages" },
    ],
  },
  {
    num: "07",
    title: "Headers, footers, sections",
    entries: [
      { name: "addHeader" },
      { name: "addFooter" },
      { name: "addPageNumberFooter" },
      { name: "setPageSize" },
      { name: "setPageMargins" },
      { name: "setPageOrientation" },
      { name: "headers" },
      { name: "footers" },
      { name: "removeAllHeaders" },
      { name: "removeAllFooters" },
    ],
  },
  {
    num: "08",
    title: "Comments, notes, hyperlinks, bookmarks",
    entries: [
      { name: "addComment" },
      { name: "addFootnote" },
      { name: "addEndnote" },
      { name: "removeAllComments" },
      { name: "removeAllFootnotes" },
      { name: "removeAllEndnotes" },
      { name: "addHyperlink" },
      { name: "addInternalHyperlink" },
      { name: "externalHyperlinks" },
      { name: "setHyperlinkUrl" },
      { name: "removeAllHyperlinks" },
      { name: "addBookmark" },
      { name: "removeBookmark" },
      { name: "removeAllBookmarks" },
      { name: "bookmarks" },
    ],
  },
  {
    num: "09",
    title: "Fields & tracked changes",
    entries: [
      { name: "appendField" },
      { name: "addTableOfContents" },
      { name: "appendMergeField" },
      { name: "acceptAllRevisions" },
      { name: "rejectAllRevisions" },
    ],
  },
  {
    num: "10",
    title: "Document properties",
    entries: [
      { name: "coreProperties" },
      { name: "setCoreProperties" },
      { name: "appProperties" },
      { name: "setAppProperties" },
      { name: "title" },
      { name: "author" },
      { name: "setTitle" },
      { name: "setAuthor" },
    ],
  },
  {
    num: "11",
    title: "Diagnostics",
    entries: [
      { name: "validate" },
      { name: "validatePackage" },
      { name: "statistics" },
      { name: "outline" },
      { name: "fields" },
      { name: "text" },
    ],
  },
  {
    num: "12",
    title: "Low-level part access",
    entries: [
      { name: "stylesPart" },
      { name: "numberingPart" },
      { name: "commentsPart" },
      { name: "footnotesPart" },
      { name: "endnotesPart" },
    ],
  },
  {
    num: "13",
    title: "Page-size & margin constants",
    entries: [
      { name: "PAGE_SIZE_A4" },
      { name: "PAGE_SIZE_LETTER" },
      { name: "MARGINS_NORMAL" },
      { name: "VERSION" },
    ],
  },
  {
    num: "14",
    title: "Browser preview (@word-kit/preview)",
    entries: [{ name: "previewToDOM", sig: "(source, container, options?) => Promise<Handle>" }],
  },
];

export const apiTotalCount: number = apiGroups.reduce((n, g) => n + g.entries.length, 0);
