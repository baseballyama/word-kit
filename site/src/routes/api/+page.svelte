<script lang="ts">
  import { base } from '$app/paths';

  type Group = { num: string; title: string; entries: Array<{ name: string; sig?: string }> };

  const groups: Group[] = [
    {
      num: '01',
      title: 'Lifecycle',
      entries: [
        { name: 'createDocx', sig: '({ paragraphs? }?) => Docx' },
        { name: 'openDocx', sig: '(bytes: Uint8Array) => Docx' },
        { name: 'fromBlob', sig: '(blob: Blob) => Promise<Docx>' },
        { name: 'toUint8Array', sig: '(doc: Docx) => Uint8Array' },
        { name: 'toBlob', sig: '(doc: Docx) => Blob' },
        { name: 'clone', sig: '(doc: Docx) => Docx' },
      ],
    },
    {
      num: '02',
      title: 'Paragraphs & blocks',
      entries: [
        { name: 'appendParagraph' },
        { name: 'insertParagraphAt' },
        { name: 'removeParagraph' },
        { name: 'appendHeading' },
        { name: 'appendPageBreak' },
        { name: 'appendLineBreak' },
        { name: 'appendSectionBreak' },
        { name: 'clearBody' },
        { name: 'paragraphs' },
      ],
    },
    {
      num: '03',
      title: 'Inline & text',
      entries: [
        { name: 'replaceText' },
        { name: 'replaceTextEverywhere' },
        { name: 'findText' },
        { name: 'findTextEverywhere' },
        { name: 'appendTextRun' },
        { name: 'setParagraphText' },
        { name: 'paragraphText' },
        { name: 'setRunFormat' },
        { name: 'clearRunFormat' },
        { name: 'getRunFormat' },
        { name: 'setParagraphAlignment' },
        { name: 'getParagraphAlignment' },
        { name: 'setParagraphIndent' },
        { name: 'setParagraphSpacing' },
        { name: 'setParagraphBorders' },
        { name: 'setParagraphShading' },
        { name: 'getParagraphStyle' },
        { name: 'getParagraphNumbering' },
        { name: 'mergeAdjacentRuns' },
        { name: 'mergeAdjacentRunsInBody' },
      ],
    },
    {
      num: '04',
      title: 'Styles & numbering',
      entries: [
        { name: 'addStyle' },
        { name: 'removeStyle' },
        { name: 'listStyles' },
        { name: 'ensureHeadingStyles' },
        { name: 'findStyleIdByName' },
        { name: 'setParagraphStyle' },
        { name: 'addBulletList' },
        { name: 'addNumberedList' },
        { name: 'applyListToParagraph' },
        { name: 'mergeStylesFromTemplate' },
      ],
    },
    {
      num: '05',
      title: 'Tables',
      entries: [
        { name: 'addTable' },
        { name: 'tables' },
        { name: 'removeTable' },
        { name: 'removeAllTables' },
        { name: 'unwrapTable' },
        { name: 'appendTableRow' },
        { name: 'removeTableRow' },
        { name: 'setTableRowAsHeader' },
        { name: 'setTableRowHeight' },
        { name: 'setTableBorders' },
        { name: 'setTableCellText' },
        { name: 'getTableCellText' },
        { name: 'setTableCellShading' },
        { name: 'setTableCellVerticalAlign' },
      ],
    },
    {
      num: '06',
      title: 'Images',
      entries: [
        { name: 'addImage' },
        { name: 'addImageRun' },
        { name: 'insertImageInto' },
        { name: 'images' },
        { name: 'imageReferences' },
        { name: 'replaceImage' },
        { name: 'replaceImageByAltText' },
        { name: 'removeAllImages' },
      ],
    },
    {
      num: '07',
      title: 'Headers, footers, sections',
      entries: [
        { name: 'addHeader' },
        { name: 'addFooter' },
        { name: 'addPageNumberFooter' },
        { name: 'setPageSize' },
        { name: 'setPageMargins' },
        { name: 'setPageOrientation' },
        { name: 'headers' },
        { name: 'footers' },
        { name: 'removeAllHeaders' },
        { name: 'removeAllFooters' },
      ],
    },
    {
      num: '08',
      title: 'Comments, notes, hyperlinks, bookmarks',
      entries: [
        { name: 'addComment' },
        { name: 'addFootnote' },
        { name: 'addEndnote' },
        { name: 'removeAllComments' },
        { name: 'removeAllFootnotes' },
        { name: 'removeAllEndnotes' },
        { name: 'addHyperlink' },
        { name: 'addInternalHyperlink' },
        { name: 'externalHyperlinks' },
        { name: 'setHyperlinkUrl' },
        { name: 'removeAllHyperlinks' },
        { name: 'addBookmark' },
        { name: 'removeBookmark' },
        { name: 'removeAllBookmarks' },
        { name: 'bookmarks' },
      ],
    },
    {
      num: '09',
      title: 'Fields & tracked changes',
      entries: [
        { name: 'appendField' },
        { name: 'addTableOfContents' },
        { name: 'appendMergeField' },
        { name: 'acceptAllRevisions' },
        { name: 'rejectAllRevisions' },
      ],
    },
    {
      num: '10',
      title: 'Document properties',
      entries: [
        { name: 'coreProperties' },
        { name: 'setCoreProperties' },
        { name: 'appProperties' },
        { name: 'setAppProperties' },
        { name: 'title' },
        { name: 'author' },
        { name: 'setTitle' },
        { name: 'setAuthor' },
      ],
    },
    {
      num: '11',
      title: 'Diagnostics',
      entries: [
        { name: 'validate' },
        { name: 'validatePackage' },
        { name: 'statistics' },
        { name: 'outline' },
        { name: 'fields' },
        { name: 'text' },
      ],
    },
    {
      num: '12',
      title: 'Low-level part access',
      entries: [
        { name: 'stylesPart' },
        { name: 'numberingPart' },
        { name: 'commentsPart' },
        { name: 'footnotesPart' },
        { name: 'endnotesPart' },
      ],
    },
    {
      num: '13',
      title: 'Page-size & margin constants',
      entries: [
        { name: 'PAGE_SIZE_A4' },
        { name: 'PAGE_SIZE_LETTER' },
        { name: 'MARGINS_NORMAL' },
        { name: 'VERSION' },
      ],
    },
    {
      num: '14',
      title: 'Browser preview (@word-kit/preview)',
      entries: [
        { name: 'previewToDOM', sig: '(source, container, options?) => Promise<Handle>' },
      ],
    },
  ];

  const totalCount = groups.reduce((n, g) => n + g.entries.length, 0);
</script>

<svelte:head>
  <title>API · word-kit</title>
</svelte:head>

<section class="head">
  <div class="head-inner">
    <p class="eyebrow">§ 03 · API reference</p>
    <h1>Every public export, grouped by area.</h1>
    <p class="lede">
      <code>@word-kit/core</code> exposes <strong>{totalCount}</strong> standalone functions
      and constants. Each one takes a <code>Docx</code> as its first argument (where
      applicable) and is side-effect-free at the module level, so bundlers can tree-shake
      anything you don't import. For full signatures and parameter shapes, see
      <code>packages/core/src/index.ts</code> and the
      <a href="https://github.com/baseballyama/word-kit">repository</a>.
    </p>
  </div>
</section>

<section class="api">
  <div class="api-inner">
    {#each groups as g (g.num)}
      <article class="group">
        <header class="group-head">
          <span class="g-num">{g.num}</span>
          <h2>{g.title}</h2>
          <span class="g-count">{g.entries.length.toString().padStart(2, '0')}</span>
        </header>
        <ul class="entries">
          {#each g.entries as e (e.name)}
            <li class="entry">
              <code class="name">{e.name}</code>
              {#if e.sig}<span class="sig">{e.sig}</span>{/if}
            </li>
          {/each}
        </ul>
      </article>
    {/each}

    <p class="more">
      Need usage shapes? Read the <a href="{base}/docs/getting-started">getting-started page</a>
      or the type-checked <a href="{base}/docs/recipes">recipes</a>.
    </p>
  </div>
</section>

<style>
  .head {
    padding: 3rem 1.5rem 1.75rem;
    border-bottom: 1px solid var(--border);
  }

  .head-inner {
    max-width: var(--max-wide);
    margin: 0 auto;
  }

  h1 {
    font-family: var(--display);
    font-weight: 460;
    font-size: clamp(1.9rem, 4.5vw, 2.85rem);
    line-height: 1.05;
    margin: 0 0 0.8rem;
    font-variation-settings: 'opsz' 144, 'SOFT' 30;
    letter-spacing: -0.025em;
    max-width: 22ch;
  }

  .lede {
    color: var(--fg-soft);
    font-size: 1.04rem;
    line-height: 1.55;
    max-width: 70ch;
    margin: 0;
  }

  .api {
    padding: 2.5rem 1.5rem 5rem;
  }

  .api-inner {
    max-width: var(--max-wide);
    margin: 0 auto;
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(320px, 1fr));
    gap: 1rem;
  }

  .group {
    border: 1px solid var(--border);
    border-radius: var(--radius);
    background: var(--bg-elev);
    overflow: hidden;
  }

  .group-head {
    display: flex;
    align-items: baseline;
    gap: 0.7rem;
    padding: 0.85rem 1rem 0.7rem;
    background: var(--bg-paper);
    border-bottom: 1px solid var(--border);
  }

  .g-num {
    font-family: var(--mono);
    font-size: 11px;
    color: var(--accent);
    letter-spacing: 0.08em;
  }

  .group-head h2 {
    margin: 0;
    border: none;
    padding: 0;
    font-family: var(--display);
    font-size: 1.05rem;
    font-weight: 540;
    flex: 1;
    font-variation-settings: 'opsz' 32, 'SOFT' 25;
  }

  .g-count {
    font-family: var(--mono);
    font-size: 11px;
    color: var(--fg-muted);
    letter-spacing: 0.1em;
  }

  .entries {
    list-style: none;
    padding: 0;
    margin: 0;
  }

  .entry {
    display: flex;
    align-items: baseline;
    gap: 0.7rem;
    padding: 0.45rem 1rem;
    border-bottom: 1px dashed var(--rule);
    font-size: 0.9rem;
  }

  .entry:last-child {
    border-bottom: none;
  }

  .name {
    font-family: var(--mono);
    font-size: 12.5px;
    background: transparent;
    border: none;
    padding: 0;
    color: var(--fg);
    flex: none;
  }

  .sig {
    font-family: var(--mono);
    font-size: 11.5px;
    color: var(--fg-muted);
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }

  .more {
    grid-column: 1 / -1;
    margin-top: 1.25rem;
    color: var(--fg-soft);
    border-top: 1px solid var(--border);
    padding-top: 1.25rem;
    font-size: 0.95rem;
  }
</style>
