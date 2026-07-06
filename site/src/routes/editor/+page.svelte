<script lang="ts">
  import { onMount } from 'svelte';
  import {
    createEditor,
    openEditor,
    runCommand,
    commands,
    runAtPath,
    paragraphAt,
    type EditorModel,
    type Command,
  } from '@office-kit/docx-editor';
  import {
    createDocx,
    appendHeading,
    appendParagraph,
    addTable,
    toUint8Array,
    text as documentText,
    getRunFormat,
    getParagraphStyle,
    replaceText,
    PAGE_SIZE_A4,
    PAGE_SIZE_LETTER,
    MARGINS_NORMAL,
    setPageSize,
    setPageMargins,
    ensureHeadingStyles,
  } from '@office-kit/docx';
  import { editorFor } from '@office-kit/docx-editor';
  import EditorCanvas from '$lib/editor/EditorCanvas.svelte';
  import RawXmlInspector from '$lib/editor/RawXmlInspector.svelte';
  import { t, locale, setLocale, LOCALES, type LocaleId } from '$lib/editor/i18n.svelte';

  type TabId = 'home' | 'insert' | 'layout' | 'references' | 'review';

  let model = $state<EditorModel | null>(null);
  let version = $state(0);
  let tick = $state(0);
  let tab = $state<TabId>('home');
  let fileName = $state('Untitled.docx');
  let status = $state('');
  let showXml = $state(false);
  let zoom = $state(1);
  let wordCount = $state(0);
  let charCount = $state(0);

  // Find & replace.
  let showFind = $state(false);
  let findQuery = $state('');
  let replaceValue = $state('');
  let findStatus = $state('');

  // Ribbon input values (kept in sync with the caret).
  let fontSize = $state(11);
  let fontColor = $state('#111111');
  let highlightColor = $state('#ffff00');
  let fontFamily = $state('Calibri');
  let currentStyle = $state('');

  onMount(() => {
    model = editorFor(sampleDoc());
    version++;
    refreshCounts();
    // Dev-only handle so the editor can be driven/inspected from the console
    // (and by the e2e verification). Stripped from production builds.
    if (import.meta.env.DEV)
      (window as unknown as { wkEditorModel?: () => EditorModel | null }).wkEditorModel = () => model;
  });

  function sampleDoc() {
    const doc = createDocx({ paragraphs: [] });
    setPageSize(doc, PAGE_SIZE_A4);
    setPageMargins(doc, MARGINS_NORMAL);
    ensureHeadingStyles(doc);
    appendHeading(doc, 'Welcome to the word-kit editor', 1);
    appendParagraph(
      doc,
      'This is a Word-like editor. Every edit routes through @office-kit/docx, so what you save is a real .docx. Type here, use the ribbon above, then Download.',
    );
    appendHeading(doc, 'Try it', 2);
    appendParagraph(doc, 'Select text and click Bold, or change alignment and color.');
    addTable(doc, [
      ['Feature', 'Status'],
      ['Text formatting', 'Editable'],
      ['Tables', 'Editable'],
    ]);
    return doc;
  }

  /** Run a command and force a structural re-render. */
  function apply<P>(cmd: Command<P>, params: P): void {
    if (!model) return;
    runCommand(model, cmd, params);
    version++;
    tick++;
    refreshCounts();
    status = `${cmd.label} applied.`;
  }

  function active(cmd: Command<unknown>): boolean {
    // `tick` is read so the button state recomputes on selection/edit changes.
    return tick >= 0 && !!model && (cmd.isActive?.(model) ?? false);
  }

  function enabled(cmd: Command<unknown>): boolean {
    return tick >= 0 && !!model && (cmd.isEnabled?.(model) ?? true);
  }

  function refreshCounts(): void {
    if (!model) return;
    const t = documentText(model.doc);
    charCount = t.length;
    const words = t.trim().match(/\S+/g);
    wordCount = words ? words.length : 0;
  }

  /** Reflect the caret's run formatting + paragraph style in the ribbon inputs. */
  function syncToolbarFromCaret(): void {
    const pos = model?.selection?.focus;
    if (!model || !pos) return;
    const run = runAtPath(model.doc, pos);
    if (run) {
      const fmt = getRunFormat(run);
      if (fmt.font) fontFamily = fmt.font;
      if (fmt.fontSizeHalfPoints) fontSize = fmt.fontSizeHalfPoints / 2;
      if (fmt.color) fontColor = `#${fmt.color}`;
    }
    const para = paragraphAt(model.doc, pos);
    currentStyle = (para && getParagraphStyle(para)) || '';
  }

  function onCanvasSelection(): void {
    tick++;
    syncToolbarFromCaret();
  }

  /** Canvas-originated edit (typing / Enter / paste / undo): refresh chrome only. */
  function onCanvasEdit(): void {
    tick++;
    refreshCounts();
    syncToolbarFromCaret();
    status = t('status.editing');
  }

  function undo(): void {
    if (!model) return;
    model.undo();
    version++;
    tick++;
    refreshCounts();
    status = 'Undo.';
  }

  function redo(): void {
    if (!model) return;
    model.redo();
    version++;
    tick++;
    refreshCounts();
    status = 'Redo.';
  }

  function doReplaceAll(): void {
    if (!model || !findQuery) return;
    const n = replaceText(model.doc, findQuery, replaceValue);
    version++;
    tick++;
    refreshCounts();
    findStatus = `Replaced ${n} occurrence${n === 1 ? '' : 's'}.`;
  }

  function setZoom(z: number): void {
    zoom = Math.max(0.5, Math.min(2, z));
  }

  function onGlobalKey(e: KeyboardEvent): void {
    const mod = e.ctrlKey || e.metaKey;
    if (mod && (e.key.toLowerCase() === 'h' || e.key.toLowerCase() === 'f')) {
      e.preventDefault();
      showFind = e.key.toLowerCase() === 'h' ? !showFind : true;
    }
  }

  async function openFile(e: Event): Promise<void> {
    const input = e.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;
    const bytes = new Uint8Array(await file.arrayBuffer());
    try {
      model = openEditor(bytes);
      fileName = file.name;
      version++;
      status = `Opened ${file.name}.`;
    } catch (err) {
      status = `Failed to open: ${(err as Error).message}`;
    }
  }

  function download(): void {
    if (!model) return;
    const bytes = toUint8Array(model.doc);
    // Copy into a plain ArrayBuffer so the Blob part type is unambiguous
    // (Uint8Array<ArrayBufferLike> is not assignable to BlobPart under strict DOM libs).
    const buffer = bytes.slice().buffer;
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName.endsWith('.docx') ? fileName : `${fileName}.docx`;
    a.click();
    URL.revokeObjectURL(url);
    status = `Downloaded ${a.download}.`;
  }

  function newDoc(): void {
    model = createEditor();
    version++;
    fileName = 'Untitled.docx';
    status = 'New document.';
  }

  function insertTable(): void {
    const spec = prompt('Table size (rows x cols)', '3x3');
    if (!spec) return;
    const [r, c] = spec.split(/[x×,]/).map((n) => Number.parseInt(n.trim(), 10));
    apply(commands.insertTableCommand, { rows: r || 2, cols: c || 2 });
  }

  function insertLink(): void {
    const url = prompt('Link URL', 'https://');
    if (!url) return;
    const text = prompt('Link text', url) ?? url;
    apply(commands.insertHyperlinkCommand, { url, text });
  }

  function insertHeading(level: number): void {
    const text = prompt(`Heading ${level} text`, 'Heading');
    if (text != null) apply(commands.insertHeadingCommand, { text, level });
  }

  function addComment(): void {
    const text = prompt('Comment');
    if (text) apply(commands.addCommentCommand, { author: 'You', initials: 'Y', text });
  }

  async function insertImage(e: Event): Promise<void> {
    const input = e.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file || !model) return;
    const bytes = new Uint8Array(await file.arrayBuffer());
    apply(commands.insertImageCommand, { bytes, options: { widthEmu: 2743200, heightEmu: 2057400 } });
    input.value = '';
  }

  const TAB_IDS: TabId[] = ['home', 'insert', 'layout', 'references', 'review'];
</script>

<svelte:head><title>word-kit editor</title></svelte:head>
<svelte:window onkeydown={onGlobalKey} />

<div class="app">
  <!-- Title bar -->
  <div class="titlebar">
    <span class="brand">word-kit</span>
    <input class="filename" bind:value={fileName} aria-label="File name" />
    <button onclick={undo} disabled={!model?.canUndo()} title={t('action.undo') + ' (Ctrl+Z)'}>↶</button>
    <button onclick={redo} disabled={!model?.canRedo()} title={t('action.redo') + ' (Ctrl+Y)'}>↷</button>
    <div class="spacer"></div>
    <select
      class="lang"
      value={locale()}
      onchange={(e) => setLocale((e.currentTarget as HTMLSelectElement).value as LocaleId)}
      title={t('language')}
      aria-label={t('language')}
    >
      {#each Object.entries(LOCALES) as [id, name] (id)}
        <option value={id}>{name}</option>
      {/each}
    </select>
    <button onclick={newDoc}>{t('action.new')}</button>
    <label class="btn-file">
      {t('action.open')}
      <input type="file" accept=".docx" onchange={openFile} hidden />
    </label>
    <button class:primary={showFind} onclick={() => (showFind = !showFind)} title="Ctrl+H">🔍 {t('action.find')}</button>
    <button class:primary={showXml} onclick={() => (showXml = !showXml)} title={t('xml.title')}>&lt;/&gt; {t('action.xml')}</button>
    <button class="primary" onclick={download}>{t('action.download')}</button>
  </div>

  {#if showFind}
    <div class="findbar">
      <input placeholder={t('find.find')} bind:value={findQuery} aria-label={t('find.find')} />
      <input placeholder={t('find.replaceWith')} bind:value={replaceValue} aria-label={t('find.replaceWith')} />
      <button onclick={doReplaceAll} disabled={!findQuery}>{t('find.replaceAll')}</button>
      <span class="find-status">{findStatus}</span>
      <button class="find-close" onclick={() => (showFind = false)} aria-label={t('find.close')}>✕</button>
    </div>
  {/if}

  <!-- Ribbon tabs -->
  <div class="tabs">
    {#each TAB_IDS as id (id)}
      <button class="tab" class:active={tab === id} onclick={() => (tab = id)}>{t(`tab.${id}`)}</button>
    {/each}
  </div>

  <!-- Ribbon body -->
  <div class="ribbon">
    {#if tab === 'home'}
      <div class="group">
        <div class="group-body">
          <select
            bind:value={fontFamily}
            onchange={() => apply(commands.setFontCommand, { font: fontFamily })}
            aria-label={t('group.font')}
          >
            <option>Calibri</option><option>Arial</option><option>Times New Roman</option>
            <option>Georgia</option><option>Courier New</option><option>Yu Gothic</option>
          </select>
          <input
            type="number"
            min="6"
            max="96"
            bind:value={fontSize}
            onchange={() => apply(commands.setFontSizeCommand, { points: fontSize })}
            aria-label="Font size"
            class="num"
          />
        </div>
        <div class="group-label">{t('group.font')}</div>
      </div>

      <div class="group">
        <div class="group-body">
          <button class:on={active(commands.toggleBoldCommand)} disabled={!enabled(commands.toggleBoldCommand)} onclick={() => apply(commands.toggleBoldCommand, undefined)}><b>B</b></button>
          <button class:on={active(commands.toggleItalicCommand)} disabled={!enabled(commands.toggleItalicCommand)} onclick={() => apply(commands.toggleItalicCommand, undefined)}><i>I</i></button>
          <button class:on={active(commands.toggleUnderlineCommand)} disabled={!enabled(commands.toggleUnderlineCommand)} onclick={() => apply(commands.toggleUnderlineCommand, undefined)}><u>U</u></button>
          <button class:on={active(commands.toggleStrikeCommand)} onclick={() => apply(commands.toggleStrikeCommand, undefined)}><s>S</s></button>
          <label class="color">A<input type="color" bind:value={fontColor} onchange={() => apply(commands.setColorCommand, { color: fontColor })} aria-label="Font color" /></label>
          <label class="color hl">▨<input type="color" bind:value={highlightColor} onchange={() => apply(commands.setHighlightCommand, { color: highlightColor })} aria-label="Highlight" /></label>
          <button onclick={() => apply(commands.clearFormattingCommand, undefined)} title="Clear formatting">⌫</button>
        </div>
        <div class="group-label">{t('group.text')}</div>
      </div>

      <div class="group">
        <div class="group-body">
          <button class:on={active(commands.alignLeftCommand)} onclick={() => apply(commands.alignLeftCommand, undefined)} title="Align left">⯇</button>
          <button class:on={active(commands.alignCenterCommand)} onclick={() => apply(commands.alignCenterCommand, undefined)} title="Center">≡</button>
          <button class:on={active(commands.alignRightCommand)} onclick={() => apply(commands.alignRightCommand, undefined)} title="Align right">⯈</button>
          <button class:on={active(commands.alignJustifyCommand)} onclick={() => apply(commands.alignJustifyCommand, undefined)} title="Justify">☰</button>
          <button onclick={() => apply(commands.applyListCommand, { kind: 'bullet' })} title="Bullet list">• —</button>
          <button onclick={() => apply(commands.applyListCommand, { kind: 'numbered' })} title="Numbered list">1. —</button>
          <button onclick={() => apply(commands.setIndentCommand, { left: 720 })} title="Indent">⇥</button>
        </div>
        <div class="group-label">{t('group.paragraph')}</div>
      </div>

      <div class="group">
        <div class="group-body">
          <select value={currentStyle} onchange={(e) => { const v = (e.currentTarget as HTMLSelectElement).value; apply(commands.setParagraphStyleCommand, { styleId: v || undefined }); }} aria-label="Paragraph style">
            <option value="">{t('style.normal')}</option>
            <option value="Heading1">{t('style.heading1')}</option>
            <option value="Heading2">{t('style.heading2')}</option>
            <option value="Heading3">{t('style.heading3')}</option>
          </select>
        </div>
        <div class="group-label">{t('group.styles')}</div>
      </div>
    {:else if tab === 'insert'}
      <div class="group">
        <div class="group-body">
          <button onclick={() => apply(commands.insertParagraphCommand, { text: '' })}>¶ {t('ins.paragraph')}</button>
          <button onclick={insertTable}>▦ {t('ins.table')}</button>
          <button onclick={() => apply(commands.insertPageBreakCommand, undefined)}>⤓ {t('ins.pageBreak')}</button>
          <button onclick={insertLink}>🔗 {t('ins.link')}</button>
          <label class="btn-inline">
            🖼 {t('ins.picture')}
            <input type="file" accept="image/*" onchange={insertImage} hidden />
          </label>
        </div>
        <div class="group-label">{t('group.insert')}</div>
      </div>
      <div class="group">
        <div class="group-body">
          <button onclick={() => insertHeading(1)}>H1</button>
          <button onclick={() => insertHeading(2)}>H2</button>
          <button onclick={() => insertHeading(3)}>H3</button>
        </div>
        <div class="group-label">{t('group.headings')}</div>
      </div>
    {:else if tab === 'layout'}
      <div class="group">
        <div class="group-body">
          <button onclick={() => apply(commands.setPageSizeCommand, { size: PAGE_SIZE_A4 })}>A4</button>
          <button onclick={() => apply(commands.setPageSizeCommand, { size: PAGE_SIZE_LETTER })}>Letter</button>
          <button onclick={() => apply(commands.setOrientationCommand, { orientation: 'portrait' })}>Portrait</button>
          <button onclick={() => apply(commands.setOrientationCommand, { orientation: 'landscape' })}>Landscape</button>
        </div>
        <div class="group-label">{t('group.pageSetup')}</div>
      </div>
      <div class="group">
        <div class="group-body">
          <button onclick={() => apply(commands.insertSectionBreakCommand, { type: 'nextPage' })}>{t('layout.sectionBreak')}</button>
          <button onclick={() => apply(commands.addPageNumberFooterCommand, {})}>{t('layout.pageNumbers')}</button>
        </div>
        <div class="group-label">{t('group.sections')}</div>
      </div>
    {:else if tab === 'references'}
      <div class="group">
        <div class="group-body">
          <button onclick={() => apply(commands.insertTocCommand, {})}>{t('ref.toc')}</button>
          <button onclick={() => { const txt = prompt('Footnote text'); if (txt) apply(commands.addFootnoteCommand, { text: txt }); }}>{t('ref.footnote')}</button>
          <button onclick={() => { const txt = prompt('Endnote text'); if (txt) apply(commands.addEndnoteCommand, { text: txt }); }}>{t('ref.endnote')}</button>
          <button onclick={() => { const n = prompt('Bookmark name'); if (n) apply(commands.addBookmarkCommand, { name: n }); }}>{t('ref.bookmark')}</button>
        </div>
        <div class="group-label">{t('group.references')}</div>
      </div>
    {:else if tab === 'review'}
      <div class="group">
        <div class="group-body">
          <button onclick={addComment}>💬 {t('review.newComment')}</button>
          <button onclick={() => apply(commands.acceptAllRevisionsCommand, undefined)}>✓ {t('review.acceptAll')}</button>
          <button onclick={() => apply(commands.rejectAllRevisionsCommand, undefined)}>✗ {t('review.rejectAll')}</button>
        </div>
        <div class="group-label">{t('group.tracking')}</div>
      </div>
    {/if}
  </div>

  <!-- Editing surface -->
  <div class="workspace">
    <div class="surface">
      {#if model}
        <EditorCanvas
          {model}
          {version}
          {zoom}
          onselectionchange={onCanvasSelection}
          onedit={onCanvasEdit}
        />
      {:else}
        <p class="loading">Loading editor…</p>
      {/if}
    </div>
    {#if showXml && model}
      <aside class="xml-panel">
        <div class="xml-head">{t('xml.title')}</div>
        <RawXmlInspector {model} onchange={() => { version++; tick++; status = 'Raw XML edited.'; }} />
      </aside>
    {/if}
  </div>

  <div class="statusbar">
    <span>{status || t('status.ready')}</span>
    <div class="spacer"></div>
    <span class="count">{wordCount} {t('status.words')} · {charCount} {t('status.chars')}</span>
    <div class="zoom">
      <button onclick={() => setZoom(zoom - 0.1)} aria-label="Zoom out">−</button>
      <span class="zoom-val">{Math.round(zoom * 100)}%</span>
      <button onclick={() => setZoom(zoom + 0.1)} aria-label="Zoom in">+</button>
    </div>
  </div>
</div>

<style>
  .app {
    display: flex;
    flex-direction: column;
    height: 100vh;
    background: #f3f2f1;
    font-family: 'Segoe UI', system-ui, sans-serif;
    color: #201f1e;
  }
  .titlebar {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 6px 12px;
    background: #2b579a;
    color: #fff;
  }
  .brand { font-weight: 700; letter-spacing: 0.02em; }
  .filename {
    background: rgba(255, 255, 255, 0.15);
    border: 1px solid transparent;
    color: #fff;
    padding: 3px 8px;
    border-radius: 3px;
    width: 180px;
  }
  .filename:focus { background: #fff; color: #201f1e; outline: none; }
  .spacer { flex: 1; }
  .titlebar button, .btn-file {
    background: rgba(255, 255, 255, 0.15);
    color: #fff;
    border: none;
    padding: 5px 12px;
    border-radius: 4px;
    cursor: pointer;
    font-size: 13px;
  }
  .titlebar button.primary { background: #fff; color: #2b579a; font-weight: 600; }
  .btn-file { display: inline-block; }
  .lang {
    background: rgba(255, 255, 255, 0.15);
    color: #fff;
    border: none;
    padding: 4px 6px;
    border-radius: 4px;
    font-size: 13px;
    cursor: pointer;
  }
  .lang option { color: #201f1e; }
  .tabs {
    display: flex;
    gap: 2px;
    background: #fff;
    padding: 0 8px;
    border-bottom: 1px solid #e1dfdd;
  }
  .tab {
    background: none;
    border: none;
    padding: 8px 16px;
    cursor: pointer;
    font-size: 13px;
    color: #444;
    border-bottom: 2px solid transparent;
  }
  .tab.active { color: #2b579a; border-bottom-color: #2b579a; font-weight: 600; }
  .ribbon {
    display: flex;
    gap: 2px;
    align-items: stretch;
    background: #faf9f8;
    padding: 6px 8px;
    border-bottom: 1px solid #e1dfdd;
    min-height: 72px;
    overflow-x: auto;
  }
  .group {
    display: flex;
    flex-direction: column;
    padding: 0 8px;
    border-right: 1px solid #edebe9;
  }
  .group-body {
    display: flex;
    flex-wrap: wrap;
    gap: 3px;
    align-items: center;
    flex: 1;
    max-width: 260px;
  }
  .group-label { text-align: center; font-size: 11px; color: #888; padding-top: 4px; }
  .ribbon button {
    min-width: 30px;
    height: 30px;
    padding: 0 8px;
    border: 1px solid transparent;
    background: transparent;
    border-radius: 4px;
    cursor: pointer;
    font-size: 14px;
    color: #201f1e;
  }
  .ribbon button:hover { background: #edebe9; border-color: #d2d0ce; }
  .ribbon button.on { background: #cfe0f4; border-color: #2b579a; }
  .ribbon button:disabled { opacity: 0.4; cursor: default; }
  .ribbon select, .ribbon .num {
    height: 28px;
    border: 1px solid #d2d0ce;
    border-radius: 4px;
    padding: 0 6px;
    background: #fff;
  }
  .ribbon .num { width: 52px; }
  .color { display: inline-flex; align-items: center; gap: 2px; cursor: pointer; font-size: 13px; }
  .color input[type='color'] { width: 20px; height: 20px; border: none; background: none; padding: 0; cursor: pointer; }
  .workspace {
    flex: 1;
    display: flex;
    min-height: 0;
  }
  .surface {
    flex: 1;
    overflow-y: auto;
    padding: 24px;
  }
  .xml-panel {
    width: 340px;
    border-left: 1px solid #e1dfdd;
    background: #fff;
    display: flex;
    flex-direction: column;
    min-height: 0;
  }
  .xml-head {
    font-size: 11px;
    font-weight: 600;
    color: #605e5c;
    padding: 6px 10px;
    border-bottom: 1px solid #edebe9;
    background: #faf9f8;
  }
  .btn-inline {
    display: inline-flex;
    align-items: center;
    height: 30px;
    padding: 0 8px;
    border: 1px solid transparent;
    border-radius: 4px;
    cursor: pointer;
    font-size: 14px;
  }
  .btn-inline:hover { background: #edebe9; border-color: #d2d0ce; }
  .loading { text-align: center; color: #888; }
  .statusbar {
    display: flex;
    align-items: center;
    gap: 12px;
    background: #2b579a;
    color: #fff;
    font-size: 12px;
    padding: 3px 12px;
  }
  .statusbar .spacer { flex: 1; }
  .statusbar .count { opacity: 0.9; }
  .zoom { display: flex; align-items: center; gap: 4px; }
  .zoom button {
    background: rgba(255, 255, 255, 0.15);
    color: #fff;
    border: none;
    width: 20px;
    height: 20px;
    border-radius: 3px;
    cursor: pointer;
    line-height: 1;
  }
  .zoom-val { min-width: 38px; text-align: center; }
  .findbar {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 6px 12px;
    background: #fff;
    border-bottom: 1px solid #e1dfdd;
  }
  .findbar input {
    height: 28px;
    border: 1px solid #d2d0ce;
    border-radius: 4px;
    padding: 0 8px;
    font-size: 13px;
  }
  .findbar button {
    height: 28px;
    padding: 0 12px;
    border: 1px solid #d2d0ce;
    border-radius: 4px;
    background: #f3f2f1;
    cursor: pointer;
    font-size: 13px;
  }
  .findbar button:disabled { opacity: 0.5; cursor: default; }
  .find-status { font-size: 12px; color: #605e5c; }
  .find-close { margin-left: auto; }
</style>
