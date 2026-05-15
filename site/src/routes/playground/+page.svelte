<script lang="ts">
  import { base } from '$app/paths';
  import { onMount } from 'svelte';

  let container = $state<HTMLDivElement | null>(null);
  let status = $state<string>('Ready.');
  let busy = $state<boolean>(false);
  let fileName = $state<string>('built-in sample');
  let docxBytes = $state<number>(0);
  let dropping = $state<boolean>(false);
  let docxModule = $state<typeof import('@word-kit/core') | null>(null);
  let previewModule = $state<typeof import('@word-kit/preview') | null>(null);
  // The actual bytes currently mounted in the preview, kept so that the
  // download button can return the same .docx the user is looking at —
  // whether that's the built-in sample or a file they just uploaded.
  let lastBytes = $state<Uint8Array | null>(null);
  let currentDispose: (() => void) | null = null;

  async function ensureModules(): Promise<{
    core: typeof import('@word-kit/core');
    preview: typeof import('@word-kit/preview');
  }> {
    if (docxModule && previewModule) return { core: docxModule, preview: previewModule };
    status = 'Loading @word-kit/core + @word-kit/preview…';
    const [coreMod, previewMod] = await Promise.all([
      import('@word-kit/core'),
      import('@word-kit/preview'),
    ]);
    docxModule = coreMod;
    previewModule = previewMod;
    return { core: coreMod, preview: previewMod };
  }

  function buildSampleBytes(core: typeof import('@word-kit/core')): Uint8Array {
    const {
      addBulletList,
      addTable,
      appendHeading,
      appendParagraph,
      createDocx,
      MARGINS_NORMAL,
      PAGE_SIZE_A4,
      setCoreProperties,
      setPageMargins,
      setPageSize,
      toUint8Array,
    } = core;

    const doc = createDocx({ paragraphs: [] });
    setPageSize(doc, PAGE_SIZE_A4);
    setPageMargins(doc, MARGINS_NORMAL);
    setCoreProperties(doc, {
      title: 'word-kit playground sample',
      creator: 'word-kit',
      description: 'Built in the browser by @word-kit/core, rendered by @word-kit/preview.',
    });

    appendHeading(doc, 'word-kit — playground sample', 1);
    appendParagraph(
      doc,
      'This document was generated entirely in the browser by @word-kit/core, then mounted into the page by @word-kit/preview. No server, no Word, no PDF.',
    );

    appendHeading(doc, 'What you can do here', 2);
    addBulletList(doc, [
      'Paragraphs, headings, and run formatting',
      'Numbered and bulleted lists',
      'Tables with rows and columns',
      'A4 / Letter page sizes and margins',
      'Tracked changes, comments, footnotes (see Recipes)',
      'Lossless round-trip — open and re-save without breaking the file',
    ]);

    appendHeading(doc, 'A small table', 2);
    addTable(doc, [
      ['Quarter', 'Shipped', 'Open'],
      ['Q1', '14', '3'],
      ['Q2', '19', '5'],
      ['Q3', '22', '2'],
    ]);

    appendHeading(doc, 'How to drive this from your own code', 2);
    appendParagraph(
      doc,
      "Drop any .docx onto the surface to the right, or open one with the file picker. The bytes never leave the page — preview runs entirely in your browser.",
    );

    return toUint8Array(doc);
  }

  async function renderBytes(bytes: Uint8Array): Promise<void> {
    if (!container) return;
    busy = true;
    try {
      const { preview } = await ensureModules();
      currentDispose?.();
      container.innerHTML = '';
      status = `Rendering ${bytes.byteLength.toLocaleString()} bytes…`;
      const handle = await preview.previewToDOM(bytes, container, {
        classPrefix: 'wk-',
        inWrapper: true,
        breakPages: true,
        renderFonts: true,
      });
      currentDispose = () => handle.dispose();
      docxBytes = bytes.byteLength;
      lastBytes = bytes;
      status = `Rendered ${fileName} — ${bytes.byteLength.toLocaleString()} bytes.`;
    } catch (err) {
      status = `Failed: ${(err as Error).message}`;
    } finally {
      busy = false;
    }
  }

  async function loadSample(): Promise<void> {
    const { core } = await ensureModules();
    fileName = 'built-in sample';
    const bytes = buildSampleBytes(core);
    await renderBytes(bytes);
  }

  async function loadFromFile(file: File): Promise<void> {
    fileName = file.name;
    const buf = await file.arrayBuffer();
    await renderBytes(new Uint8Array(buf));
  }

  function handleFileInput(event: Event): void {
    const input = event.currentTarget as HTMLInputElement;
    const file = input.files?.[0];
    if (file) void loadFromFile(file);
  }

  function handleDrop(event: DragEvent): void {
    event.preventDefault();
    dropping = false;
    const file = event.dataTransfer?.files[0];
    if (file) void loadFromFile(file);
  }

  function handleDragOver(event: DragEvent): void {
    event.preventDefault();
    dropping = true;
  }

  function handleDragLeave(): void {
    dropping = false;
  }

  function downloadCurrent(): void {
    if (!lastBytes) return;
    // Slice the in-memory bytes into a fresh ArrayBuffer so Blob's
    // BlobPart type (which wants ArrayBuffer-backed views, not
    // SharedArrayBuffer) is satisfied under strict TS lib settings.
    const ab = lastBytes.slice().buffer as ArrayBuffer;
    const blob = new Blob([ab], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName.endsWith('.docx') ? fileName : `${fileName || 'word-kit'}.docx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  onMount(() => {
    void loadSample();
    return () => currentDispose?.();
  });
</script>

<svelte:head>
  <title>Playground · word-kit</title>
</svelte:head>

<section class="head">
  <div class="head-inner">
    <p class="eyebrow">Playground</p>
    <h1>Render a <em>.docx</em> in your browser.</h1>
    <p class="lede">
      Drop a Word document onto the canvas, or generate a built-in sample with
      <code>@word-kit/core</code>. Preview is rendered by
      <a href="{base}/api">@word-kit/preview</a> entirely client-side — your file never
      leaves the page.
    </p>
  </div>
</section>

<section class="workbench">
  <aside class="controls" aria-label="Playground controls">
    <header class="controls-head">
      <span class="bracket">[</span>
      <span class="title">controls</span>
      <span class="bracket">]</span>
    </header>

    <button class="btn primary" onclick={loadSample} disabled={busy}>
      <span>Generate sample</span>
      <span class="arrow">↻</span>
    </button>

    <label class="file-btn">
      <input type="file" accept=".docx" onchange={handleFileInput} />
      <span>Open .docx…</span>
    </label>

    <button class="btn ghost" onclick={downloadCurrent} disabled={!lastBytes}>
      Download current bytes
    </button>

    <dl class="status">
      <div>
        <dt>file</dt>
        <dd>{fileName}</dd>
      </div>
      <div>
        <dt>size</dt>
        <dd>{docxBytes ? `${docxBytes.toLocaleString()} B` : '—'}</dd>
      </div>
      <div>
        <dt>state</dt>
        <dd class:busy>{status}</dd>
      </div>
    </dl>

    <p class="note">
      <strong>What is this?</strong>
      <br />
      A read-only render. To edit, build a <code>Docx</code> with
      <code>@word-kit/core</code>; <em>Download current bytes</em> hands you back
      whatever <code>.docx</code> is mounted right now — sample or upload.
    </p>

    <p class="note muted">
      Bundle: <code>@word-kit/core</code> ~131 KB full · <code>@word-kit/preview</code> wraps the OSS
      <code>docx-preview</code> renderer.
    </p>
  </aside>

  <div class="stage" class:dropping ondragover={handleDragOver} ondragleave={handleDragLeave} ondrop={handleDrop} role="region" aria-label="Document preview">
    <div class="stage-shell">
      <div class="stage-head">
        <span class="dot dot-a"></span>
        <span class="dot dot-b"></span>
        <span class="dot dot-c"></span>
        <span class="stage-name">{fileName}</span>
      </div>
      <div class="stage-scroll">
        <div bind:this={container} class="stage-canvas">
          <p class="stage-placeholder">Loading word-kit modules…</p>
        </div>
      </div>
      {#if dropping}
        <div class="drop-overlay">
          <span>Drop the .docx to render</span>
        </div>
      {/if}
    </div>
  </div>
</section>

<style>
  .head {
    padding: 2.6rem 1.5rem 1.5rem;
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
    margin: 0 0 0.85rem;
    font-variation-settings: 'opsz' 144, 'SOFT' 30;
    letter-spacing: -0.025em;
    max-width: 18ch;
  }

  h1 em {
    color: var(--accent);
    font-style: italic;
    font-variation-settings: 'opsz' 144, 'SOFT' 70;
  }

  .lede {
    color: var(--fg-soft);
    font-size: 1.04rem;
    line-height: 1.55;
    max-width: 64ch;
    margin: 0;
  }

  .workbench {
    display: grid;
    grid-template-columns: 280px 1fr;
    gap: 0;
    max-width: var(--max-wide);
    margin: 0 auto;
    padding: 1.75rem 1.5rem 4rem;
  }

  @media (max-width: 880px) {
    .workbench {
      grid-template-columns: 1fr;
      gap: 1.25rem;
    }
  }

  .controls {
    border: 1px solid var(--border);
    border-right: none;
    border-radius: var(--radius) 0 0 var(--radius);
    background: var(--bg-elev);
    padding: 1.1rem 1.1rem 1.25rem;
    display: flex;
    flex-direction: column;
    gap: 0.75rem;
  }

  @media (max-width: 880px) {
    .controls {
      border-right: 1px solid var(--border);
      border-radius: var(--radius);
    }
  }

  .controls-head {
    font-family: var(--mono);
    font-size: 11px;
    color: var(--fg-muted);
    letter-spacing: 0.18em;
    text-transform: uppercase;
    margin-bottom: 0.25rem;
    display: flex;
    gap: 0.4rem;
    align-items: center;
  }

  .controls-head .bracket {
    color: var(--accent);
  }

  .btn {
    display: inline-flex;
    align-items: center;
    justify-content: space-between;
    gap: 0.5rem;
    padding: 0.65rem 0.85rem;
    border-radius: var(--radius-sm);
    border: 1px solid var(--border);
    background: var(--bg-soft);
    color: var(--fg);
    font-family: var(--sans);
    font-size: 0.9rem;
    font-weight: 540;
    cursor: pointer;
    transition:
      background 120ms ease,
      border-color 120ms ease,
      color 120ms ease,
      transform 120ms ease;
  }

  .btn:hover:not([disabled]) {
    background: var(--bg-paper);
    border-color: var(--border-strong);
    transform: translateY(-1px);
  }

  .btn[disabled] {
    opacity: 0.55;
    cursor: not-allowed;
  }

  .btn.primary {
    background: var(--accent);
    color: var(--bg);
    border-color: var(--accent);
    box-shadow: 0 8px 24px -14px var(--accent-glow);
  }

  .btn.primary:hover:not([disabled]) {
    background: var(--accent-hot);
    border-color: var(--accent-hot);
  }

  .btn.ghost {
    background: transparent;
  }

  .btn .arrow {
    color: var(--bg);
    opacity: 0.7;
  }

  .file-btn {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    padding: 0.65rem 0.85rem;
    border-radius: var(--radius-sm);
    border: 1px dashed var(--border-strong);
    background: var(--bg-soft);
    color: var(--fg);
    font-family: var(--sans);
    font-size: 0.9rem;
    font-weight: 540;
    cursor: pointer;
    text-align: center;
    transition: background 120ms ease, border-color 120ms ease;
  }

  .file-btn:hover {
    background: var(--bg-paper);
    border-color: var(--accent);
    color: var(--fg);
  }

  .file-btn input {
    display: none;
  }

  .status {
    margin: 0.5rem 0 0;
    border: 1px solid var(--border);
    border-radius: var(--radius-sm);
    background: var(--code-bg);
    padding: 0.7rem 0.75rem;
    display: flex;
    flex-direction: column;
    gap: 0.45rem;
    font-family: var(--mono);
    font-size: 11.5px;
  }

  .status > div {
    display: grid;
    grid-template-columns: 4ch 1fr;
    align-items: baseline;
    gap: 0.6rem;
  }

  .status dt {
    color: var(--fg-muted);
    text-transform: uppercase;
    letter-spacing: 0.1em;
    font-size: 10px;
  }

  .status dd {
    color: var(--fg);
    margin: 0;
    word-break: break-word;
  }

  .status dd.busy {
    color: var(--accent);
  }

  .note {
    color: var(--fg-soft);
    font-size: 0.86rem;
    line-height: 1.5;
    margin: 0;
  }

  .note.muted {
    color: var(--fg-muted);
    font-size: 0.78rem;
  }

  .note strong {
    color: var(--fg);
  }

  .stage {
    border: 1px solid var(--border);
    border-radius: 0 var(--radius) var(--radius) 0;
    background: var(--bg-paper);
    overflow: hidden;
    position: relative;
    min-height: 70vh;
  }

  @media (max-width: 880px) {
    .stage {
      border-radius: var(--radius);
    }
  }

  .stage-shell {
    display: flex;
    flex-direction: column;
    height: 100%;
  }

  .stage-head {
    display: flex;
    align-items: center;
    gap: 0.45rem;
    padding: 0.6rem 0.85rem;
    background: var(--bg-elev);
    border-bottom: 1px solid var(--border);
    font-family: var(--mono);
    font-size: 12px;
    color: var(--fg-muted);
  }

  .dot {
    width: 9px;
    height: 9px;
    border-radius: 50%;
    background: var(--bg-soft);
    border: 1px solid var(--border-strong);
  }

  .dot-a {
    background: color-mix(in oklab, var(--accent) 60%, var(--bg-soft));
    border-color: var(--accent-soft);
  }

  .dot-b {
    background: color-mix(in oklab, var(--brass) 50%, var(--bg-soft));
    border-color: var(--brass-soft);
  }

  .stage-name {
    margin-left: 0.4rem;
    color: var(--fg-soft);
  }

  /* The scroll slot has zero padding so docx-preview's own wrapper (which
   * applies its own padding + centring) sits flush against the edges.
   * Overflow auto handles documents whose page width exceeds the viewport. */
  .stage-scroll {
    flex: 1;
    overflow: auto;
    padding: 0;
    background: var(--bg-paper);
  }

  /* Transparent slot for docx-preview's rendered tree. The renderer paints
   * its own grey wrapper + white pages + page shadows; we deliberately
   * don't fight those styles with our own border / background / padding —
   * doing so used to push the page off the top-left of the canvas. */
  .stage-canvas {
    min-height: 60vh;
  }

  .stage-placeholder {
    color: var(--fg-muted);
    font-family: var(--mono);
    font-size: 12px;
    padding: 1.2rem 1.4rem;
    margin: 0;
  }

  .drop-overlay {
    position: absolute;
    inset: 0;
    display: flex;
    align-items: center;
    justify-content: center;
    background: color-mix(in oklab, var(--accent) 18%, transparent);
    border: 2px dashed var(--accent);
    color: var(--fg);
    font-family: var(--mono);
    font-size: 14px;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    backdrop-filter: blur(2px);
    pointer-events: none;
  }

  .stage.dropping {
    border-color: var(--accent);
  }
</style>
