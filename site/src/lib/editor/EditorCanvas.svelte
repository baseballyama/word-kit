<script lang="ts">
  import type { EditorModel, DocPosition } from '@office-kit/docx-editor';
  import {
    renderDocumentHtml,
    readDomSelection,
    runAtPath,
    setSimpleRunText,
    runCommand,
    commands,
  } from '@office-kit/docx-editor';

  interface Props {
    model: EditorModel;
    /** Bumped by the parent whenever a command mutates the doc structurally. */
    version: number;
    /** Zoom factor (1 = 100%). */
    zoom?: number;
    /** Called after selection changes so the ribbon can refresh active states. */
    onselectionchange?: () => void;
    /** Called after the canvas itself mutates the doc (typing / Enter / paste). */
    onedit?: () => void;
  }

  let { model, version, zoom = 1, onselectionchange, onedit }: Props = $props();

  let canvas = $state<HTMLDivElement | null>(null);
  // Bumped by structural edits made *inside* the canvas (Enter / Backspace /
  // undo / paste) to force a re-render + caret restore, without the parent
  // having to bump `version`. Plain typing does NOT bump this — the browser
  // owns the caret during input and re-rendering mid-keystroke would reset it.
  let renderTick = $state(0);

  /** Re-render after a structural edit the canvas performed itself. */
  function rerender(): void {
    renderTick++;
  }

  // Re-render the canvas HTML only when `version` changes (a structural/format
  // edit), never on every keystroke — that would reset the caret. Character
  // typing is reconciled back into the AST by syncFromDom() on input. After a
  // structural re-render we restore the caret to the model's selection so edits
  // like Enter / Backspace / undo don't dump the caret at the top.
  $effect(() => {
    if (!canvas) return;
    // Depend on both the parent's structural `version` and the canvas-local
    // `renderTick` so either source of structural change re-renders.
    canvas.dataset.version = `${version}.${renderTick}`;
    canvas.innerHTML = renderDocumentHtml(model.doc);
    restoreCaret();
  });

  /** The run span (or paragraph fallback) that hosts a document position. */
  function hostFor(pos: DocPosition): HTMLElement | null {
    if (!canvas) return null;
    const sel = [`.wk-run[data-wk-block="${pos.block}"]`];
    if (pos.cell) sel.push(`[data-wk-cell="${pos.cell.row},${pos.cell.col}"]`);
    else sel.push(':not([data-wk-cell])');
    sel.push(`[data-wk-inline="${pos.inline ?? 0}"]`);
    const span = canvas.querySelector<HTMLElement>(sel.join(''));
    if (span) return span;
    // Empty paragraph: no run span, caret lives in the <p> itself.
    const pSel = pos.cell
      ? `.wk-p[data-wk-block="${pos.block}"][data-wk-cell="${pos.cell.row},${pos.cell.col}"]`
      : `.wk-p[data-wk-block="${pos.block}"]:not([data-wk-cell])`;
    return canvas.querySelector<HTMLElement>(pSel);
  }

  /** Map a document position to a concrete DOM (textNode, offset) caret point. */
  function domPoint(pos: DocPosition): { node: Node; offset: number } | null {
    const host = hostFor(pos);
    if (!host) return null;
    const text = host.firstChild;
    if (text && text.nodeType === Node.TEXT_NODE) {
      const raw = text.textContent ?? '';
      // The renderer uses a zero-width space as an empty placeholder; clamp the
      // caret to real content length so it lands correctly.
      const len = raw.replace(/​/g, '').length || raw.length;
      return { node: text, offset: Math.min(pos.offset ?? 0, len) };
    }
    return { node: host, offset: 0 };
  }

  function restoreCaret(): void {
    const sel = model.selection;
    if (!sel) return;
    const focus = domPoint(sel.focus);
    if (!focus) return;
    const anchor = domPoint(sel.anchor) ?? focus;
    const range = document.createRange();
    range.setStart(anchor.node, anchor.offset);
    range.setEnd(focus.node, focus.offset);
    const domSel = window.getSelection();
    if (!domSel) return;
    domSel.removeAllRanges();
    domSel.addRange(range);
  }

  /** Push edited run text from the DOM back into the AST (no re-render). */
  function syncFromDom(): void {
    if (!canvas) return;
    const spans = canvas.querySelectorAll<HTMLElement>('.wk-run[data-wk-block]');
    for (const span of spans) {
      const block = Number.parseInt(span.getAttribute('data-wk-block') ?? '', 10);
      if (Number.isNaN(block)) continue;
      const cellAttr = span.getAttribute('data-wk-cell');
      const inlineAttr = span.getAttribute('data-wk-inline');
      const inline = inlineAttr ? Number.parseInt(inlineAttr, 10) : 0;
      let cell: { row: number; col: number } | undefined;
      if (cellAttr) {
        const [r, c] = cellAttr.split(',').map((n) => Number.parseInt(n, 10));
        if (r !== undefined && c !== undefined && !Number.isNaN(r) && !Number.isNaN(c))
          cell = { row: r, col: c };
      }
      const pos = { block, inline, ...(cell ? { cell } : {}) };
      const run = runAtPath(model.doc, pos);
      if (!run) continue;
      const text = (span.textContent ?? '').replace(/​/g, '');
      setSimpleRunText(run, text);
    }
    model.doc.dirty = true;
  }

  function onInput(): void {
    syncFromDom();
    onedit?.();
  }

  function onSelChange(): void {
    const sel = readDomSelection(document);
    if (sel) model.setSelection(sel);
    onselectionchange?.();
  }

  /** Whether the caret is a collapsed cursor at the very start of its paragraph. */
  function caretAtParagraphStart(): boolean {
    const sel = model.selection;
    if (!sel) return false;
    const domSel = window.getSelection();
    if (!domSel || !domSel.isCollapsed) return false;
    return (sel.focus.inline ?? 0) === 0 && (sel.focus.offset ?? 0) === 0;
  }

  /** Whether the caret is a collapsed cursor at the very end of its paragraph. */
  function caretAtParagraphEnd(): boolean {
    const domSel = window.getSelection();
    if (!domSel || !domSel.isCollapsed || !canvas) return false;
    const pos = model.selection?.focus;
    if (!pos) return false;
    const host = hostFor(pos);
    if (!host) return false;
    // Last run span in the same paragraph?
    const runs = canvas.querySelectorAll<HTMLElement>(
      pos.cell
        ? `.wk-run[data-wk-block="${pos.block}"][data-wk-cell="${pos.cell.row},${pos.cell.col}"]`
        : `.wk-run[data-wk-block="${pos.block}"]:not([data-wk-cell])`,
    );
    const isLast = runs.length === 0 || runs[runs.length - 1] === host;
    const len = (host.textContent ?? '').replace(/​/g, '').length;
    return isLast && (pos.offset ?? 0) >= len;
  }

  function onKeydown(e: KeyboardEvent): void {
    const mod = e.ctrlKey || e.metaKey;

    // Undo / redo.
    if (mod && e.key.toLowerCase() === 'z') {
      e.preventDefault();
      syncFromDom();
      if (e.shiftKey) model.redo();
      else model.undo();
      rerender();
      onedit?.();
      return;
    }
    if (mod && e.key.toLowerCase() === 'y') {
      e.preventDefault();
      syncFromDom();
      model.redo();
      rerender();
      onedit?.();
      return;
    }

    // Enter splits the paragraph at the caret (Shift+Enter = soft line break).
    if (e.key === 'Enter') {
      e.preventDefault();
      syncFromDom();
      if (e.shiftKey) runCommand(model, commands.insertLineBreakCommand, { kind: 'line' });
      else runCommand(model, commands.splitParagraphCommand, undefined);
      rerender();
      onedit?.();
      return;
    }

    // Backspace at paragraph start merges into the previous paragraph.
    if (e.key === 'Backspace' && caretAtParagraphStart() && (model.selection?.focus.block ?? 0) > 0) {
      e.preventDefault();
      syncFromDom();
      runCommand(model, commands.mergeBackCommand, undefined);
      rerender();
      onedit?.();
      return;
    }

    // Delete at paragraph end pulls the next paragraph up into this one.
    if (e.key === 'Delete' && caretAtParagraphEnd()) {
      const here = model.selection?.focus;
      if (here && canvas) {
        const nextBlock = here.block + 1;
        const nextExists = !!canvas.querySelector(`[data-wk-block="${nextBlock}"]`);
        if (nextExists) {
          e.preventDefault();
          syncFromDom();
          // Merge the next paragraph into this one, keeping the caret at the join.
          model.setSelection({ anchor: { block: nextBlock }, focus: { block: nextBlock } });
          runCommand(model, commands.mergeBackCommand, undefined);
          rerender();
          onedit?.();
          return;
        }
      }
    }

    // Formatting shortcuts.
    if (mod && ['b', 'i', 'u'].includes(e.key.toLowerCase())) {
      e.preventDefault();
      syncFromDom();
      const cmd =
        e.key.toLowerCase() === 'b'
          ? commands.toggleBoldCommand
          : e.key.toLowerCase() === 'i'
            ? commands.toggleItalicCommand
            : commands.toggleUnderlineCommand;
      runCommand(model, cmd, undefined);
      rerender();
      onedit?.();
    }
  }

  function onPaste(e: ClipboardEvent): void {
    const text = e.clipboardData?.getData('text/plain');
    if (text === undefined) return;
    e.preventDefault();
    syncFromDom();
    const lines = text.split(/\r?\n/);
    // First line: insert into the caret run at the offset.
    const pos = model.selection?.focus;
    if (!pos) return;
    const run = runAtPath(model.doc, pos);
    if (run) {
      const current = run.pieces
        .filter((p): p is { kind: 'text'; value: string; preserveSpace: boolean } => p.kind === 'text')
        .map((p) => p.value)
        .join('');
      const at = pos.offset ?? 0;
      const merged = current.slice(0, at) + lines[0] + current.slice(at);
      setSimpleRunText(run, merged);
      model.setSelection({
        anchor: { ...pos, offset: at + (lines[0]?.length ?? 0) },
        focus: { ...pos, offset: at + (lines[0]?.length ?? 0) },
      });
    }
    model.doc.dirty = true;
    // Remaining lines become new paragraphs.
    for (let i = 1; i < lines.length; i++) {
      runCommand(model, commands.splitParagraphCommand, undefined);
      const np = model.selection?.focus;
      const npr = np ? runAtPath(model.doc, np) : undefined;
      if (npr && lines[i]) {
        setSimpleRunText(npr, lines[i]!);
        model.setSelection({
          anchor: { ...np!, offset: lines[i]!.length },
          focus: { ...np!, offset: lines[i]!.length },
        });
      }
    }
    rerender();
    onedit?.();
  }

  $effect(() => {
    document.addEventListener('selectionchange', onSelChange);
    return () => document.removeEventListener('selectionchange', onSelChange);
  });
</script>

<div class="wk-page" style="--zoom: {zoom}">
  <div
    bind:this={canvas}
    class="wk-canvas"
    contenteditable="true"
    role="textbox"
    tabindex="0"
    aria-multiline="true"
    aria-label="Document editor"
    oninput={onInput}
    onkeydown={onKeydown}
    onpaste={onPaste}
  ></div>
</div>

<style>
  .wk-page {
    display: flex;
    justify-content: center;
    transform: scale(var(--zoom));
    transform-origin: top center;
  }
  .wk-canvas {
    background: #fff;
    color: #111;
    width: 100%;
    max-width: 816px;
    min-height: 1056px;
    margin: 0 auto;
    padding: 96px 96px;
    box-shadow:
      0 1px 4px rgba(0, 0, 0, 0.12),
      0 8px 24px rgba(0, 0, 0, 0.08);
    border-radius: 2px;
    outline: none;
    font-family: 'Calibri', 'Segoe UI', system-ui, sans-serif;
    font-size: 11pt;
    line-height: 1.5;
  }

  .wk-canvas :global(.wk-p) {
    margin: 0 0 8px;
    min-height: 1.4em;
  }

  .wk-canvas :global(.wk-table) {
    border-collapse: collapse;
    margin: 8px 0;
    width: 100%;
  }

  .wk-canvas :global(.wk-td) {
    border: 1px solid #bbb;
    padding: 4px 8px;
    vertical-align: top;
  }

  .wk-canvas :global(.wk-raw) {
    color: #888;
    font-style: italic;
    user-select: none;
    background: #f3f4f6;
    padding: 2px 6px;
    border-radius: 3px;
    display: inline-block;
  }
</style>
