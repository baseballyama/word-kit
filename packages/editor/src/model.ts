/**
 * The editor's stateful controller.
 *
 * An editor is inherently stateful (current document, selection, undo history,
 * subscribers), so unlike `@office-kit/docx`'s function API this is modelled as
 * a small class. It never serializes OOXML itself: every document mutation is
 * performed by a {@link Command} that calls the `@office-kit/docx` public API.
 * The library remains the single source of truth for the document ("one way to
 * do one thing").
 */

import { clone, type Docx } from "@office-kit/docx";
import type { Selection } from "./selection.js";

export interface EditorSnapshot {
  readonly doc: Docx;
  readonly selection: Selection | null;
}

export type ChangeListener = (model: EditorModel) => void;

/** How deep the undo/redo history is allowed to grow. */
const DEFAULT_HISTORY_LIMIT = 200;

export class EditorModel {
  private docState: Docx;
  private selectionState: Selection | null = null;
  private readonly undoStack: EditorSnapshot[] = [];
  private readonly redoStack: EditorSnapshot[] = [];
  private readonly listeners = new Set<ChangeListener>();
  private readonly historyLimit: number;

  constructor(doc: Docx, options: { historyLimit?: number } = {}) {
    this.docState = doc;
    this.historyLimit = options.historyLimit ?? DEFAULT_HISTORY_LIMIT;
  }

  get doc(): Docx {
    return this.docState;
  }

  get selection(): Selection | null {
    return this.selectionState;
  }

  setSelection(selection: Selection | null): void {
    this.selectionState = selection;
    this.emit();
  }

  /**
   * Snapshot the current document onto the undo stack before a mutating command
   * runs. Redo history is cleared because a new edit forks the timeline. The
   * snapshot clones the document so later mutations don't alias it.
   */
  beginEdit(): void {
    this.undoStack.push({ doc: clone(this.docState), selection: this.selectionState });
    if (this.undoStack.length > this.historyLimit) this.undoStack.shift();
    this.redoStack.length = 0;
  }

  /** Notify subscribers that the document (or selection) changed. */
  commit(nextSelection?: Selection | null): void {
    if (nextSelection !== undefined) this.selectionState = nextSelection;
    this.docState.dirty = true;
    this.emit();
  }

  canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  undo(): void {
    const prev = this.undoStack.pop();
    if (!prev) return;
    this.redoStack.push({ doc: clone(this.docState), selection: this.selectionState });
    this.docState = prev.doc;
    this.selectionState = prev.selection;
    this.emit();
  }

  redo(): void {
    const next = this.redoStack.pop();
    if (!next) return;
    this.undoStack.push({ doc: clone(this.docState), selection: this.selectionState });
    this.docState = next.doc;
    this.selectionState = next.selection;
    this.emit();
  }

  /** Replace the whole document (e.g. after opening a new file). Clears history. */
  replaceDocument(doc: Docx): void {
    this.docState = doc;
    this.selectionState = null;
    this.undoStack.length = 0;
    this.redoStack.length = 0;
    this.emit();
  }

  subscribe(listener: ChangeListener): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  private emit(): void {
    for (const listener of this.listeners) listener(this);
  }
}
