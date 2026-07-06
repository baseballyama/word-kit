/**
 * Command layer types.
 *
 * A command is the *only* way the editor mutates a document. Each command's
 * `run` calls the `@office-kit/docx` public API through {@link EditorModel},
 * wrapping the mutation in the model's undo snapshot. The set of command ids is
 * the single source of truth the coverage ledger validates against: a
 * WordprocessingML element may be classified `edit` only if a real command id
 * here handles it.
 */

import type { EditorModel } from "../model.js";

/** Ribbon groupings, mirroring Word's tabs. */
export type FeatureGroup =
  | "text"
  | "paragraph"
  | "structure"
  | "list"
  | "table"
  | "image"
  | "style"
  | "section"
  | "headerFooter"
  | "references"
  | "review"
  | "advanced";

/**
 * A command definition. `run` receives the model plus command-specific params;
 * it must call {@link EditorModel.beginEdit} before mutating and
 * {@link EditorModel.commit} after (the {@link runCommand} helper does this).
 */
export interface Command<P = void> {
  readonly id: string;
  readonly group: FeatureGroup;
  /** Human label for the UI (English; the UI layer localizes). */
  readonly label: string;
  /** Longer description for tooltips / command palette. */
  readonly description?: string;
  /** Perform the mutation. Return value is ignored; report failures by throwing. */
  run(model: EditorModel, params: P): void;
  /** Whether the command applies to the current selection/state. */
  isEnabled?(model: EditorModel): boolean;
  /** For toggle commands (bold, italic…): whether it is currently on. */
  isActive?(model: EditorModel): boolean;
}

/** Run a command with automatic undo-snapshot + commit around the mutation. */
export function runCommand<P>(model: EditorModel, command: Command<P>, params: P): void {
  if (command.isEnabled && !command.isEnabled(model)) return;
  model.beginEdit();
  try {
    command.run(model, params);
  } catch (err) {
    // The mutation failed after we snapshotted; roll back so the model is not
    // left half-edited, then rethrow so the failure surfaces (anti-swallow).
    model.undo();
    throw err;
  }
  model.commit();
}
