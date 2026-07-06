/**
 * `@office-kit/docx-editor` — an MS Office-like WYSIWYG editing core for
 * word-kit `.docx` documents.
 *
 * The editor never serializes OOXML itself: every mutation runs through a
 * {@link Command} that calls the `@office-kit/docx` public API, keeping that
 * library the single source of truth. Completeness against the WordprocessingML
 * spec is tracked and enforced by the capability ledger under `./coverage`.
 *
 * @packageDocumentation
 */

import { createDocx, type Docx, openDocx } from "@office-kit/docx";
import { EditorModel } from "./model.js";

export { EditorModel, type ChangeListener, type EditorSnapshot } from "./model.js";
export { type Command, type FeatureGroup, runCommand } from "./commands/types.js";
export { ALL_COMMANDS, COMMAND_IDS, commandsInGroup, getCommand } from "./commands/registry.js";
// Named command objects + group arrays for typed UI wiring.
export * as commands from "./commands/index.js";
export {
  type CellCoord,
  type DocPosition,
  type OrderedSelection,
  type Selection,
  caretAt,
  orderSelection,
} from "./selection.js";
export { renderDocumentHtml, paragraphPlainText } from "./render.js";
export { positionFromDom, readDomSelection } from "./dom-selection.js";
export { runAtPath, setSimpleRunText } from "./text-edit.js";
export { type RawNode, rawTrees, allRawElements, xmlParts, partRawTree } from "./raw-tree.js";
export { blocks, paragraphAt, paragraphsInRange, runsInRange } from "./doc-access.js";

// Coverage / capability ledger — the completeness-guarantee surface.
export {
  type Capability,
  type Disposition,
  classify,
  coverageSummary,
  LEDGER,
} from "./capability/ledger.js";
export { renderCoverageMarkdown } from "./capability/report.js";

/** Create an editor over a fresh, empty document. */
export function createEditor(): EditorModel {
  return new EditorModel(createDocx({ paragraphs: [""] }));
}

/** Create an editor over an existing `.docx` (bytes). */
export function openEditor(bytes: Uint8Array): EditorModel {
  return new EditorModel(openDocx(bytes));
}

/** Create an editor over an already-open {@link Docx}. */
export function editorFor(doc: Docx): EditorModel {
  return new EditorModel(doc);
}
