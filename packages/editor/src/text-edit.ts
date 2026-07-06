/**
 * Low-level text-sync helpers used by the canvas to push typed text back into
 * the AST. The canvas renders each run as an editable span; on input it reads
 * the span's text and writes it to the matching run here.
 */

import type { Docx, WmlRun } from "@office-kit/docx";
import { paragraphAt } from "./doc-access.js";
import type { DocPosition } from "./selection.js";

/** Resolve the run addressed by a position's `inline` index. */
export function runAtPath(doc: Docx, pos: DocPosition): WmlRun | undefined {
  const para = paragraphAt(doc, pos);
  if (!para) return undefined;
  const runs = para.children.filter((c): c is WmlRun => c.kind === "run");
  return runs[pos.inline ?? 0];
}

/**
 * Replace a run's text when it is a simple text-only run. Returns false (and
 * changes nothing) for runs carrying tabs, breaks, drawings, fields, etc. — the
 * canvas leaves those to command-level edits so nothing is silently dropped.
 */
export function setSimpleRunText(run: WmlRun, text: string): boolean {
  const nonText = run.pieces.some((p) => p.kind !== "text");
  if (nonText && run.pieces.length > 0) return false;
  const preserveSpace = /^\s|\s$|\s\s/.test(text);
  run.pieces = [{ kind: "text", value: text, preserveSpace }];
  return true;
}
