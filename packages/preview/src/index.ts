/**
 * `@word-kit/preview` — render a `.docx` produced (or opened) by
 * `@word-kit/core` into a browser DOM container so callers can show a
 * read-only preview of the document next to a form / editor / pipeline
 * that produced it.
 *
 * @packageDocumentation
 *
 * **Implementation status:** v0 = bridge over the OSS `docx-preview`
 * library (Apache-2.0). The public API ({@link previewToDOM}) is the
 * stable surface; v1 will swap the implementation to render directly
 * from the WML AST without re-parsing the OOXML. See
 * `docs/PLAN-PREVIEW.md` in the monorepo for the design.
 */

import { type Docx, toUint8Array } from "@word-kit/core";

/**
 * Options accepted by {@link previewToDOM}. Names mirror the
 * underlying renderer where they map directly so callers familiar with
 * `docx-preview` see no surprises.
 */
export interface PreviewToDOMOptions {
  /**
   * Prefix applied to every CSS class the renderer emits (so multiple
   * previews on the same page can't collide with each other or the
   * host's stylesheet). Defaults to `"wk-"`.
   */
  readonly classPrefix?: string;
  /**
   * When `true` (default), wrap each rendered page in a
   * `.wk-document` block so the host can scroll a multi-page preview.
   * When `false`, content is appended flat into the container.
   */
  readonly inWrapper?: boolean;
  /**
   * When `true` (default), insert source-declared page breaks
   * (`<w:br w:type="page"/>` and `<w:lastRenderedPageBreak/>`) as
   * visible page boundaries. When `false`, the document renders as
   * one long flow.
   */
  readonly breakPages?: boolean;
  /**
   * When `true` (default), inject `@font-face` declarations from the
   * document's embedded font table. When `false`, falls back to the
   * host's available fonts.
   */
  readonly renderFonts?: boolean;
  /**
   * When `true`, render comments alongside the body (left-margin
   * markers). Defaults to `false` because the implementation is still
   * experimental upstream.
   */
  readonly experimentalComments?: boolean;
  /**
   * When `true`, render tracked-change markup. Defaults to `false`
   * because the implementation is still experimental upstream.
   */
  readonly experimentalChanges?: boolean;
}

/**
 * Handle returned by {@link previewToDOM}. Calling {@link Handle.dispose}
 * detaches the renderer's CSS, removes the rendered DOM nodes from the
 * container, and releases internal references.
 *
 * @remarks Plain data shape — no class. Matches `@word-kit/core`'s
 * function-API posture.
 */
export interface Handle {
  /**
   * Detach the renderer's CSS and remove the rendered DOM. Safe to
   * call more than once; subsequent calls are no-ops.
   */
  dispose(): void;
}

/** Source the caller hands to {@link previewToDOM}. */
export type PreviewSource = Docx | Uint8Array | Blob | ArrayBuffer;

const DEFAULT_OPTIONS = {
  classPrefix: "wk-",
  inWrapper: true,
  breakPages: true,
  renderFonts: true,
  experimentalComments: false,
  experimentalChanges: false,
} satisfies Required<PreviewToDOMOptions>;

/**
 * Render `source` into `container` as a read-only preview.
 *
 * Accepts either a `Docx` value (just-built or just-opened) or raw
 * bytes (`Uint8Array` / `Blob` / `ArrayBuffer`). The promise resolves
 * after the renderer finishes painting; call `handle.dispose()` to
 * tear down when the preview is no longer needed.
 *
 * @example
 * ```ts
 * import { openDocx } from "@word-kit/core";
 * import { previewToDOM } from "@word-kit/preview";
 *
 * const doc = openDocx(bytes);
 * const handle = await previewToDOM(doc, document.querySelector("#preview")!);
 * // …later, when the preview is no longer needed:
 * handle.dispose();
 * ```
 *
 * @remarks
 * v0 implementation is a thin bridge over the OSS `docx-preview`
 * library. v1 will swap in a renderer driven directly from the WML
 * AST (no second OOXML parse). The signature on this function will
 * not change across the swap.
 */
export async function previewToDOM(
  source: PreviewSource,
  container: HTMLElement,
  options: PreviewToDOMOptions = {},
): Promise<Handle> {
  const opts = { ...DEFAULT_OPTIONS, ...options };
  const bytes = await coerceToBytes(source);

  // Lazy-import `docx-preview` so consumers who only import types from
  // this package don't pull the renderer into their bundle prematurely.
  const docxPreview = await import("docx-preview");

  // Allocate a sub-container so dispose() can wipe just our output and
  // leave anything the host appended around it untouched.
  const root = container.ownerDocument.createElement("div");
  root.setAttribute("data-word-kit-preview", "");
  container.appendChild(root);

  const styleHost = container.ownerDocument.createElement("div");
  styleHost.setAttribute("data-word-kit-preview-styles", "");
  root.appendChild(styleHost);

  const bodyHost = container.ownerDocument.createElement("div");
  bodyHost.setAttribute("data-word-kit-preview-body", "");
  root.appendChild(bodyHost);

  await docxPreview.renderAsync(bytes, bodyHost, styleHost, {
    className: opts.classPrefix,
    inWrapper: opts.inWrapper,
    breakPages: opts.breakPages,
    ignoreFonts: !opts.renderFonts,
    experimental: opts.experimentalComments || opts.experimentalChanges,
    renderComments: opts.experimentalComments,
    renderChanges: opts.experimentalChanges,
  });

  let disposed = false;
  return {
    dispose(): void {
      if (disposed) return;
      disposed = true;
      root.remove();
    },
  };
}

async function coerceToBytes(source: PreviewSource): Promise<Uint8Array> {
  if (source instanceof Uint8Array) return source;
  if (source instanceof ArrayBuffer) return new Uint8Array(source);
  if (typeof Blob !== "undefined" && source instanceof Blob) {
    const buf = await source.arrayBuffer();
    return new Uint8Array(buf);
  }
  // Treat the remaining case as `Docx`. It's the function-API surface
  // from `@word-kit/core` so we can reach `toUint8Array` without a
  // type assertion.
  return toUint8Array(source as Docx);
}
