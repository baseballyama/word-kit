// Render any Docx — built from scratch or opened from bytes — into a
// DOM container using @word-kit/preview. The wrap is intentional: we
// share the renderer with docx-preview upstream but pin the contract
// behind word-kit's stable function-API surface.

import { openDocx } from "@word-kit/core";
import { previewToDOM } from "@word-kit/preview";

declare const bytes: Uint8Array;
declare const container: HTMLElement;

const doc = openDocx(bytes);
const handle = await previewToDOM(doc, container, {
  classPrefix: "wk-",
  inWrapper: true,
  breakPages: true,
  renderFonts: true,
});

// when you're done, detach + release internal references:
handle.dispose();
