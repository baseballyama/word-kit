// Injected at build time via the `__DOCX_KIT_VERSION__` define in
// tsdown.config.ts. The literal fallback keeps vitest (which has no define)
// and any un-bundled consumer working with a recognizable dev sentinel.
// oxlint-disable-next-line no-underscore-dangle -- build-time define injected by tsdown
declare const __DOCX_KIT_VERSION__: string;

export const VERSION: string =
  typeof __DOCX_KIT_VERSION__ === "string" ? __DOCX_KIT_VERSION__ : "0.0.0-dev";
