# @word-kit/wml

WordprocessingML semantic AST and parser/writer for word-kit.

> **Internal package.** The public entry point for word-kit is
> [`@word-kit/core`](../core/README.md).

## Scope

Maps `word/document.xml` (and related parts) between raw XML AST (from
[`@word-kit/ooxml-xml`](../ooxml-xml)) and a typed Word document model:

- `WmlDocument` → `WmlBody` → blocks (paragraphs, tables, …)
- `WmlParagraph` → runs and other inlines
- `WmlRun` → text pieces, breaks, tabs, drawings, …

Unknown elements are kept as raw XML nodes so that nothing is lost on
round-trip. Higher-fidelity property structures (typed `pPr`/`rPr`) land
incrementally; see `PLAN.md`.

## License

[MIT](../../LICENSE)
