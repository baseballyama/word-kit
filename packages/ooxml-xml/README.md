# @word-kit/ooxml-xml

Namespace-aware XML parser and serializer used by word-kit.

> **Internal package.** The public entry point for word-kit is
> [`@word-kit/core`](../core/README.md). This package is documented for
> contributors only; its API may change without notice.

## What it does

OOXML files are XML, but not all XML libraries preserve the things OOXML
round-trips care about:

- **Attribute order** — Word and other consumers are sometimes order-sensitive.
- **Namespace prefixes** — `<w:document>` and `<w:body>` are conventional; the
  source prefix should round-trip back unchanged.
- **Whitespace** — `xml:space="preserve"` on `<w:t>` is load-bearing.
- **Comments and processing instructions** — should round-trip even if word-kit
  does not interpret them.

This package implements a small purpose-built parser/serializer that preserves
all of the above, so higher layers (AST builder, OPC writer) can rely on
faithful round-trips.

## License

[MIT](../../LICENSE)
