# @office-kit/docx-editor

An MS Office-like WYSIWYG editing core for [`@office-kit/docx`](https://github.com/office-kit/docx)
`.docx` documents. Framework-agnostic; a Svelte/SvelteKit UI lives in the
monorepo's `site/` (`/editor` route).

## Design

The editor never serializes OOXML itself. Every mutation runs through a
**command** that calls the `@office-kit/docx` public API, so that library stays
the single source of truth for the document. Open a `.docx`, drive commands from
a ribbon or keyboard, and save a real, valid `.docx` back.

```ts
import { openEditor, runCommand, commands, renderDocumentHtml } from "@office-kit/docx-editor";
import { toUint8Array } from "@office-kit/docx";

const model = openEditor(bytes); // EditorModel over a live Docx
model.setSelection({ anchor: { block: 0 }, focus: { block: 0 } });
runCommand(model, commands.toggleBoldCommand, undefined);
container.innerHTML = renderDocumentHtml(model.doc);
const out = toUint8Array(model.doc); // a valid .docx
```

## The completeness guarantee

"Cover every docx representation" is a _checkable_ claim here, not an aspiration.
Every element in the ECMA-376 schema universe (extracted into
`src/capability/element-universe.json`) is classified in the capability ledger
(`src/capability/ledger.ts`) into exactly one disposition:

- **`edit`** — a real command creates/modifies it (the command id is recorded).
- **`render`** — the canvas shows it faithfully, read-only.
- **`preserve`** — round-tripped losslessly by `@office-kit/docx` (raw XML
  pass-through), not yet surfaced in the UI.

`src/capability/coverage.test.ts` fails the build if any element is unclassified,
or if an `edit` element points at a command that no longer exists. The live
scoreboard is regenerated into [`COVERAGE.md`](./COVERAGE.md). Nothing can
silently fall through: an element is either editable, rendered, or explicitly
preserved with the round-trip test that proves it.

Regenerate the universe from the schemas with:

```sh
node scripts/extract-ooxml-elements.mjs
```

## Command groups

`text`, `paragraph`, `structure`, `list`, `table`, `image`, `style`, `section`,
`headerFooter`, `references`, `review`. See `ALL_COMMANDS` for the full list.
