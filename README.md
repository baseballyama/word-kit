# word-kit

OOXML-compliant (ECMA-376) **`.docx`** generation library. Runs in modern
browsers and Node.js.

- Build a Word document from scratch with a programmatic API.
- Open an existing `.docx` as a template, edit it, and serialize it back.
- Target: **full WordprocessingML coverage** of ECMA-376 Part 1.

> Status: pre-1.0. Core functionality (`openDocx` / `createDocx` /
> `toUint8Array`) is stable; details may shift before v1.
> See [`CLAUDE.md`](./CLAUDE.md) for engineering principles and scope rules
> and [`PLAN.md`](./PLAN.md) for the milestone roadmap.

## Install

```bash
pnpm add @word-kit/core
# or: npm install @word-kit/core / yarn add @word-kit/core
```

`@word-kit/core` ships ESM-only with bundled `.d.ts` types. It has no
Node-only dependencies and works in browsers (including with the `Blob` and
`File` APIs).

## Quick start

```ts
import {
  addBulletList,
  addTable,
  appendParagraph,
  createDocx,
  openDocx,
  PAGE_SIZE_A4,
  replaceTextEverywhere,
  setPageSize,
  toUint8Array,
} from "@word-kit/core";

// From scratch
const doc = createDocx({ paragraphs: ["Hello, world."] });
appendParagraph(doc, "Bullets:");
addBulletList(doc, ["one", "two", "three"]);
addTable(doc, [
  ["Name", "Score"],
  ["Alice", "90"],
]);
setPageSize(doc, PAGE_SIZE_A4);
const bytes = toUint8Array(doc);

// From an existing .docx template
const tpl = openDocx(existingDocxBytes);
replaceTextEverywhere(tpl, /\{\{(\w+)\}\}/g, (m) => values[m.captures[0]!] ?? "");
const out = toUint8Array(tpl);
```

> **Why standalone functions instead of methods?** Each operation is its
> own export, so a bundler can tree-shake any function you don't import.
> A minimal `createDocx + appendParagraph + toUint8Array` slice bundles to
> ~40 KB minified; the full surface is ~115 KB. CI enforces both numbers.

See [`packages/core/README.md`](./packages/core/README.md) for the full API
walkthrough (images, comments, footnotes, headers/footers, bookmarks,
hyperlinks, tracked changes, core document properties).

## Scope

**In scope**

- WordprocessingML (`.docx`) — read, edit, write.
- OPC packaging (ECMA-376 Part 2), DrawingML — as the underlying layers.

**Out of scope (for now)**

- `.pptx` (PresentationML) and `.xlsx` (SpreadsheetML).

**Out of scope (permanent)**

- Rendering to PDF / HTML, headless Word automation, binary `.doc` (pre-2007).

If a feature request only makes sense for pptx / xlsx, or for rendering, it
will be redirected to a more appropriate library. See [`CLAUDE.md`](./CLAUDE.md)
("Scope discipline").

## Repository layout

```
.
├── packages/
│   ├── core/             # @word-kit/core      — public Docx wrapper
│   ├── opc/              # @word-kit/opc       — OPC (ZIP + Content Types + rels)
│   ├── ooxml-xml/        # @word-kit/ooxml-xml — namespace-aware XML parser/serializer
│   └── wordprocessingml/ # @word-kit/wml       — WML AST + parsers + builders
├── docs/
│   └── specs/                # Distilled spec notes + ECMA-376 fetcher target
├── references/               # External OSS / spec material (submodules)
├── scripts/
│   └── fetch-specs.sh        # Downloads ECMA-376 PDFs + XSDs into docs/specs/
├── .changeset/               # Changesets — drives versioning + npm release
├── .github/                  # Issue / PR templates and CI / release workflows
├── .claude/skills/           # Workflow guides for Claude Code agents
├── CLAUDE.md                 # Engineering principles for contributors and AI
├── PLAN.md                   # Living implementation plan + progress table
├── SECURITY.md               # Private vulnerability reporting
├── .oxlintrc.json            # oxlint config
└── .oxfmtrc.json             # oxfmt config (Prettier-compatible)
```

## Development

```bash
pnpm install           # one-time setup
pnpm typecheck         # tsc --noEmit across all packages
pnpm lint              # oxlint
pnpm format:check      # oxfmt --check
pnpm test              # vitest run
pnpm build             # tsup, all packages
```

See [`.claude/skills/run-check-and-test/SKILL.md`](./.claude/skills/run-check-and-test/SKILL.md)
for the canonical pre-PR quality gate.

## Contributing

- Issues and PRs **must** follow the templates under
  [`.github/`](./.github/). Submissions that strip the template structure are
  auto-closed by the template-compliance workflow.
- Read [`CLAUDE.md`](./CLAUDE.md) before opening a PR — it covers the
  "one way to do one thing" rule, defensive programming, comments, and the
  hard "no"s (no `as unknown as T`, no N+1, etc.).

## License

[MIT](./LICENSE)
