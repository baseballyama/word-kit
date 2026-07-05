# Contributing to word-kit

Thanks for your interest in contributing. This document covers the day-to-day
workflow: how to set up a development environment, run the test suite, and
land a change.

## Reporting issues

- **Bug reports** — please use the bug-report template under
  [`.github/ISSUE_TEMPLATE/`](.github/ISSUE_TEMPLATE/). Include the
  word-kit version, a minimal reproduction, and the offending `.docx` (or at
  least the snippet that triggers it). Bugs in docx output ("Word says the
  file needs to be repaired") almost always need the file attached to be
  actionable.
- **Feature requests** — use the feature-request template. Tell us the
  use case first; the API shape we land on usually depends on it.
- **Security issues** — please do **not** file a public issue. Follow the
  process in [`SECURITY.md`](SECURITY.md).

## Development environment

```sh
# Requires Node 22+ and pnpm.
git clone --recurse-submodules https://github.com/office-kit/docx
cd docx
pnpm install
```

The `--recurse-submodules` flag pulls the [`mammoth.js`](references/mammoth.js)
and [`python-docx`](references/python-docx) test corpora into `references/`;
the fixture-coverage and fixture-python-docx tests load `.docx` files from
those checkouts. If you clone without submodules the corresponding test
files will skip rather than fail, but you'll miss real-world round-trip
coverage.

### Useful scripts

| Command                 | What it does                                                                 |
| ----------------------- | ---------------------------------------------------------------------------- |
| `pnpm -r run typecheck` | Run `tsc --noEmit` across the library + preview with the strict settings.    |
| `pnpm lint`             | Run `oxlint`.                                                                |
| `pnpm format`           | Run `oxfmt` (writes). `pnpm format:check` validates without writing.         |
| `pnpm build`            | Bundle the `@office-kit/docx` library with `tsdown` (rolldown) into `dist/`. |
| `pnpm build:all`        | Build every workspace package (library, preview, site).                      |
| `pnpm test`             | Run vitest — 460 tests, < 20 s. No build step needed (source-resolved).      |
| `pnpm test:watch`       | Same in watch mode.                                                          |
| `pnpm check:tree-shake` | Build the minimal entry and budget it against the full surface (~42/131 KB). |
| `pnpm sample`           | Write demonstration `.docx` files into `./samples/` for inspection.          |
| `pnpm changeset`        | Add a changeset describing your change (drives release notes + versioning).  |

CI mirrors the same gate: format check + lint + build + typecheck +
tree-shake budget on Node 22, then a vitest run on Node 22 and 24. Run them
locally before opening a PR to catch most failures up front.

## Workflow

1. **Open or claim an issue.** For non-trivial work (new public API,
   breaking change, large refactor), please discuss the shape in an issue
   first so we can align before code is written.
2. **Branch from `main`.** Keep branches focused — one logical change per
   branch makes review and reverts straightforward.
3. **Write a test.** Bug fixes should add a regression test; features
   should ship with unit + integration coverage. The test files live next
   to their source under `src/` (and `packages/preview/src/`).
4. **Run the full gate locally.** `pnpm format:check && pnpm lint && pnpm -r
run typecheck && pnpm test && pnpm check:tree-shake` is what CI runs.
5. **Add a changeset.** word-kit uses
   [Changesets](https://github.com/changesets/changesets) for release notes
   and version bumps. Run `pnpm changeset` and pick `patch` for fixes,
   `minor` for additive features, `major` only after coordination with the
   maintainer. Pre-1.0 we still try to call out breaking changes in `minor`
   bumps clearly.
6. **Open a PR.** Use the template; fill in what the change does and how
   you tested it. Link the issue you're closing.

## Code style

The discipline lives in [`CLAUDE.md`](CLAUDE.md). The short version:

- **TypeScript strict mode is non-negotiable.** No `as any`, no
  `@ts-ignore`, no `as unknown as T` outside the few irreducible
  `Buffer ↔ Uint8Array` / WHATWG bridges.
- **Function API, no classes.** Every operation is a standalone, named
  export that takes a `Docx` (or other plain data type) as its first
  argument. Classes can't be tree-shaken once an instance escapes; we
  designed away from them deliberately.
- **One way to do one thing.** Don't add a parallel API that reaches an
  already-reachable capability. Rename instead of parallelising; graduate
  the helper instead of exposing two paths.
- **Validate at boundaries; trust internally.** Schema-validate external
  input and third-party calls; don't re-check values upstream code already
  proved correct. Let internal bugs throw — don't swallow exceptions.
- **Round-trip is sacred.** Anything the AST doesn't yet model must pass
  through as a raw node so re-saving an untouched template doesn't trigger
  Word's "needs repair" prompt. New OOXML support should land with at
  least one round-trip fixture.
- **Comments explain _why_, not _what_.** Identifier names + types
  document _what_; comments earn their keep only when they capture
  invariants, trade-offs, or non-obvious history.

`oxlint`, `oxfmt`, and `tsc` enforce the rest; please don't disable rules
locally without flagging it in the PR description.

## Tests

- **Unit and integration tests** live next to the code under `src/`
  (and `packages/preview/src/`). The file naming is `<thing>.test.ts`.
- **Round-trip tests** load a fixture, save it, and assert the output is
  semantically equal to the input. They're how we keep parity with Word /
  LibreOffice / mammoth.js / python-docx output. When you add support for
  a new OOXML element, add at least one round-trip fixture.
- **Performance smoke tests** (`src/api/perf-smoke.test.ts`)
  guard against accidental O(n²) regressions on common shapes (10k
  paragraphs, 100 tables, large find/replace). Budgets are loose so CI on
  slow runners doesn't false-fail.

The CI matrix runs against Node 22 and 24 on Ubuntu. Node 20 was dropped
after it reached EOL on 2026-04-30. The browser-preview tests run under
`happy-dom`; jsdom's cross-realm `Uint8Array` confused `fflate`'s type
guards (see commit `8039586` for the history).

## Packages

`@office-kit/docx` ships as a single self-contained package built from
layered modules. The public library lives at the repo root (`src/`); the
lower layers are inlined under `src/internal/` and bundled in — they are
**not** separate npm packages. The browser preview is the only other
published package, under `packages/preview`.

| Location                                        | What it owns                                                       |
| ----------------------------------------------- | ------------------------------------------------------------------ |
| `src/api` (`@office-kit/docx`)                  | The public `Docx` interface + the function API users import.       |
| `src/internal/wordprocessingml`                 | WordprocessingML AST: parsers, writers, builders.                  |
| `src/internal/xml`                              | Namespace-aware XML parser / serializer with round-trip fidelity.  |
| `src/internal/opc`                              | OPC packaging (ZIP + content types + rels) with byte-stable parts. |
| `packages/preview` (`@office-kit/docx-preview`) | Browser-side read-only preview (wraps `docx-preview`).             |

Most contributions touch `src/api` or `src/internal/wordprocessingml`; the
lower two layers (`xml`, `opc`) are stable and only get changes when the
schema below them moves.

The browser preview is **intentionally** a thin wrap over
[`docx-preview`](https://github.com/VolodymyrBaydalka/docxjs). Don't
propose re-implementing the renderer in `@office-kit/docx-preview` — see
[`docs/PLAN-PREVIEW.md`](docs/PLAN-PREVIEW.md) for the rationale.

## Releases

word-kit publishes via npm using Changesets:

1. PRs land on `main` with their associated changeset.
2. The Changesets action opens a "Version Packages" PR that bumps versions
   and consumes the pending changesets.
3. Merging that PR triggers the release workflow, which runs the full gate
   one more time and publishes to npm.

Maintainers: never publish from a local machine.

## Questions

If anything in this document is unclear, open an issue or start a
discussion. PRs welcome.
