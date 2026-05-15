# word-kit `.docx` browser preview — design plan

> **Status (2026-05-15 — final): SHIPPED at v0. v1 and v2 are not
> planned.**
>
> The `@word-kit/preview` package now wraps `docx-preview` behind the
> stable `previewToDOM(...)` function-API entry. That meets the user
> requirement ("プレビューできれば OK") so the more ambitious v1/v2 phases
> originally described below are explicitly dropped: implementing our
> own renderer would re-do years of `docx-preview` work for purely
> architectural-purity reasons that do not change what the user sees.
>
> The v1/v2 sections are kept below as a **historical record** of
> options that were considered and rejected, so a future contributor
> doesn't re-derive the same plan and walk down it. Re-open the
> question only if a concrete `docx-preview` bug blocks a real user
> workflow we cannot work around — at which point the right move is
> usually upstreaming a fix to `docx-preview`, not forking the
> renderer.

## 1. Why

`word-kit` today is a write-and-edit library: `createDocx`, `openDocx`,
`appendParagraph`, `addTable`, `toUint8Array`. The next missing piece — the
one we keep hearing about — is **show this docx in the browser**:

> "I produced a docx with `word-kit`. I want my web app to render a
> faithful preview of it next to the form/editor that produced it,
> without spinning up a server."

That is a real workflow (template-based document generation, mail-merge
preview, document review UIs, in-app proofreading) and `word-kit` is the
natural place to solve it because **it already owns the WML AST** the
preview would need to render.

## 2. Scope

### In scope (v1)

- Read-only preview of `.docx` content in any modern browser.
- Renders body text + paragraph styling, lists, tables, images,
  hyperlinks, headers/footers, footnotes/endnotes (as a flat tail
  list), bookmarks, comments (as side-bar markers), tracked changes
  visualised in `accept-all` form.
- Page-sized blocks broken at _source-declared_ breaks
  (`<w:br w:type="page"/>`, `<w:lastRenderedPageBreak/>`) — the same
  bar `docx-preview` clears today.
- Custom paragraph / character / table styles defined in the docx
  resolve correctly (since we already parse `styles.xml`).
- Built-in CSS classes are namespaced (configurable prefix) so
  preview output coexists with the host page's CSS.
- ESM-only, no Node-only deps in the published bundle.
- Tree-shake friendly — sits behind its own subpath import
  (`@word-kit/preview`) so consumers who don't need preview pay zero
  bundle cost.

### Out of scope (v1)

- **Live re-pagination** as content changes. Pages are drawn at
  source breaks. Same compromise `docx-preview` ships with — no
  serious in-browser docx renderer does this today.
- **Field evaluation.** `TOC`, `PAGE`, `NUMPAGES`, etc. show their
  cached display value (`<w:fldChar w:fldCharType="separate"/>` →
  `<w:t>`) when present, otherwise the field instruction. Filing
  this as a v2 follow-up.
- **Editing.** Read-only. Editing is a much larger commitment
  (selection model, IME, undo/redo, collab) that competes head-on
  with SuperDoc / `eigenpal/docx-js-editor`. Out of scope until v3
  at the earliest.
- **Pixel-perfect Word fidelity.** HTML/CSS cannot replicate every
  WML page semantic; that's a known industry limit. Goal is "looks
  like a Word doc," not "byte-identical to Word's renderer."
- **Print-perfect output.** If you need that, convert to PDF via
  LibreOffice headless and render with `pdf.js`. We document the
  recipe; we do not ship it.
- **DrawingML beyond inline pictures.** Floating shapes, SmartArt,
  charts, equations remain pass-through for now.

### Permanent non-goals

- Server-rendered tile streaming (OnlyOffice / Collabora model).
- A WebAssembly LibreOffice port. Tens of MB; wrong fit for a
  library.
- An `altChunk` HTML-fragment renderer. Word-only feature, edge
  case.

## 3. Survey conclusions (informing the plan)

Five OSS approaches exist; the trade-offs are summarised below — the
full survey lives in the planning notes from this commit.

| Approach                           | Fidelity                                                                                | Cost to adopt            | License                 | Browser-only     | Verdict                                                                                             |
| ---------------------------------- | --------------------------------------------------------------------------------------- | ------------------------ | ----------------------- | ---------------- | --------------------------------------------------------------------------------------------------- |
| **`docx-preview`** wrap            | Medium (good enough for ~80% of docx)                                                   | Days                     | Apache-2.0              | ✓                | Ship-fast bridge. Carries a second OOXML parser in the bundle (we already have ours).               |
| **Build from `word-kit` WML AST**  | Match `docx-preview` at v1, beat it on tab stops / lists / tables we know it gets wrong | Months                   | own (MIT)               | ✓                | On-brand: makes the renderer the _reverse_ of the writer; round-trip tests now bidirectional.       |
| **SuperDoc** wrap                  | High (paginated, real layout engine)                                                    | Days                     | **AGPLv3** / commercial | ✓                | Hard pass: AGPL propagates to anyone serving an app over the network.                               |
| **`eigenpal/docx-js-editor`** wrap | Medium-high but the authors flag tables/images as incomplete                            | Days                     | MIT                     | ✓                | Too young (~656 stars, launched 2025), small team, fidelity gaps acknowledged. Watch, don't depend. |
| **LibreOffice → PDF → `pdf.js`**   | Print-perfect                                                                           | Weeks for the recipe doc | MPL/Apache              | ✗ (needs server) | Out of scope as a `word-kit` deliverable; documented separately.                                    |

Headline observations:

- **Wrapping anyone forfeits long-tail bugs.** Every wrapped
  renderer has a known set of "we don't fix that" gaps (no live
  pagination, no TOC, list/tab edge cases, table backgrounds). We
  inherit those forever.
- **Owning the renderer is on-brand.** `word-kit`'s `CLAUDE.md` is
  explicit about avoiding parallel APIs and second-implementation
  parsers. A second OOXML parser in the bundle is precisely the
  shape of dependency the project is set up to refuse.
- **A bridge buys time.** We can ship preview in days by wrapping
  `docx-preview` _behind a stable function-API entry point_, then
  swap the implementation later without breaking consumers.

## 4. Final decision (2026-05-15): stop at v0

The wrap-`docx-preview` v0 ships the user-visible feature
("プレビューできれば OK"). The originally-planned v1 (own renderer)
and v2 (layout engine) phases are **dropped** — they would re-do work
`docx-preview` already does for purely architectural-purity reasons
that do not change what the user sees.

If a concrete `docx-preview` bug blocks a real user workflow that
can't be worked around, the right next move is **upstream a fix**,
not re-implement. Re-opening v1/v2 requires that level of evidence.

The v0 phase below is the actual shipped state. The v1/v2 sections
are kept as a **historical record of considered-and-rejected
options** so a future contributor doesn't re-derive the same plan
and walk down it.

## 4a. Original recommendation (HISTORICAL — superseded by §4)

A three-phase rollout, each phase shipping a usable feature, with the
public API stable across all three.

### Public API surface (stable across phases)

A single new package: **`@word-kit/preview`**.

A single function entry — consistent with `word-kit`'s no-classes /
function-API rule:

```ts
import { previewToDOM } from "@word-kit/preview";
import { openDocx } from "@word-kit/core";

const doc = openDocx(bytes);
const handle = await previewToDOM(doc, document.getElementById("preview")!, {
  classPrefix: "wk-", // CSS namespace
  inWrapper: true, // emit a scrolling page wrapper
  ignoreFonts: false, // skip @font-face injection
  breakPages: true, // honour source page breaks
  experimentalComments: false,
  experimentalTrackedChanges: false,
});

handle.dispose(); // detaches CSS, drops DOM
```

`previewToDOM` accepts:

- the `Docx` returned by `openDocx` / `createDocx`, **or**
- raw bytes (`Uint8Array | Blob`) — internally calls `openDocx` for the
  caller, since most preview callers just hand over what they have.

It returns a small handle object (still a plain interface, no class),
similar in spirit to the `Docx` interface itself: `{ dispose(): void }`.

### Phase 0 — bridge (week 1)

- New package `packages/preview` published as `@word-kit/preview`.
- Implementation: thin adapter over `docx-preview`. Convert a
  `Docx` value to bytes via `toUint8Array(doc)`, then hand to
  `docx-preview.renderAsync`. (Yes, this re-parses the OOXML —
  that's what "bridge" means.)
- Same single function `previewToDOM` is the entry point.
- Tree-shake check: importing only `previewToDOM` should NOT pull
  the rest of `@word-kit/core` into the bundle for a consumer who
  also doesn't import `@word-kit/core` themselves. (Achievable
  because `@word-kit/preview` only imports `toUint8Array` and
  `openDocx` from core — both already proven to tree-shake.)
- Documentation: clearly mark v0 as a stop-gap, no semver-1.0
  commitment to behaviour, swap planned for v1.

**Ship goal: a usable preview within one working day of starting.**

### Phase 1 — own the renderer (weeks 2-8)

Replace the v0 implementation with a renderer driven by `word-kit`'s
own WML AST. The function signature stays identical so the swap is
invisible to consumers.

Architecture (mirrors `docx-preview`'s, but native to our types):

```
@word-kit/preview
├── src/
│   ├── index.ts              ← previewToDOM (the only public export)
│   ├── render-document.ts    ← walks WmlDocument.body
│   ├── render-paragraph.ts   ← <w:p> → <p class="wk-p"> (plus pStyle, alignment, indent, spacing)
│   ├── render-run.ts         ← <w:r> → <span class="wk-r"> (plus rPr → inline style or class)
│   ├── render-table.ts       ← <w:tbl> → <table>
│   ├── render-list.ts        ← numbering.xml → CSS counter chains
│   ├── render-image.ts       ← <w:drawing> → <img src="data:..."> (re-uses media parts already in opc)
│   ├── render-hyperlink.ts   ← <w:hyperlink> → <a>
│   ├── render-header-footer.ts ← top/bottom-of-page slots
│   ├── render-footnote.ts    ← tail block at end of preview
│   ├── render-comment.ts     ← side-bar markers
│   ├── style-resolver.ts     ← styles.xml + theme.xml → CSS class declarations
│   ├── font-loader.ts        ← @font-face injection
│   └── pages.ts              ← split body into page-sized wrappers at source breaks
└── tests/
    ├── render-fixtures.test.ts ← snapshot the produced DOM HTML for every sample under ./samples/
    └── round-trip.test.ts      ← preview(openDocx(toUint8Array(createDocx(...)))) == preview(createDocx(...))
```

#### Implementation order (pull bugs out one feature area at a time)

1. **Bare body**: paragraphs + runs + character formatting (bold,
   italic, underline, color, size, font). Visual diff against
   `docx-preview` rendering of the same fixture.
2. **Lists**: bullet + numbered, multi-level, using CSS counters.
   Close the `docx-preview` issue #181 gap as a stretch goal.
3. **Tables**: `<w:tblBorders>`, `<w:shd>`, vertical alignment, row
   height — all the helpers we already write through `setTableX`
   should _render_ correctly here.
4. **Images**: pull bytes from the OPC media part, base64-encode
   into `<img src="data:...">`. (Could use `URL.createObjectURL`
   later for memory efficiency; v1 keeps it simple.)
5. **Hyperlinks + bookmarks**.
6. **Headers / footers**: render at top/bottom of each page-sized
   wrapper.
7. **Footnotes / endnotes**: tail block, anchored from inline
   reference markers.
8. **Comments**: side-bar with anchored highlighters.
9. **Tracked changes**: render in `accept-all` form (insertions
   shown, deletions hidden) for v1; full review-pane visualisation
   in v2.
10. **Sections**: page size / orientation / margins per section.

#### Test strategy

Existing `samples/` script already produces ~32 docx files covering
every feature combination. The preview tests will:

- Snapshot the produced DOM HTML for each `samples/*.docx` and check
  it into the repo.
- Diff future rendering changes against those snapshots in CI.
- For the styled-base templates specifically: confirm that opening
  `30-styled-base-template.docx` and rendering it shows the
  template's `AcmeTitle` / `AcmeBody` styles applied to the title
  page text (proves style resolution works from `styles.xml`, not
  from inline rPr).

A **visual snapshot** (Playwright + visual diff against `docx-preview`'s
rendering of the same fixture) is a v1 stretch goal. If we run it,
diffs become "we render this differently — better, worse, or
neutral?" decisions, not pass/fail.

### Phase 2 — layout engine (only if demand justifies)

`docx-preview` and we (in v1) both punt on live pagination because it
requires a real layout engine: measure each block, accumulate height,
break to a new page when the page-height budget is exceeded. SuperDoc
has built one; it's a substantial multi-month investment.

Trigger to start v2:

- Three or more reported bugs of the form "I edited the doc in the
  preview's parent app and the page breaks didn't move" (or similar
  re-flow expectations).
- A user willing to dogfood the alpha.

Architecture would adopt SuperDoc's pattern (`FlowBlocks → Layout →
DomPainter`) without taking the AGPL code; the _idea_ is unencumbered.
Risk: this is where projects stall. Don't start until necessary.

### Permanent non-deliverable: PDF rendering

Documented as a recipe under `docs/recipes/pdf-preview.md`:

> If you need print-perfect rendering, run LibreOffice headless on
> your server (`libreoffice --headless --convert-to pdf
input.docx`), serve the PDF, and render with `pdf.js` in the
> browser. Memory and cold-start tradeoffs are described at
> [LibreOffice serverless guide](https://levelup.gitconnected.com/libreoffice-docker-express-lambda-convert-office-to-pdf-serverless-for-free-8781bc2f0c55).

We never ship the conversion ourselves.

## 5. Detailed task list

### Phase 0 — bridge

| #   | Task                                                                                                                                            | Estimate | Notes                                                       |
| --- | ----------------------------------------------------------------------------------------------------------------------------------------------- | -------- | ----------------------------------------------------------- | ---------- | --------------------------------------- |
| 0.1 | Add `packages/preview` workspace package, `@word-kit/preview`, ESM-only, function-API.                                                          | 1h       | Mirrors the layout of the other 4 packages.                 |
| 0.2 | Add `docx-preview` as a runtime dependency.                                                                                                     | 5min     | Apache-2.0; bundle-size impact disclosed in CHANGELOG.      |
| 0.3 | Implement `previewToDOM(doc, container, options?)` as a thin wrapper around `docx-preview.renderAsync`.                                         | 4h       | Accept `Docx                                                | Uint8Array | Blob`; convert `Docx`via`toUint8Array`. |
| 0.4 | Tree-shake budget: extend `pnpm check:tree-shake` so `import { previewToDOM } from "@word-kit/preview"` is bounded (e.g. ≤1.2MB minified gzip). | 2h       | Two budgets: minimal core, minimal core+preview.            |
| 0.5 | Add `samples` walk that renders each `samples/*.docx` to a static HTML page under `samples/preview/` for human review.                          | 2h       | Uses jsdom in Node so the sample script is still pure node. |
| 0.6 | Vitest integration test: `previewToDOM(createDocx({paragraphs:["hello"]}), container)` produces a DOM whose textContent contains "hello".       | 1h       | Uses jsdom test environment.                                |
| 0.7 | Doc: `docs/preview.md` with the API + a "this is v0, the implementation will be replaced" footnote.                                             | 1h       |                                                             |
| 0.8 | Changeset entry: feat(preview).                                                                                                                 | 5min     |                                                             |

### Phase 1 — own the renderer

| #   | Task                                                                                                                              | Estimate                             |
| --- | --------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------ |
| 1.1 | Replace internal implementation of `previewToDOM` with native renderer. Keep public signature unchanged.                          | 1d (skeleton)                        |
| 1.2 | Per-feature renderers (paragraph, run, list, table, image, hyperlink, header/footer, footnote, comment, tracked change, section). | 1-2 weeks each block, parallelisable |
| 1.3 | `style-resolver.ts`: emit `<style>` block driven by `styles.xml`, with class-prefix namespacing.                                  | 3d                                   |
| 1.4 | `font-loader.ts`: extract embedded fonts from `word/fonts/`, inject `@font-face`.                                                 | 2d                                   |
| 1.5 | `pages.ts`: split body into page-sized wrappers at source breaks.                                                                 | 2d                                   |
| 1.6 | DOM snapshot tests for every fixture under `samples/`.                                                                            | 2d                                   |
| 1.7 | Drop `docx-preview` runtime dep. Update tree-shake budget downwards.                                                              | 1d                                   |
| 1.8 | Side-by-side visual diff (Playwright) against `docx-preview` for each fixture; treat differences as findings, not failures.       | 3d                                   |

### Phase 2 — layout engine (deferred)

Only sized when triggered. Rough order: `FlowBlock` model → block
height measurement (in headless DOM via `getBoundingClientRect`) →
page-budget allocator → `DomPainter` → re-flow on content change.

## 6. Open questions (please decide before phase 0 starts)

1. **Subpath vs separate package?** Plan above is a separate package
   (`@word-kit/preview`). Subpath import (`@word-kit/core/preview`)
   is also viable but couples preview lifecycle to core. Separate
   package matches the existing OPC / WML / XML separation.

2. **Class prefix default.** `wk-`? `word-kit-`? `wkpreview-`?
   Defaulting to `wk-` is shortest; document that consumers can
   override if it collides.

3. **`<style>` block placement.** Inject into `<head>` (cleanest,
   risks duplicate emission across multiple `previewToDOM` calls)
   or scope it to a `<style>` element inside the container (slightly
   less clean DOM but multi-instance-safe). Recommend the latter.

4. **Image handling for large media.** Base64 inline (simple, no
   blob URL lifecycle to manage) or `URL.createObjectURL` (memory-
   efficient, requires `dispose()` to revoke). Recommend base64 in
   v0; revisit for v1 if memory profiling shows a problem.

5. **Synchronous or async?** `previewToDOM` returns a Promise even
   if the v0 wrapper is async-by-coincidence (`docx-preview` is
   async). v1 native renderer doesn't strictly need to be async —
   but staying async future-proofs us for font loading, image
   decoding, and worker-based layout.

6. **Worker offload (v1)?** Renderer could parse + emit DOM strings
   in a worker, then `container.innerHTML = ...` on the main
   thread. Adds complexity; do not do in v0; consider for v1 if
   render time on large docs blocks the main thread.

7. **Does the v0 wrapper acknowledge it re-parses?** Be transparent
   in the docstring: "v0 sends bytes through `docx-preview`. v1 will
   render directly from the AST without a second parse."

## 7. Risks & mitigations

| Risk                                                                                         | Likelihood                                               | Impact                   | Mitigation                                                                                                                                                                              |
| -------------------------------------------------------------------------------------------- | -------------------------------------------------------- | ------------------------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| v1 native renderer takes longer than estimated.                                              | High (renderers always do)                               | Schedule slip            | Phase 0 ships first. v1 can be partial — replace areas that have parity-with-`docx-preview`, leave the rest delegated.                                                                  |
| Bundle size balloons because we ship two parsers in v0.                                      | Certain in v0                                            | Tree-shake budget breach | Sit `@word-kit/preview` behind its own package so only consumers who import it pay the cost. CI budget checks both.                                                                     |
| `docx-preview` upstream breaks our wrapper.                                                  | Low                                                      | Schedule slip            | Pin major version. Wrapper is shallow; backports are cheap.                                                                                                                             |
| v1 visual diff against `docx-preview` shows we render _worse_ on some fixtures.              | Likely on long-tail features (tab stops, complex tables) | Perception               | Document each known difference explicitly in `docs/preview.md`. We're allowed to be different where `docx-preview` is wrong; we are not allowed to be silently different.               |
| HTML/CSS cannot reproduce some Word page semantics (well-known limitation, see TextControl). | Certain                                                  | Some fixtures look "off" | Out of scope for v1. Goal is "looks like a Word doc", not "byte-identical to Word's renderer".                                                                                          |
| Live pagination requested.                                                                   | Medium                                                   | Big v2 commitment        | Document it as out of scope until enough demand accumulates.                                                                                                                            |
| Font availability. Documents reference fonts the user's machine doesn't have.                | Certain                                                  | Visual drift             | v0 inherits `docx-preview`'s behavior. v1: extract embedded fonts (`word/fonts/`) and inject `@font-face`; warn in dev when a referenced font is neither embedded nor system-installed. |

## 8. Success metrics

A "we shipped v1" acceptance set:

- Every `samples/*.docx` renders without throwing, and at least its
  body text appears in the DOM.
- The styled-base template (`30-styled-base-template.docx`) renders
  the title in `AcmeTitle`'s bold dark-blue 22pt look, proving
  style resolution from `styles.xml` works.
- Tree-shake budget for a consumer that imports only
  `previewToDOM` (no other `@word-kit/core` symbols) stays under
  600 KB minified, ungzipped (estimate; tighten after first build).
- Suite passes 100% in CI across Node 20 / 22 / 24, and the new
  jsdom-based render tests pass.
- Manual audit: open every `samples/*.docx` in Microsoft Word AND
  in our preview side-by-side; record per-fixture differences in
  `docs/preview-fidelity.md`.

## 9. Decision log (to be updated as decisions firm up)

| Date       | Decision                                           | Rationale                                                                                                                |
| ---------- | -------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------ |
| 2026-05-15 | Recommend "v0 bridge → v1 own → v2 layout engine". | Ships fast, then converges with `word-kit`'s "own the OOXML stack" posture without locking in a second parser long-term. |
|            | Subpath import vs separate package.                | Pending — see open question 1.                                                                                           |
|            | Default class prefix.                              | Pending — see open question 2.                                                                                           |
|            | Base64 vs blob URL for images.                     | Pending — see open question 4.                                                                                           |

## 10. References

- [VolodymyrBaydalka/docx-preview](https://github.com/VolodymyrBaydalka/docxjs)
- [mwilliamson/mammoth.js](https://github.com/mwilliamson/mammoth.js)
- [Harbour-Enterprises/SuperDoc](https://github.com/Harbour-Enterprises/SuperDoc) (architecture reference)
- [SuperDoc CONTRIBUTING — pipeline architecture](https://github.com/Harbour-Enterprises/SuperDoc/blob/main/CONTRIBUTING.md)
- [eigenpal/docx-js-editor](https://github.com/eigenpal/docx-js-editor) (MIT alternative to SuperDoc)
- [ONLYOFFICE/DocumentServer](https://github.com/ONLYOFFICE/DocumentServer) (server-side reference)
- [Collabora Online](https://github.com/CollaboraOnline/online) (LibreOffice-based reference)
- [LibreOffice-headless serverless guide](https://levelup.gitconnected.com/libreoffice-docker-express-lambda-convert-office-to-pdf-serverless-for-free-8781bc2f0c55)
- [Mozilla pdf.js](https://github.com/mozilla/pdf.js)
- [TextControl — Why HTML is not a substitute for DOCX](https://www.textcontrol.com/blog/2025/08/19/why-html-is-not-a-substitute-for-page-oriented-formats-like-docx/)
