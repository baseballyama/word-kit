<script lang="ts">
  import { base } from "$app/paths";
  import CodeBlock from "$lib/components/CodeBlock.svelte";
  import type { PageProps } from "./$types";

  const { data }: PageProps = $props();

  // Cross-reference pointers: which sample file (from `pnpm sample`) or
  // integration test demonstrates each recipe end-to-end. Kept in a
  // separate map from the type-checked examples so the snippet stays
  // minimal and the cross-reference is a hairline pointer underneath.
  const pointers: Record<string, string> = {
    recipeMailMerge: "samples/20-mailmerge-text-replace-*.docx (run `pnpm sample`)",
    recipeStyledBase: "samples/30-styled-base-*.docx",
    recipeTrackedChanges: "samples/09-tracked-changes.docx",
    previewEmbed: "see live in the /playground page",
  };
</script>

<svelte:head>
  <title>Recipes · word-kit</title>
</svelte:head>

<article class="prose">
  <p class="eyebrow">§ 02 · Recipes</p>
  <h1>Common scenarios, type-checked snippets.</h1>

  <p class="lede">
    Every snippet below lives under <code>site/src/lib/examples/</code> and is
    type-checked by <code>svelte-check</code> against the live
    <code>@word-kit/core</code> / <code>@word-kit/preview</code> surface. An API
    rename breaks this page before anything ships. Pointers under each snippet
    name the matching sample file produced by <code>pnpm sample</code> (or the
    integration test that exercises the same path).
  </p>

  {#each data.recipes as r, i (r.key)}
    <section class="recipe">
      <header class="r-head">
        <span class="r-num">{String(i + 1).padStart(2, "0")}</span>
        <div class="r-text">
          <h2>{r.title}</h2>
          <p>{r.description}</p>
          <p class="r-pointer">
            <span class="label">where:</span>
            <code>{pointers[r.key] ?? "—"}</code>
          </p>
        </div>
      </header>
      <CodeBlock
        html={r.html}
        source={r.source}
        title={r.path}
        coord={String.fromCharCode(64 + i + 1) + "1"}
      />
    </section>
  {/each}

  <p class="more">
    Want to drive a doc end-to-end? Walk the
    <a href="{base}/docs/getting-started">Getting started</a> page, then poke at
    <a href="{base}/playground">the playground</a> to see the rendered output.
  </p>
</article>

<style>
  .prose {
    max-width: var(--max-content);
    margin: 0 auto;
    padding: 3rem 1.5rem 5rem;
  }

  h1 {
    font-family: var(--display);
    font-weight: 460;
    font-size: clamp(2rem, 4.6vw, 2.95rem);
    line-height: 1.05;
    letter-spacing: -0.026em;
    margin: 0 0 1rem;
    font-variation-settings: "opsz" 144, "SOFT" 30;
    max-width: 22ch;
  }

  .lede {
    color: var(--fg-soft);
    font-size: 1.06rem;
    line-height: 1.55;
    max-width: 64ch;
    margin: 0 0 2rem;
  }

  .recipe {
    margin: 2.5rem 0 3rem;
    padding-top: 2rem;
    border-top: 1px solid var(--border);
  }

  .r-head {
    display: grid;
    grid-template-columns: 4ch 1fr;
    gap: 1.25rem;
    margin: 0 0 0.5rem;
  }

  .r-num {
    font-family: var(--mono);
    font-size: 11.5px;
    color: var(--accent);
    font-weight: 500;
    letter-spacing: 0.06em;
    margin-top: 0.55rem;
  }

  .r-text h2 {
    margin: 0 0 0.5rem;
    border: none;
    padding: 0;
    font-family: var(--display);
    font-size: 1.45rem;
    font-weight: 540;
    font-variation-settings: "opsz" 64, "SOFT" 25;
  }

  .r-text p {
    margin: 0;
    color: var(--fg-soft);
    font-size: 1rem;
    line-height: 1.55;
  }

  .r-pointer {
    margin-top: 0.55rem !important;
    font-family: var(--mono);
    font-size: 11.5px;
    color: var(--fg-muted);
  }

  .r-pointer .label {
    text-transform: uppercase;
    letter-spacing: 0.12em;
    margin-right: 0.4rem;
    color: var(--fg-faint);
  }

  .more {
    margin-top: 2.5rem;
    border-top: 1px solid var(--border);
    padding-top: 1.25rem;
    color: var(--fg-soft);
    font-size: 0.95rem;
  }
</style>
