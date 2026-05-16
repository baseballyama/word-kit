<script lang="ts">
  import { base } from '$app/paths';
  import { apiGroups, apiTotalCount } from '$lib/api-groups';

  // Data lives in `$lib/api-groups` so `/llms-full.txt` and the
  // `check:api-page` CI gate consume the same source.
  const groups = apiGroups;
  const totalCount = apiTotalCount;
</script>

<svelte:head>
  <title>API · word-kit</title>
</svelte:head>

<section class="head">
  <div class="head-inner">
    <p class="eyebrow">§ 03 · API reference</p>
    <h1>Every public export, grouped by area.</h1>
    <p class="lede">
      <code>@word-kit/core</code> exposes <strong>{totalCount}</strong> standalone functions
      and constants. Each one takes a <code>Docx</code> as its first argument (where
      applicable) and is side-effect-free at the module level, so bundlers can tree-shake
      anything you don't import. For full signatures and parameter shapes, see
      <code>packages/core/src/index.ts</code> and the
      <a href="https://github.com/baseballyama/word-kit">repository</a>.
    </p>
  </div>
</section>

<section class="api">
  <div class="api-inner">
    {#each groups as g (g.num)}
      <article class="group">
        <header class="group-head">
          <span class="g-num">{g.num}</span>
          <h2>{g.title}</h2>
          <span class="g-count">{g.entries.length.toString().padStart(2, '0')}</span>
        </header>
        <ul class="entries">
          {#each g.entries as e (e.name)}
            <li class="entry">
              <code class="name">{e.name}</code>
              {#if e.sig}<span class="sig">{e.sig}</span>{/if}
            </li>
          {/each}
        </ul>
      </article>
    {/each}

    <p class="more">
      Need usage shapes? Read the <a href="{base}/docs/getting-started">getting-started page</a>
      or the type-checked <a href="{base}/docs/recipes">recipes</a>.
    </p>
  </div>
</section>

<style>
  .head {
    padding: 3rem 1.5rem 1.75rem;
    border-bottom: 1px solid var(--border);
  }

  .head-inner {
    max-width: var(--max-wide);
    margin: 0 auto;
  }

  h1 {
    font-family: var(--display);
    font-weight: 460;
    font-size: clamp(1.9rem, 4.5vw, 2.85rem);
    line-height: 1.05;
    margin: 0 0 0.8rem;
    font-variation-settings: 'opsz' 144, 'SOFT' 30;
    letter-spacing: -0.025em;
    max-width: 22ch;
  }

  .lede {
    color: var(--fg-soft);
    font-size: 1.04rem;
    line-height: 1.55;
    max-width: 70ch;
    margin: 0;
  }

  .api {
    padding: 2.5rem 1.5rem 5rem;
  }

  .api-inner {
    max-width: var(--max-wide);
    margin: 0 auto;
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(320px, 1fr));
    gap: 1rem;
  }

  .group {
    border: 1px solid var(--border);
    border-radius: var(--radius);
    background: var(--bg-elev);
    overflow: hidden;
  }

  .group-head {
    display: flex;
    align-items: baseline;
    gap: 0.7rem;
    padding: 0.85rem 1rem 0.7rem;
    background: var(--bg-paper);
    border-bottom: 1px solid var(--border);
  }

  .g-num {
    font-family: var(--mono);
    font-size: 11px;
    color: var(--accent);
    letter-spacing: 0.08em;
  }

  .group-head h2 {
    margin: 0;
    border: none;
    padding: 0;
    font-family: var(--display);
    font-size: 1.05rem;
    font-weight: 540;
    flex: 1;
    font-variation-settings: 'opsz' 32, 'SOFT' 25;
  }

  .g-count {
    font-family: var(--mono);
    font-size: 11px;
    color: var(--fg-muted);
    letter-spacing: 0.1em;
  }

  .entries {
    list-style: none;
    padding: 0;
    margin: 0;
  }

  .entry {
    display: flex;
    align-items: baseline;
    gap: 0.7rem;
    padding: 0.45rem 1rem;
    border-bottom: 1px dashed var(--rule);
    font-size: 0.9rem;
  }

  .entry:last-child {
    border-bottom: none;
  }

  .name {
    font-family: var(--mono);
    font-size: 12.5px;
    background: transparent;
    border: none;
    padding: 0;
    color: var(--fg);
    flex: none;
  }

  .sig {
    font-family: var(--mono);
    font-size: 11.5px;
    color: var(--fg-muted);
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }

  .more {
    grid-column: 1 / -1;
    margin-top: 1.25rem;
    color: var(--fg-soft);
    border-top: 1px solid var(--border);
    padding-top: 1.25rem;
    font-size: 0.95rem;
  }
</style>
