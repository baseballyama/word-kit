<script lang="ts">
  import { base } from '$app/paths';
  import CodeBlock from '$lib/components/CodeBlock.svelte';
  import type { PageProps } from './$types';

  const { data }: PageProps = $props();
</script>

<svelte:head>
  <title>Getting started · word-kit</title>
</svelte:head>

<article class="prose">
  <p class="eyebrow">§ 01 · Getting started</p>
  <h1>Install, build a document, write the bytes.</h1>

  <p class="lede">
    word-kit is a function-first WordprocessingML library. Every operation is a standalone
    export that takes a <code>Docx</code> value as its first argument; nothing is hidden behind
    a class instance. That keeps the public surface tree-shakeable and makes the library work
    the same way in Node and in browsers.
  </p>

  <h2>Install</h2>

  <pre class="install"><span class="dollar">$</span> pnpm add @word-kit/core @word-kit/preview</pre>

  <p>
    <code>@word-kit/core</code> is the authoring API. <code>@word-kit/preview</code> is an
    optional companion that mounts a read-only preview of any <code>Docx</code> value into a DOM
    container. Both ship as ESM with bundled <code>.d.ts</code> types and have no Node-only
    dependencies.
  </p>

  <h2>Build a document from scratch</h2>

  <p>
    The "hello world" of word-kit. Every helper here lives on
    <code>@word-kit/core</code>; <code>createDocx</code> hands back a plain
    <code>Docx</code> object that the rest of the API treats as a value.
  </p>

  <CodeBlock
    html={data.fromScratch.html}
    source={data.fromScratch.source}
    title={data.fromScratch.path}
    coord="A1"
  />

  <h2>Open a template, fill placeholders</h2>

  <p>
    Existing <code>.docx</code> files can be opened with <code>openDocx</code> and edited
    in place. word-kit preserves every XML element it does not yet model as a pass-through
    node, so re-saving an unmodified template doesn't trip Word's "needs repair" prompt.
  </p>

  <p>
    <code>replaceTextEverywhere</code> walks every story — body, headers, footers, footnotes,
    endnotes, comments, textboxes — not just the main document. Run-spanning matches like
    <code>{'{{name}}'}</code> split across multiple runs are joined before the regex sees them.
  </p>

  <CodeBlock
    html={data.templateFill.html}
    source={data.templateFill.source}
    title={data.templateFill.path}
    coord="A2"
  />

  <h2>Render in the browser</h2>

  <p>
    The companion package <code>@word-kit/preview</code> mounts a read-only preview of any
    <code>Docx</code> (or raw bytes) into a DOM container. It wraps the OSS
    <code>docx-preview</code> renderer behind a stable function-API entry point.
  </p>

  <p>
    You can play with it on the <a href="{base}/playground">playground</a>, or read the
    <a href="{base}/docs/recipes">recipes</a> for embedding patterns.
  </p>

  <h2>What's next</h2>

  <ul>
    <li><a href="{base}/docs/recipes">Recipes</a> — common scenarios, copy-paste ready</li>
    <li><a href="{base}/api">API reference</a> — every public export, grouped by area</li>
    <li><a href="{base}/playground">Playground</a> — drop a .docx and see preview</li>
    <li><a href="https://github.com/baseballyama/word-kit">GitHub</a> — source + issues</li>
  </ul>
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
    font-variation-settings: 'opsz' 144, 'SOFT' 30;
    max-width: 22ch;
  }

  .lede {
    color: var(--fg-soft);
    font-size: 1.06rem;
    line-height: 1.55;
    max-width: 64ch;
    margin: 0 0 1.5rem;
  }

  .install {
    border: 1px solid var(--border);
    background: var(--code-bg);
    padding: 0.7rem 0.95rem;
    border-radius: var(--radius-sm);
    font-family: var(--mono);
    font-size: 0.9rem;
    color: var(--fg);
    margin: 1rem 0 1.5rem;
    overflow-x: auto;
  }

  .install .dollar {
    color: var(--accent);
    font-weight: 600;
    margin-right: 0.55rem;
  }
</style>
