<script lang="ts">
  type Props = {
    /** Pre-rendered Shiki HTML (the full <pre>...</pre>). */
    html: string;
    /** Raw source text — what the copy button writes to the clipboard. */
    source: string;
    /** Optional file path / caption shown above the snippet. */
    title?: string;
    /** Optional coordinate-style label (defaults to next sequential index). */
    coord?: string;
    /** Language tag shown on the right of the caption. */
    lang?: string;
  };

  const { html, source, title, coord, lang = 'ts' }: Props = $props();

  let copied = $state(false);
  let resetTimer: ReturnType<typeof setTimeout> | null = null;

  async function copyToClipboard(): Promise<void> {
    try {
      await navigator.clipboard.writeText(source);
      copied = true;
      if (resetTimer) clearTimeout(resetTimer);
      resetTimer = setTimeout(() => {
        copied = false;
      }, 1500);
    } catch {
      // Permissions denied / insecure context — silently ignore. The
      // text is still selectable, so the user can fall back to manual
      // copy.
    }
  }
</script>

<figure class="code-block">
  <figcaption>
    {#if coord}
      <span class="coord">{coord}</span>
    {/if}
    {#if title}
      <span class="path">{title}</span>
    {:else}
      <span class="path"></span>
    {/if}
    <span class="lang">.{lang}</span>
    <button
      type="button"
      class="copy"
      class:done={copied}
      onclick={copyToClipboard}
      aria-label={copied ? 'Copied to clipboard' : 'Copy code to clipboard'}
    >
      {#if copied}
        <svg viewBox="0 0 16 16" width="13" height="13" aria-hidden="true">
          <path d="M3.5 8.5l3 3 6-6" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>
        <span>Copied</span>
      {:else}
        <svg viewBox="0 0 16 16" width="13" height="13" aria-hidden="true">
          <rect x="4" y="4" width="9" height="10" rx="1.5" fill="none" stroke="currentColor" stroke-width="1.4"/>
          <path d="M2.5 10.5V3a.5.5 0 0 1 .5-.5h7" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/>
        </svg>
        <span>Copy</span>
      {/if}
    </button>
  </figcaption>
  <div class="body">{@html html}</div>
</figure>

<style>
  .code-block {
    margin: 1.4rem 0;
    border: 1px solid var(--border);
    border-radius: var(--radius);
    background: var(--code-bg);
    overflow: hidden;
    box-shadow: 0 1px 0 0 var(--border-strong) inset;
  }

  figcaption {
    display: flex;
    align-items: center;
    gap: 0.6rem;
    padding: 0.45rem 0.5rem 0.45rem 0.9rem;
    background: var(--bg-paper);
    border-bottom: 1px solid var(--border);
    font-family: var(--mono);
    font-size: 12px;
    color: var(--fg-soft);
    letter-spacing: 0.02em;
  }

  .coord {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    min-width: 2.6ch;
    padding: 0.1em 0.4em;
    font-size: 11px;
    font-weight: 500;
    color: var(--accent);
    background: var(--accent-soft);
    border: 1px solid var(--accent-soft);
    border-radius: 3px;
    letter-spacing: 0.04em;
  }

  .path {
    color: var(--fg);
    flex: 1;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }

  .lang {
    color: var(--fg-muted);
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    font-weight: 500;
  }

  .copy {
    display: inline-flex;
    align-items: center;
    gap: 0.35rem;
    background: var(--bg-soft);
    border: 1px solid var(--border);
    border-radius: var(--radius-sm);
    padding: 0.25rem 0.55rem;
    color: var(--fg-soft);
    font-family: var(--mono);
    font-size: 11px;
    letter-spacing: 0.04em;
    cursor: pointer;
    transition:
      color 120ms ease,
      background 120ms ease,
      border-color 120ms ease;
  }

  .copy:hover {
    color: var(--fg);
    background: var(--bg);
    border-color: var(--border-strong);
  }

  .copy:focus-visible {
    outline: 1px solid var(--accent);
    outline-offset: 1px;
  }

  .copy.done {
    color: var(--accent);
    border-color: var(--accent-soft);
    background: var(--accent-soft);
  }

  .body :global(pre) {
    margin: 0;
    border: none;
    border-radius: 0;
    background: transparent !important;
    padding: 1rem 1.1rem;
  }

  .body :global(pre code) {
    font-family: var(--mono);
  }
</style>
