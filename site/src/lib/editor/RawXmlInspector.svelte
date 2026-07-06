<script lang="ts">
  import {
    rawTrees,
    partRawTree,
    xmlParts,
    runCommand,
    commands,
    type EditorModel,
    type RawNode,
  } from '@office-kit/docx-editor';
  import { t } from './i18n.svelte';

  let { model, onchange }: { model: EditorModel; onchange?: () => void } = $props();

  // `''` = the document inline view (DrawingML / OMML / VML in the body);
  // any other value is an XML part name edited through the `rawpart.*` commands.
  let source = $state('');
  let refresh = $state(0);

  const parts = $derived.by<string[]>(() => (refresh >= 0 ? xmlParts(model.doc) : []));
  const trees = $derived.by<RawNode[]>(() => {
    // Read `refresh` so the tree re-collects after each raw edit.
    if (refresh < 0) return [];
    if (source === '') return rawTrees(model.doc);
    const tree = partRawTree(model.doc, source);
    return tree ? [tree] : [];
  });

  function setAttr(node: RawNode, local: string, value: string): void {
    if (source === '')
      runCommand(model, commands.setAttributeCommand, { target: node.element, local, value });
    else
      runCommand(model, commands.setPartAttributeCommand, {
        partName: source,
        target: node.element,
        local,
        value,
      });
    refresh++;
    onchange?.();
  }

  function addChild(node: RawNode): void {
    const local = prompt('New child element local name (e.g. rot, uiPriority, sig)');
    if (!local) return;
    if (source === '')
      runCommand(model, commands.addChildCommand, { target: node.element, local });
    else
      runCommand(model, commands.addPartChildCommand, {
        partName: source,
        target: node.element,
        local,
      });
    refresh++;
    onchange?.();
  }
</script>

<div class="picker">
  <select bind:value={source} aria-label="XML source">
    <option value="">{t('xml.documentSource')}</option>
    {#each parts as part (part)}
      <option value={part}>{part}</option>
    {/each}
  </select>
</div>

{#snippet nodeRow(node: RawNode, depth: number)}
  <div class="node" style="--depth: {depth}">
    <div class="tag">
      <span class="name">{node.prefix ? `${node.prefix}:` : ''}{node.local}</span>
      {#each node.element.attrs as attr (attr.name.local)}
        <label class="attr">
          <span class="attr-name">{attr.name.local}</span>
          <input
            value={attr.value}
            onchange={(e) => setAttr(node, attr.name.local, (e.currentTarget as HTMLInputElement).value)}
          />
        </label>
      {/each}
      <button class="add" title="Add child element" onclick={() => addChild(node)}>+</button>
    </div>
    {#each node.children as child (child)}
      {@render nodeRow(child, depth + 1)}
    {/each}
  </div>
{/snippet}

<div class="inspector">
  {#if trees.length === 0}
    <p class="empty">
      {#if source === ''}
        No inline raw XML yet. Insert an image (or open a document with drawings,
        equations, or shapes) to edit DrawingML / OMML / VML here — or pick an XML
        part above (fontTable, settings, styles, numbering, …) to edit it directly.
      {:else}
        This part has no elements.
      {/if}
    </p>
  {:else}
    {#each trees as tree (tree)}
      {@render nodeRow(tree, 0)}
    {/each}
  {/if}
</div>

<style>
  .picker {
    padding: 6px 8px;
    border-bottom: 1px solid #edebe9;
  }
  .picker select {
    width: 100%;
    height: 26px;
    border: 1px solid #d2d0ce;
    border-radius: 4px;
    padding: 0 6px;
    background: #fff;
    font-size: 12px;
  }
  .inspector {
    font-family: 'Cascadia Code', 'Consolas', monospace;
    font-size: 12px;
    padding: 8px;
    overflow: auto;
    flex: 1;
  }
  .empty { color: #888; font-family: 'Segoe UI', system-ui, sans-serif; line-height: 1.5; }
  .node { padding-left: calc(var(--depth) * 14px); }
  .tag {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 4px;
    padding: 1px 0;
  }
  .name { color: #2b579a; font-weight: 600; }
  .attr { display: inline-flex; align-items: center; gap: 2px; }
  .attr-name { color: #a31515; }
  .attr input {
    width: 72px;
    border: 1px solid #d2d0ce;
    border-radius: 3px;
    padding: 0 3px;
    font: inherit;
  }
  .add {
    border: 1px solid #d2d0ce;
    background: #fff;
    border-radius: 3px;
    cursor: pointer;
    width: 18px;
    height: 18px;
    line-height: 1;
  }
  .add:hover { background: #edebe9; }
</style>
