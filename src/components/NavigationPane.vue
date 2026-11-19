<template>
  <nav class="nav-pane" aria-label="Headings">
    <header>
      <h2>Headings</h2>
      <button type="button" class="close" aria-label="Close the headings" title="Close" @click="$emit('close')">×</button>
    </header>
    <p v-if="!headings.length" class="empty">No headings yet. Give a paragraph a heading style (Heading 1, Heading 2…) to see it here and move around the document by it.</p>
    <ol v-else>
      <li v-for="h in headings" :key="h.index" :class="[`level${Math.min(h.level, 6)}`, { on: h.index === current }]">
        <button type="button" :title="h.text" :aria-current="h.index === current ? 'location' : null" @click="$emit('go', h.index)">{{ h.text }}</button>
      </li>
    </ol>
  </nav>
</template>

<script setup>
// The Navigation pane's Headings, as Word's: the document's headings by level, the one the caret is
// under marked, and a click goes to a heading (the caret at its start).
defineProps({
  headings: { type: Array, default: () => [] }, // [{ index, level (1-9), text }]
  current: { type: Number, default: -1 },
});
defineEmits(['close', 'go']);
</script>

<style scoped>
.nav-pane {
  width: 240px; flex: none; box-sizing: border-box; overflow: auto; border-right: 1px solid var(--border);
  background: var(--surface); color: var(--text); font-size: 13px;
}
header { display: flex; align-items: center; justify-content: space-between; padding: 10px 8px 6px 14px; position: sticky; top: 0; background: var(--surface); }
h2 { margin: 0; font-size: 14px; }
.close { border: 0; background: none; color: inherit; font-size: 20px; line-height: 1; width: 28px; height: 28px; border-radius: 4px; cursor: pointer; }
.close:hover { background: var(--hover); }
.empty { margin: 8px 14px; color: var(--muted); line-height: 1.4; }
ol { list-style: none; margin: 0; padding: 0 6px 12px; }
li button {
  display: block; width: 100%; text-align: left; padding: 4px 8px; border: 0; border-radius: 4px; background: none;
  color: inherit; font: inherit; cursor: pointer; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
}
li button:hover { background: var(--hover); }
li.on button { background: var(--active); color: var(--accent); }
.level1 button { font-weight: 600; }
.level2 button { padding-left: 20px; }
.level3 button { padding-left: 32px; }
.level4 button, .level5 button, .level6 button { padding-left: 44px; color: var(--muted); }
@media (max-width: 700px) { .nav-pane { display: none; } }
</style>
