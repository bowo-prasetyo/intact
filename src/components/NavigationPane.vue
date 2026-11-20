<template>
  <nav class="nav-pane" aria-label="Headings">
    <header>
      <h2>Headings</h2>
      <button type="button" class="close" aria-label="Close the headings" title="Close" @click="$emit('close')">×</button>
    </header>
    <p v-if="!headings.length" class="empty">No headings yet. Give a paragraph a heading style (Heading 1, Heading 2…) to see it here and move around the document by it.</p>
    <ol v-else>
      <li
        v-for="h in headings"
        :key="h.index"
        :class="[`level${Math.min(h.level, 6)}`, { on: h.index === current, dragging: h.index === dragFrom, target: h.index === dropAt && dragFrom !== null && h.index !== dragFrom }]"
        :draggable="movable"
        @dragstart="onDragStart($event, h.index)"
        @dragend="dragFrom = null; dropAt = null"
        @dragover="onDragOver($event, h.index)"
        @dragleave="dropAt === h.index && (dropAt = null)"
        @drop="onDrop($event, h.index)"
      >
        <button type="button" class="go" :title="movable ? `${h.text} – drag to move it with its text, or Alt+Shift+↑ / ↓` : h.text" :aria-current="h.index === current ? 'location' : null" @click="$emit('go', h.index)" @keydown="onKeydown($event, h.index)">{{ h.text }}</button>
        <span v-if="movable" class="moves">
          <button type="button" aria-label="Move up" title="Move up, with its text (Alt+Shift+↑)" @click="$emit('move', h.index, -1)"><svg viewBox="0 0 10 10" width="10" height="10" aria-hidden="true"><path d="M2 6.5 5 3.5l3 3" stroke="currentColor" stroke-width="1.5" fill="none" /></svg></button>
          <button type="button" aria-label="Move down" title="Move down, with its text (Alt+Shift+↓)" @click="$emit('move', h.index, 1)"><svg viewBox="0 0 10 10" width="10" height="10" aria-hidden="true"><path d="M2 3.5 5 6.5l3-3" stroke="currentColor" stroke-width="1.5" fill="none" /></svg></button>
        </span>
      </li>
    </ol>
  </nav>
</template>

<script setup>
// The Navigation pane's Headings, as Word's: the document's headings by level, the one the caret is
// under marked, and a click goes to a heading (the caret at its start). As in Word, a heading can be
// moved with everything under it (a chapter with its text): dragged onto another heading, it goes
// before that one; ▲ and ▼ (or Alt+Shift+↑ and ↓) move it past the heading of its level before or
// after it. The editor does the moving (see moveHeading), or says why it can't.
import { ref } from 'vue';

defineProps({
  headings: { type: Array, default: () => [] }, // [{ index, level (1-9), text }]
  current: { type: Number, default: -1 },
  movable: { type: Boolean, default: false },
});
const emit = defineEmits(['close', 'go', 'move', 'move-before']);

const dragFrom = ref(null);
const dropAt = ref(null);

function onDragStart(e, i) {
  dragFrom.value = i;
  e.dataTransfer.effectAllowed = 'move';
  e.dataTransfer.setData('application/x-intact-heading', String(i));
}

function onDragOver(e, i) {
  if (dragFrom.value === null) return; // not a heading of this pane: a file, text…
  e.preventDefault();
  e.dataTransfer.dropEffect = 'move';
  dropAt.value = i;
}

function onDrop(e, i) {
  if (dragFrom.value === null) return;
  e.preventDefault();
  const from = dragFrom.value;
  dragFrom.value = null;
  dropAt.value = null;
  if (from !== i) emit('move-before', from, i);
}

function onKeydown(e, i) {
  if (!e.altKey || !e.shiftKey || e.ctrlKey || e.metaKey) return;
  if (e.key !== 'ArrowUp' && e.key !== 'ArrowDown') return;
  e.preventDefault();
  emit('move', i, e.key === 'ArrowUp' ? -1 : 1);
}
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
li { display: flex; align-items: center; border-top: 2px solid transparent; }
li.dragging { opacity: .5; }
li.target { border-top-color: var(--accent); } /* where the heading dragged goes: before this one */
li .go {
  display: block; flex: 1; min-width: 0; text-align: left; padding: 4px 8px; border: 0; border-radius: 4px; background: none;
  color: inherit; font: inherit; cursor: pointer; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
}
li .go:hover { background: var(--hover); }
li.on .go { background: var(--active); color: var(--accent); }
.moves { display: flex; flex: none; opacity: 0; }
li:hover .moves, li:focus-within .moves { opacity: 1; }
.moves button {
  width: 20px; height: 22px; padding: 0; border: 0; border-radius: 3px; background: none; color: var(--muted);
  display: grid; place-items: center; cursor: pointer;
}
.moves button:hover { background: var(--hover); color: inherit; }
.level1 .go { font-weight: 600; }
.level2 .go { padding-left: 20px; }
.level3 .go { padding-left: 32px; }
.level4 .go, .level5 .go, .level6 .go { padding-left: 44px; color: var(--muted); }
@media (max-width: 700px) { .nav-pane { display: none; } }
</style>
