<template>
  <form class="fill-in" role="dialog" aria-label="Fill in placeholders" aria-describedby="fill-in-how" @submit.prevent="fill" @keydown.esc="emit('close')">
    <p id="fill-in-how" class="how">
      <template v-if="items.length">Type what goes in each placeholder: it is filled in everywhere it appears. Leave one empty to keep it.</template>
      <template v-else>No placeholders found. Placeholders are text in brackets, such as {{ EXAMPLES }}.</template>
    </p>
    <ul v-if="items.length" class="fields">
      <li v-for="item in items" :key="item.text">
        <label :for="`fill-${item.id}`">{{ item.label }}</label>
        <input
          :id="`fill-${item.id}`"
          v-model="values[item.text]"
          type="text"
          spellcheck="true"
          :placeholder="item.text"
          :disabled="!item.editable"
          :aria-describedby="`fill-${item.id}-where`"
          @focus="show(item)"
        />
        <span :id="`fill-${item.id}-where`" class="where">{{ where(item) }}</span>
      </li>
    </ul>
    <p v-if="note" class="note" role="status">{{ note }}</p>
    <div class="actions">
      <button v-if="items.length" type="submit" class="confirm" :disabled="!ready">Fill in</button>
      <button type="button" @click="emit('close')">{{ items.length ? 'Cancel' : 'Close' }}</button>
    </div>
  </form>
</template>

<script setup>
// Fill in placeholders: every "[Client name]", "{{date}}", «Amount» or <<Amount>> in the document,
// each listed once with a box for what goes in it, filled in everywhere it appears as one undoable
// edit (tracked while tracking), as Replace all would do for each of them in turn.
import { ref, reactive, computed, onMounted, onBeforeUnmount, nextTick } from 'vue';
import { findPlaceholders } from '../editor/placeholders.js';
import { matchRange, replaceMatch } from '../editor/find.js';

const props = defineProps({
  /** The elements to look in, in order (the page: its header, text, notes and footer). */
  roots: { type: Function, required: true },
  /** Runs a document change as one undoable edit; the change returns the text nodes it touched. */
  edit: { type: Function, required: true },
  /** The author and date if a replacement in `node` is a tracked change, or null. */
  tracked: { type: Function, default: () => null },
});
// reveal(range): bring a placeholder into view; filled(message): what was filled in, for the status bar.
const emit = defineEmits(['close', 'reveal', 'filled']);

const EXAMPLES = '[Client name], {{date}}, «Amount» or <<Amount>>';
const items = ref([]);
const values = reactive({});
const note = ref('');
const ready = computed(() => items.value.some((i) => i.editable && values[i.text]?.trim()));

function scan() {
  items.value = findPlaceholders(props.roots()).map((item, id) => ({ ...item, id }));
}

function where(item) {
  const n = item.matches.length;
  const places = `${n} ${n === 1 ? 'place' : 'places'}`;
  const locked = n - item.editable;
  if (!item.editable) return `${places}, read-only here`;
  return locked ? `${places}, ${locked} read-only` : places;
}

const highlights = typeof CSS !== 'undefined' && CSS.highlights && typeof Highlight !== 'undefined';

/** Highlights every place a placeholder is (as Find does), and brings the first into view. */
function show(item) {
  const ranges = item.matches.filter((m) => m.segment.nodes[0].node.isConnected).map(matchRange);
  if (highlights) {
    const all = new Highlight(...ranges);
    all.priority = 2;
    CSS.highlights.set('find-match', all);
  }
  if (ranges[0]) emit('reveal', ranges[0]);
}

function clearPaint() {
  if (highlights) CSS.highlights.delete('find-match');
}

/** Fills in each placeholder given a value, everywhere it can be changed, last place first. */
function fill() {
  scan(); // the document may have changed since the list was made
  const chosen = items.value.filter((i) => i.editable && values[i.text]?.trim());
  if (!chosen.length) return;
  const todo = chosen
    .flatMap((i) => i.matches.filter((m) => m.editable).map((m) => ({ m, text: values[i.text].trim() })))
    .sort((a, b) => b.m.order - a.m.order); // later places first, so earlier offsets stay valid
  props.edit(() => todo.map(({ m, text }) => replaceMatch(m, text, props.tracked(matchRange(m).startContainer)).node));
  for (const i of chosen) delete values[i.text];
  clearPaint();
  const places = todo.length;
  const message = `Filled in ${chosen.length} ${chosen.length === 1 ? 'placeholder' : 'placeholders'} in ${places} ${places === 1 ? 'place' : 'places'}`;
  scan();
  emit('filled', message);
  const left = items.value.filter((i) => i.editable).length;
  if (left) note.value = `${message}. ${left} left.`;
  else emit('close'); // all filled in (any left are read-only here)
}

onMounted(() => {
  scan();
  nextTick(() => document.querySelector('.fill-in input:not(:disabled), .fill-in button')?.focus());
});
onBeforeUnmount(clearPaint);

defineExpose({ scan });
</script>

<style scoped>
/* Its own look, as the editor's dialogs have theirs (the editor's styles reach only this box). */
.fill-in {
  position: fixed; z-index: 20; padding: 12px; min-width: 260px;
  background: var(--surface); border: 1px solid var(--border); border-radius: 6px; box-shadow: 0 4px 16px rgb(0 0 0 / 15%);
}
.fill-in { width: 360px; max-width: calc(100vw - 16px); max-height: calc(100vh - 140px); display: flex; flex-direction: column; gap: 8px; font-size: 13px; }
.how { margin: 0; color: var(--muted); }
.fields { list-style: none; margin: 0; padding: 0 2px; display: flex; flex-direction: column; gap: 8px; overflow-y: auto; min-height: 0; }
.fields li { display: grid; grid-template-columns: 1fr; gap: 2px; }
.fields label { font-weight: 600; overflow-wrap: anywhere; }
.fields input { font: inherit; padding: 4px 6px; border: 1px solid var(--border); border-radius: 4px; background: var(--surface); color: inherit; }
.fields input:disabled { opacity: .55; }
.where { font-size: 12px; color: var(--muted); }
.note { margin: 0; }
.actions { display: flex; gap: 8px; justify-content: flex-end; }
.actions button { font: inherit; border: 1px solid var(--border); border-radius: 4px; padding: 5px 12px; background: var(--surface); color: inherit; cursor: pointer; }
.actions button:hover { background: var(--hover); }
.actions .confirm { background: var(--brand); border-color: var(--brand); color: #fff; }
.actions .confirm:hover { background: var(--brand); filter: brightness(1.1); }
.actions button:disabled { opacity: .5; cursor: default; filter: none; }
</style>
