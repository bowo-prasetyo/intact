<template>
  <ModalDialog title="Cross-reference" @close="$emit('close')" @submit="insert">
    <label>Reference type
      <select v-model="kind" aria-label="Reference type" @change="index = 0; as = defaultAs()">
        <option v-for="k in kinds" :key="k" :value="k">{{ k }}</option>
      </select>
    </label>
    <label>Insert reference to
      <select v-model="as" aria-label="Insert reference to">
        <option v-for="o in options" :key="o.value" :value="o.value">{{ o.label }}</option>
      </select>
    </label>
    <p v-if="!list.length" class="hint">{{ EMPTY[kind] || `The document has no ${kind.toLowerCase()} captions.` }}</p>
    <div v-else class="items" role="listbox" :aria-label="`For which ${kind.toLowerCase()}`" tabindex="0" @keydown="onKeys">
      <div v-for="(item, i) in list" :key="i" role="option" :aria-selected="index === i" :class="{ on: index === i, [`level${item.level || 1}`]: true }" @click="index = i" @dblclick="index = i; insert()">{{ item.text }}</div>
    </div>
    <p class="hint">A field Word updates (REF): Update (F9) brings its text up to date after what it points to changes.</p>
    <template #actions>
      <button type="submit" class="confirm" :disabled="!list.length">Insert</button>
    </template>
  </ModalDialog>
</template>

<script setup>
// Insert Cross-reference, as on Word's References tab: to a numbered item (its number or text), a
// heading (its text, or its number when headings are numbered), a bookmark (its text, or the
// number of its paragraph) or a caption (its label and number, or all of it), or "above"/"below",
// as a REF field to a bookmark on it (\r for a paragraph's number, \p for above/below).
import { ref, computed } from 'vue';
import ModalDialog from './ModalDialog.vue';

// targets: { 'Numbered item': [{ text }], Heading: [{ text, level, numbered }], Bookmark: [{ text,
// numbered }], Figure: [{ text }], … } in the document's order.
const props = defineProps({ targets: { type: Object, required: true } });
// insert({ kind, index, as }), as: 'text', 'number' or 'position'.
const emit = defineEmits(['close', 'insert']);
const OWN = ['Numbered item', 'Heading', 'Bookmark'];
const EMPTY = {
  Heading: 'The document has no headings: give a paragraph a heading style first.',
  'Numbered item': 'The document has no numbered paragraphs: number a list first.',
  Bookmark: 'The document has no bookmarks: add one with Insert > Bookmark… first.',
};
const kinds = computed(() => [...OWN.filter((k) => k in props.targets), ...Object.keys(props.targets).filter((k) => !OWN.includes(k))]);
// As before these kinds were offered: a caption kind with captions first, else headings.
const kind = ref(Object.keys(props.targets).find((k) => !OWN.includes(k) && props.targets[k].length) || 'Heading');
const options = computed(() => {
  const anyNumbered = (props.targets[kind.value] || []).some((t) => t.numbered);
  const out = {
    'Numbered item': [['number', 'Paragraph number'], ['text', 'Paragraph text']],
    Heading: [['text', 'Heading text'], ...(anyNumbered ? [['number', 'Heading number']] : [])],
    Bookmark: [['text', 'Bookmark text'], ...(anyNumbered ? [['number', 'Paragraph number']] : [])],
  }[kind.value] || [['number', 'Only label and number'], ['text', 'Entire caption']];
  return [...out, ['position', 'Above/below']].map(([value, label]) => ({ value, label }));
});
const defaultAs = () => options.value[0].value;
const as = ref(defaultAs());
const index = ref(0);
const list = computed(() => props.targets[kind.value] || []);

function onKeys(e) {
  const step = { ArrowDown: 1, ArrowUp: -1 }[e.key];
  if (!step) return;
  e.preventDefault();
  index.value = Math.max(0, Math.min(list.value.length - 1, index.value + step));
}
function insert() {
  if (!list.value.length) return;
  emit('insert', { kind: kind.value, index: index.value, as: as.value });
}
</script>

<style scoped>
.items { max-height: 200px; overflow: auto; border: 1px solid var(--border); border-radius: 4px; margin: 8px 0; }
.items div { padding: 4px 8px; cursor: pointer; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.items div:hover { background: var(--hover); }
.items div.on { background: var(--active); color: var(--accent); }
.items .level2 { padding-left: 20px; }
.items .level3 { padding-left: 32px; }
.items .level4, .items .level5, .items .level6 { padding-left: 44px; }
.items:focus-visible { outline: 2px solid var(--accent); }
</style>
