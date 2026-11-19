<template>
  <ModalDialog title="Cross-reference" @close="$emit('close')" @submit="insert">
    <label>Reference type
      <select v-model="kind" aria-label="Reference type" @change="index = 0; as = defaultAs()">
        <option v-for="k in kinds" :key="k" :value="k">{{ k }}</option>
      </select>
    </label>
    <label>Insert reference to
      <select v-model="as" aria-label="Insert reference to">
        <template v-if="kind === 'Heading'">
          <option value="text">Heading text</option>
        </template>
        <template v-else>
          <option value="number">Only label and number</option>
          <option value="text">Entire caption</option>
        </template>
        <option value="position">Above/below</option>
      </select>
    </label>
    <p v-if="!list.length" class="hint">{{ kind === 'Heading' ? 'The document has no headings: give a paragraph a heading style first.' : `The document has no ${kind.toLowerCase()} captions.` }}</p>
    <div v-else class="items" role="listbox" :aria-label="`For which ${kind.toLowerCase()}`" tabindex="0" @keydown="onKeys">
      <div v-for="(item, i) in list" :key="i" role="option" :aria-selected="index === i" :class="{ on: index === i, [`level${item.level || 1}`]: true }" @click="index = i" @dblclick="index = i; insert()">{{ item.text }}</div>
    </div>
    <p class="hint">A field Word updates (REF): Update (F9) brings its text up to date after the heading or caption changes.</p>
    <template #actions>
      <button type="submit" class="confirm" :disabled="!list.length">Insert</button>
    </template>
  </ModalDialog>
</template>

<script setup>
// Insert Cross-reference, as on Word's References tab: to a heading (its text) or a caption (its
// label and number, or all of it), or "above"/"below", as a REF field to a bookmark on it.
import { ref, computed } from 'vue';
import ModalDialog from './ModalDialog.vue';

// targets: { Heading: [{ text, level }], Figure: [{ text }], … } in the document's order.
const props = defineProps({ targets: { type: Object, required: true } });
// insert({ kind, index, as }).
const emit = defineEmits(['close', 'insert']);
const kinds = computed(() => ['Heading', ...Object.keys(props.targets).filter((k) => k !== 'Heading')]);
const kind = ref(Object.keys(props.targets).find((k) => k !== 'Heading' && props.targets[k].length) || 'Heading');
const defaultAs = () => (kind.value === 'Heading' ? 'text' : 'number');
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
