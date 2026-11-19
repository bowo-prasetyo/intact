<template>
  <ModalDialog title="Caption" @close="$emit('close')" @submit="insert">
    <label>Label
      <select v-model="label" aria-label="Label" @change="position = label === 'Table' ? 'above' : 'below'">
        <option v-for="l in CAPTION_LABELS" :key="l" :value="l">{{ l }}</option>
      </select>
    </label>
    <label>Caption text <input ref="input" v-model="text" type="text" autocomplete="off" aria-label="Caption text" placeholder="Sales by region" /></label>
    <label>Position
      <select v-model="position" aria-label="Position">
        <option value="below">Below {{ target }}</option>
        <option value="above">Above {{ target }}</option>
      </select>
    </label>
    <p class="preview" aria-live="polite"><span>{{ label }} {{ next }}</span>{{ text.trim() ? `: ${text.trim()}` : '' }}</p>
    <p class="hint">Numbered as Word numbers captions (a SEQ field): Update (F9) renumbers them after one is added or moved, and a cross-reference can point to one.</p>
    <template #actions>
      <button type="submit" class="confirm">Insert</button>
    </template>
  </ModalDialog>
</template>

<script setup>
// Insert Caption, as on Word's References tab: "Figure 1: text" in Word's Caption style, below a
// picture or above a table as Word places them by default, numbered by a SEQ field.
import { ref, computed, onMounted, nextTick } from 'vue';
import ModalDialog from './ModalDialog.vue';
import { CAPTION_LABELS } from '../docx/references.js';

// counts: how many captions of each label the document has; target: what the caption is for.
const props = defineProps({ counts: { type: Object, default: () => ({}) }, target: { type: String, default: 'the paragraph' }, table: Boolean });
// insert({ label, text, position }).
const emit = defineEmits(['close', 'insert']);
const label = ref(props.table ? 'Table' : 'Figure');
const position = ref(props.table ? 'above' : 'below');
const text = ref('');
const input = ref(null);
const next = computed(() => (props.counts[label.value.toLowerCase()] || 0) + 1);

function insert() {
  emit('insert', { label: label.value, text: text.value.trim(), position: position.value });
}
onMounted(() => nextTick(() => input.value?.focus()));
</script>

<style scoped>
.preview { margin: 8px 0; font-size: 14px; color: var(--text); font-style: italic; }
</style>
