<template>
  <ModalDialog title="Mark index entry" @close="$emit('close')" @submit="mark(false)">
    <label>Main entry
      <input ref="mainEl" v-model="main" aria-label="Main entry" autocomplete="off" @input="error = ''" />
    </label>
    <label>Subentry
      <input v-model="sub" aria-label="Subentry" autocomplete="off" placeholder="Optional: listed under the main entry" />
    </label>
    <p v-if="error" class="error" role="alert">{{ error }}</p>
    <p v-else class="hint">An index entry (Word’s XE field) after the selection: it shows nothing in the text. Insert › Index lists the entries with their pages; Update (F9) brings it up to date.{{ count ? ` The document has ${count} ${count === 1 ? 'entry' : 'entries'} marked.` : '' }}</p>
    <template #actions>
      <button type="button" :disabled="!text" :title="text ? `Mark “${text}” in each paragraph it is in` : 'Select the text to mark everywhere first'" @click="mark(true)">Mark all</button>
      <button type="submit" class="confirm">Mark</button>
    </template>
  </ModalDialog>
</template>

<script setup>
// Word's Mark Index Entry (References › Mark Entry, Alt+Shift+X): an entry, and a subentry under it,
// for the index, marked after the selection, or (Mark all) after the selected text in each
// paragraph it is in.
import { ref, onMounted } from 'vue';
import ModalDialog from './ModalDialog.vue';

// text: the selected text (one line), or ''; count: the entries marked already.
const props = defineProps({ text: { type: String, default: '' }, count: { type: Number, default: 0 } });
const emit = defineEmits(['close', 'mark']);
const main = ref(props.text);
const sub = ref('');
const error = ref('');
const mainEl = ref(null);

function mark(all) {
  if (!main.value.trim()) return void (error.value = 'Type the entry, as the index is to list it');
  if (/[:"]/.test(main.value + sub.value)) return void (error.value = 'An entry can’t have a colon or quotation marks: Word reads a colon as the start of a subentry');
  emit('mark', { main: main.value.trim(), sub: sub.value.trim(), all });
}
onMounted(() => {
  mainEl.value?.focus();
  mainEl.value?.select();
});
</script>
