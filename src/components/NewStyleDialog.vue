<template>
  <ModalDialog title="Create a new style" @close="$emit('close')" @submit="create">
    <label>Name
      <input ref="nameEl" v-model="name" aria-label="Style name" maxlength="253" autocomplete="off" @input="error = ''" />
    </label>
    <label>Style type
      <select v-model="type" aria-label="Style type">
        <option value="paragraph">Paragraph</option>
        <option value="character">Character</option>
      </select>
    </label>
    <p v-if="error" class="error" role="alert">{{ error }}</p>
    <p v-else class="hint">{{ type === 'paragraph'
      ? `From this paragraph’s formatting (its font, size, colour, spacing, indents, borders and shading), based on ${basedOn}, and given to the selected paragraphs. Change it later with Modify style…`
      : 'From the selected text’s formatting (its font, size, colour and the like), and given to it: for code, a product’s name or a term.' }}</p>
    <template #actions>
      <button type="submit" class="confirm">OK</button>
    </template>
  </ModalDialog>
</template>

<script setup>
// Word's Create New Style from Formatting (the style gallery's Create a Style): a paragraph or
// character style made from the selection's formatting, as a "Code" or "Note" style is made.
import { ref, onMounted } from 'vue';
import ModalDialog from './ModalDialog.vue';

// names: the names styles have already (lower case); basedOn: the paragraph's style, by name.
const props = defineProps({ names: { type: Array, required: true }, basedOn: { type: String, default: 'Normal' }, selection: { type: Boolean, default: false } });
const emit = defineEmits(['close', 'create']);
const name = ref('');
const type = ref(props.selection ? 'character' : 'paragraph');
const error = ref('');
const nameEl = ref(null);

function create() {
  const n = name.value.trim();
  if (!n) return void (error.value = 'Type a name for the style');
  if (props.names.includes(n.toLowerCase())) return void (error.value = `The document has a style called “${n}” already: choose another name, or change that one with Modify style…`);
  if (/[,;]/.test(n)) return void (error.value = 'Word separates a style’s other names with commas: leave them out of its name');
  emit('create', { name: n, type: type.value });
}
onMounted(() => nameEl.value?.focus());
</script>
