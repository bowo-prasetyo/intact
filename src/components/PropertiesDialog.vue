<template>
  <ModalDialog title="Document properties" @close="$emit('close')" @submit="save">
    <div class="two">
      <label v-for="f in CORE.filter((c) => c.key !== 'description')" :key="f.key">{{ f.label }}
        <input v-model="core[f.key]" :aria-label="f.label" autocomplete="off" />
      </label>
    </div>
    <label>Comments
      <textarea v-model="core.description" aria-label="Comments" rows="2" />
    </label>
    <fieldset>
      <legend>Custom</legend>
      <p class="hint">Your own properties, such as a product name or a version number, as Word’s Advanced Properties › Custom: change one here and every field showing it follows.</p>
      <div v-for="(p, i) in custom" :key="i" class="row">
        <input v-model="p.name" :aria-label="`Name of custom property ${i + 1}`" autocomplete="off" @input="error = ''" />
        <input v-model="p.value" :aria-label="`Value of ${p.name || `custom property ${i + 1}`}`" autocomplete="off" />
        <button type="button" class="remove" :aria-label="`Remove ${p.name || `custom property ${i + 1}`}`" title="Remove this property" @click="custom.splice(i, 1)">×</button>
      </div>
      <button type="button" class="add" @click="addRow">Add a property</button>
    </fieldset>
    <label v-if="insertable.length">Insert as a field
      <select v-model="field" aria-label="Insert as a field">
        <option v-for="n in insertable" :key="n" :value="n">{{ n }}</option>
      </select>
    </label>
    <p v-if="error" class="error" role="alert">{{ error }}</p>
    <p v-else-if="insertable.length" class="hint">A DOCPROPERTY field at the caret, which Word updates: Update (F9) brings every one up to date, also in headers and footers.</p>
    <template #actions>
      <button v-if="insertable.length" type="button" @click="save(true)">Insert field</button>
      <button type="submit" class="confirm">OK</button>
    </template>
  </ModalDialog>
</template>

<script setup>
// Word's document properties (File › Info › Properties, and Advanced Properties › Custom), and
// Insert › Quick Parts › Field › DocProperty: the title, subject and the like, the document's own
// custom properties (text), and a field showing one at the caret.
import { reactive, ref, computed } from 'vue';
import ModalDialog from './ModalDialog.vue';

// properties: { core: { title, subject, keywords, category, contentStatus, description }, custom: [{ name, value }] }.
const props = defineProps({ properties: { type: Object, required: true } });
// save({ core, custom, field }), field: the property to insert as a field, or null.
const emit = defineEmits(['close', 'save']);
const CORE = [
  { key: 'title', label: 'Title', field: 'Title' }, { key: 'subject', label: 'Subject', field: 'Subject' },
  { key: 'keywords', label: 'Keywords', field: 'Keywords' }, { key: 'category', label: 'Category', field: 'Category' },
  { key: 'contentStatus', label: 'Status', field: null }, { key: 'description', label: 'Comments', field: 'Comments' },
];
const core = reactive({ ...props.properties.core });
const custom = reactive(props.properties.custom.map((p) => ({ ...p })));
const error = ref('');
const insertable = computed(() => [
  ...CORE.filter((c) => c.field).map((c) => c.field),
  ...custom.map((p) => p.name.trim()).filter(Boolean),
]);
const field = ref(custom.find((p) => p.name.trim())?.name.trim() || 'Title');

function addRow() {
  custom.push({ name: '', value: '' });
}

function save(insert = false) {
  const list = custom.map((p) => ({ name: p.name.trim(), value: p.value })).filter((p) => p.name || p.value);
  const names = new Set();
  for (const p of list) {
    const key = p.name.toLowerCase();
    if (!p.name) return void (error.value = 'Give each custom property a name');
    if (p.name.length > 255) return void (error.value = `“${p.name.slice(0, 20)}…” is too long a name: up to 255 characters`);
    if (/^MSIP_Label_/i.test(p.name)) return void (error.value = 'Names starting MSIP_Label_ are a sensitivity label’s, which Office manages');
    if (names.has(key)) return void (error.value = `There are two properties called “${p.name}”: a name is used once`);
    names.add(key);
  }
  emit('save', { core: { ...core }, custom: list, field: insert === true ? field.value : null });
}
</script>

<style scoped>
.row { display: grid; grid-template-columns: 1fr 1.4fr 28px; gap: 6px; margin: 6px 0; }
.row input { min-width: 0; }
.remove { width: 28px; border: 0; background: none; color: inherit; font-size: 18px; border-radius: 4px; cursor: pointer; }
.remove:hover { background: var(--hover); }
.add { margin-top: 4px; padding: 4px 10px; border: 1px solid var(--border); border-radius: 4px; background: var(--surface); color: inherit; font: inherit; cursor: pointer; }
textarea { resize: vertical; }
</style>
