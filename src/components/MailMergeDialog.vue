<template>
  <div class="mm-backdrop" @mousedown.self="$emit('close')">
    <form class="mm" role="dialog" aria-modal="true" aria-labelledby="mm-title" @submit.prevent="run" @keydown.esc.stop="$emit('close')">
      <header>
        <h2 id="mm-title">Mail merge</h2>
        <button type="button" class="close" aria-label="Close" @click="$emit('close')">×</button>
      </header>

      <template v-if="!placeholders.length">
        <p class="hint">This document has no placeholders for a list to fill in. Type one in brackets where each row’s value goes, such as <strong>[Name]</strong> or <strong>[Address]</strong>, then merge again.</p>
        <div class="actions"><button type="button" @click="$emit('close')">Close</button></div>
      </template>

      <template v-else>
        <p class="hint">Each row of a list fills in the placeholders, for a letter (or a document) per row. Save the list from Excel or Google Sheets as <strong>CSV</strong>, with the column names in its first row. It stays on this device.</p>
        <div class="pick">
          <button type="button" @click="fileInput.click()">{{ list ? 'Choose another list…' : 'Choose a list (.csv)…' }}</button>
          <input ref="fileInput" type="file" accept=".csv,.txt,text/csv,text/plain" hidden aria-label="List to merge (CSV)" @change="onPick" />
          <span v-if="list" class="file" role="status">{{ list.file }}: {{ list.rows.length }} {{ list.rows.length === 1 ? 'row' : 'rows' }}, {{ list.columns.length }} {{ list.columns.length === 1 ? 'column' : 'columns' }}</span>
        </div>
        <p v-if="error" class="error" role="alert">{{ error }}</p>

        <template v-if="list">
          <div class="preview-nav">
            <span>Preview:</span>
            <button type="button" aria-label="Previous row" :disabled="row === 0" @click="row--">‹</button>
            <span>Row {{ row + 1 }} of {{ list.rows.length }}</span>
            <button type="button" aria-label="Next row" :disabled="row >= list.rows.length - 1" @click="row++">›</button>
          </div>
          <table class="map">
            <thead><tr><th scope="col">Placeholder</th><th scope="col">Filled from</th><th scope="col">In this row</th></tr></thead>
            <tbody>
              <tr v-for="(p, i) in placeholders" :key="p.text">
                <th scope="row">{{ p.text }}</th>
                <td>
                  <select v-model.number="mapping[i]" :aria-label="`Column for ${p.text}`">
                    <option :value="-1">(leave as it is)</option>
                    <option v-for="(c, k) in list.columns" :key="k" :value="k">{{ c }}</option>
                  </select>
                </td>
                <td class="value">{{ mapping[i] >= 0 ? (list.rows[row][mapping[i]] || '(empty)') : p.text }}</td>
              </tr>
            </tbody>
          </table>

          <fieldset class="output">
            <legend>Make</legend>
            <label><input v-model="mode" type="radio" value="single" :disabled="!!singleBlocked" /> One document, each row’s copy on a new page</label>
            <p v-if="singleBlocked" class="note">Not for this document: {{ singleBlocked }}.</p>
            <label><input v-model="mode" type="radio" value="separate" /> A separate document for each row, in a .zip</label>
            <label v-if="mode === 'separate'" class="name-by">Name each file by
              <select v-model.number="nameColumn" aria-label="Name each file by">
                <option :value="-1">Number (1, 2, 3…)</option>
                <option v-for="(c, k) in list.columns" :key="k" :value="k">{{ c }}</option>
              </select>
            </label>
          </fieldset>
          <p v-if="unmatched" class="note">{{ unmatched }} {{ unmatched === 1 ? 'placeholder is' : 'placeholders are' }} left as {{ unmatched === 1 ? 'it is' : 'they are' }}: choose a column to fill {{ unmatched === 1 ? 'it' : 'them' }}.</p>
          <p v-if="list.rows.length > MAX_ROWS" class="note">The list has {{ list.rows.length }} rows: the first {{ MAX_ROWS }} are merged.</p>
        </template>

        <div class="actions">
          <button type="submit" class="confirm" :disabled="!ready || busy">{{ busy ? 'Merging…' : list ? `Merge ${Math.min(list.rows.length, MAX_ROWS)} ${Math.min(list.rows.length, MAX_ROWS) === 1 ? 'row' : 'rows'}` : 'Merge' }}</button>
          <button type="button" @click="$emit('close')">Cancel</button>
        </div>
      </template>
    </form>
  </div>
</template>

<script setup>
// Mail merge from a CSV (see docx/merge.js): choose the list, check which column fills each
// placeholder (matched by name, as far as the names agree) with a row to preview, and make one
// document of every row's copy or a .zip of a document for each.
import { ref, computed, onMounted, nextTick } from 'vue';
import { parseCsv, decodeCsv, matchColumns, MAX_ROWS } from '../docx/merge.js';

const props = defineProps({
  /** The document's placeholders: [{ text, label }]. */
  placeholders: { type: Array, required: true },
  /** Why one document of every row's copy can't be made here, or null. */
  singleBlocked: { type: String, default: null },
  /** Makes and downloads the merged documents: ({ rows, mapping, mode, nameColumn }) → Promise. */
  merge: { type: Function, required: true },
});
const emit = defineEmits(['close']);

const fileInput = ref(null);
const list = ref(null);
const mapping = ref([]);
const row = ref(0);
const mode = ref(props.singleBlocked ? 'separate' : 'single');
const nameColumn = ref(-1);
const error = ref('');
const busy = ref(false);

const unmatched = computed(() => mapping.value.filter((m) => m < 0).length);
const ready = computed(() => !!list.value?.rows.length && mapping.value.some((m) => m >= 0));

async function onPick(e) {
  const file = e.target.files[0];
  e.target.value = '';
  if (!file) return;
  error.value = '';
  try {
    const parsed = parseCsv(decodeCsv(await file.arrayBuffer()));
    if (!parsed.rows.length) {
      error.value = `“${file.name}” has no rows below its column names. The first row names the columns; each row after it is one letter.`;
      return;
    }
    list.value = { ...parsed, file: file.name };
    mapping.value = matchColumns(props.placeholders.map((p) => p.label), parsed.columns);
    row.value = 0;
    // A column that names a person or company names the files, where there is one.
    const named = parsed.columns.findIndex((c) => /name|nama|company|perusahaan|client|klien/i.test(c));
    nameColumn.value = named;
  } catch (err) {
    error.value = `“${file.name}” couldn’t be read as a list: ${err.message}`;
  }
}

async function run() {
  if (!ready.value || busy.value) return;
  busy.value = true;
  error.value = '';
  try {
    await props.merge({ rows: list.value.rows, mapping: mapping.value, mode: mode.value, nameColumn: nameColumn.value });
    emit('close');
  } catch (err) {
    error.value = err.message;
  } finally {
    busy.value = false;
  }
}

onMounted(() => nextTick(() => document.querySelector('.mm button:not(.close)')?.focus()));
</script>

<style scoped>
.mm-backdrop {
  position: fixed; inset: 0; z-index: 20; display: flex; align-items: center; justify-content: center;
  padding: 16px; background: rgb(0 0 0 / 35%);
}
.mm {
  width: min(620px, 100%); max-height: calc(100vh - 32px); overflow: auto; padding: 0 20px 18px; box-sizing: border-box;
  background: var(--surface); color: var(--text); border-radius: 8px; box-shadow: 0 12px 40px rgb(0 0 0 / 30%); font-size: 13px;
}
header { display: flex; align-items: center; justify-content: space-between; padding: 16px 0 6px; }
h2 { margin: 0; font-size: 18px; }
.close { border: 0; background: none; color: inherit; font-size: 22px; line-height: 1; width: 32px; height: 32px; border-radius: 4px; cursor: pointer; }
.close:hover { background: var(--hover); }
.hint { margin: 0 0 12px; color: var(--muted); }
.pick { display: flex; flex-wrap: wrap; align-items: center; gap: 10px; }
.file { color: var(--muted); overflow-wrap: anywhere; }
.preview-nav { display: flex; align-items: center; gap: 8px; margin: 14px 0 6px; color: var(--muted); }
.preview-nav button { width: 28px; height: 26px; padding: 0; }
.map { width: 100%; border-collapse: collapse; table-layout: fixed; }
.map th:nth-child(1) { width: 36%; }
.map th:nth-child(2) { width: 30%; }
.map th, .map td { text-align: left; padding: 5px 6px; border-bottom: 1px solid var(--border); vertical-align: top; }
.map thead th { font-size: 12px; color: var(--muted); font-weight: 600; }
.map tbody th { font-weight: 600; overflow-wrap: anywhere; }
.map select { width: 100%; }
.map .value { white-space: pre-line; overflow-wrap: anywhere; color: var(--muted); }
.output { margin: 14px 0 0; padding: 8px 12px 10px; border: 1px solid var(--border); border-radius: 6px; display: flex; flex-direction: column; gap: 6px; }
.output legend { font-weight: 600; font-size: 12px; padding: 0 4px; }
.output label { display: flex; align-items: center; gap: 6px; }
.name-by { margin-left: 22px; }
.note { margin: 6px 0 0; color: var(--muted); font-size: 12px; }
.error { margin: 8px 0 0; color: var(--danger, #b3261e); }
.actions { display: flex; gap: 8px; margin-top: 16px; }
button:not(.close), select {
  padding: 5px 10px; border: 1px solid var(--border); border-radius: 4px; background: var(--surface); color: inherit; font: inherit; cursor: pointer;
}
button:not(.close):hover:not(:disabled) { background: var(--hover); }
button:disabled { opacity: .5; cursor: default; }
.confirm { background: var(--brand) !important; border-color: var(--brand) !important; color: #fff !important; }
button:focus-visible, select:focus-visible, input:focus-visible { outline: 2px solid var(--accent); outline-offset: 2px; }
</style>
