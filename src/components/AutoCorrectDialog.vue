<template>
  <ModalDialog :title="t('AutoCorrect')" @close="$emit('close')" @submit="save">
    <label class="check"><input v-model="on" type="checkbox" /> {{ t('Replace text as you type') }}</label>
    <p class="hint">{{ t('A word typed then a space or punctuation is replaced, in the case you typed it (cllr → Councillor, Cllr → Councillor). Ctrl+Z straight after gives back what you typed.') }}</p>

    <fieldset>
      <legend>{{ t('Your entries') }}</legend>
      <div class="add">
        <label>{{ t('Replace') }}<input ref="fromInput" v-model="from" type="text" autocomplete="off" spellcheck="false" @keydown.enter.prevent="add" /></label>
        <label>{{ t('With') }}<input v-model="to" type="text" autocomplete="off" @keydown.enter.prevent="add" /></label>
        <button type="button" @click="add">{{ existing ? t('Replace entry') : t('Add') }}</button>
      </div>
      <p v-if="problem" class="error" role="alert">{{ t(problem) }}</p>
      <p v-if="!entries.length" class="hint">{{ t('None yet: such as a long name from its initials, or a word you often mistype.') }}</p>
      <ul v-else class="entries">
        <li v-for="([f, w], i) in entries" :key="f">
          <button type="button" class="pick" :title="t('Change it')" @click="from = f; to = w">{{ f }}</button>
          <span class="arrow" aria-hidden="true">→</span>
          <span class="with">{{ w }}</span>
          <button type="button" class="delete" :aria-label="t('Delete {entry}', { entry: f })" @click="entries.splice(i, 1)">×</button>
        </li>
      </ul>
      <div class="move">
        <button type="button" @click="file.click()">{{ t('Import…') }}</button>
        <button type="button" :disabled="!entries.length" @click="exportEntries">{{ t('Export…') }}</button>
        <input ref="file" type="file" accept=".csv,.tsv,.txt,text/csv,text/tab-separated-values,text/plain" hidden @change="importEntries" />
        <span class="hint">{{ t('A CSV file (Replace, With), to take your entries to another browser or device, or edit them in a spreadsheet') }}</span>
      </div>
      <p v-if="moved" class="note" role="status">{{ moved }}</p>
    </fieldset>

    <label class="check"><input v-model="typos" type="checkbox" /> {{ t('Correct common English typos (teh → the, adn → and…), in English text') }}</label>
    <details class="typos">
      <summary>{{ t('The typos corrected ({n})', { n: TYPOS.length }) }}</summary>
      <p>{{ TYPOS.map(([f, w]) => `${f} → ${w}`).join(', ') }}</p>
    </details>
    <p class="hint">{{ t('Kept in this browser, on this device: another browser or device doesn’t have them.') }}</p>

    <template #actions>
      <button type="submit" class="confirm">{{ t('OK') }}</button>
    </template>
  </ModalDialog>
</template>

<script setup>
// Word's AutoCorrect Options (the status bar's AutoCorrect): "Replace text as you type", your own entries
// (a word and what replaces it) and Word's common English typos. Opened with text selected, that text
// is what to replace with, as Word offers it. Export and Import take the entries to another browser or
// device as a CSV file. Saved when OK is chosen (see editor/autoCorrectList.js).
import { computed, nextTick, onMounted, ref } from 'vue';
import ModalDialog from './ModalDialog.vue';
import { t } from '../i18n.js';
import { TYPOS, entryProblem, entriesToCsv, entriesFromCsv, mergeEntries, MAX_ENTRIES, MAX_FILE } from '../editor/autoCorrectList.js';
import { parseCsv, decodeCsv } from '../docx/merge.js';

const props = defineProps({
  settings: { type: Object, required: true }, // { on, typos, entries }
  selected: { type: String, default: '' }, // the selected text, to replace with
});
const emit = defineEmits(['close', 'save']);

const on = ref(props.settings.on);
const typos = ref(props.settings.typos);
const entries = ref(props.settings.entries.map((e) => [...e]));
const from = ref('');
const to = ref(props.selected);
const problem = ref('');
const fromInput = ref(null);
const file = ref(null);
const moved = ref(''); // what Import or Export did
const existing = computed(() => entries.value.some(([f]) => f.toLowerCase() === from.value.trim().toLowerCase()));

function add() {
  const f = from.value.trim();
  problem.value = entryProblem(f, to.value);
  if (problem.value) return false;
  const at = entries.value.findIndex(([x]) => x.toLowerCase() === f.toLowerCase());
  if (at >= 0) entries.value.splice(at, 1, [f, to.value]);
  else entries.value.push([f, to.value]);
  entries.value.sort((a, b) => a[0].localeCompare(b[0]));
  from.value = '';
  to.value = '';
  fromInput.value?.focus();
  return true;
}

/** Export…: the entries in the list (as it is now, OK'd or not) saved as a CSV file, on this device. */
function exportEntries() {
  const url = URL.createObjectURL(new Blob([entriesToCsv(entries.value)], { type: 'text/csv' }));
  const a = document.createElement('a');
  a.href = url;
  a.download = 'autocorrect.csv';
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10000);
  moved.value = t(entries.value.length === 1 ? 'Exported 1 entry to autocorrect.csv, in your downloads' : 'Exported {n} entries to autocorrect.csv, in your downloads', { n: entries.value.length });
}

/** Import…: a list's entries put into this one (an entry for the same word replaced); kept once OK is chosen. */
async function importEntries(e) {
  const picked = e.target.files?.[0];
  e.target.value = ''; // the same file can be chosen again
  if (!picked) return;
  if (picked.size > MAX_FILE) {
    moved.value = t('{name} is too big for a list of entries (1 MB at most).', { name: picked.name });
    return;
  }
  const { entries: incoming, skipped } = entriesFromCsv(decodeCsv(await picked.arrayBuffer()), parseCsv);
  const merged = mergeEntries(entries.value, incoming);
  if (merged.entries.length > MAX_ENTRIES) {
    moved.value = t('{name} would make more than {max} entries: import a shorter list.', { name: picked.name, max: MAX_ENTRIES });
    return;
  }
  entries.value = merged.entries;
  const parts = [
    t(merged.added === 1 ? '1 entry added' : '{n} entries added', { n: merged.added }),
    merged.changed && t(merged.changed === 1 ? '1 changed' : '{n} changed', { n: merged.changed }),
    skipped && t(skipped === 1 ? '1 row skipped (not an entry: a word to replace with spaces, or nothing to replace it with)' : '{n} rows skipped (not entries: a word to replace with spaces, or nothing to replace it with)', { n: skipped }),
  ].filter(Boolean);
  moved.value = `${t('From {name}:', { name: picked.name })} ${parts.join(', ')}. ${t('Choose OK to keep them.')}`;
}

function save() {
  // An entry typed but not added is added, as OK in Word's dialog does.
  if ((from.value.trim() || to.value.trim()) && !add()) return;
  emit('save', { on: on.value, typos: typos.value, entries: entries.value });
}

onMounted(async () => {
  await nextTick();
  fromInput.value?.focus();
});
</script>

<style scoped>
/* A fieldset is as wide as its content by default: here, as wide as the dialog. */
fieldset { min-width: 0; }
.add { display: grid; grid-template-columns: minmax(0, 1fr) minmax(0, 2fr) auto; gap: 8px; align-items: end; }
.add label { margin: 0; min-width: 0; }
.add input { width: 100%; box-sizing: border-box; }
.add button, .entries button { padding: 6px 12px; border: 1px solid var(--border); border-radius: 4px; background: var(--surface); color: inherit; font: inherit; cursor: pointer; }
.entries { list-style: none; margin: 8px 0 0; padding: 0; max-height: 200px; overflow: auto; }
.entries li { display: grid; grid-template-columns: auto auto 1fr auto; gap: 8px; align-items: center; padding: 3px 0; border-top: 1px solid var(--border); }
.entries .pick { border: 0; padding: 2px 4px; font-weight: 600; text-decoration: underline dotted; background: none; }
.entries .with { overflow-wrap: anywhere; }
.entries .delete { border: 0; padding: 0 6px; font-size: 18px; line-height: 1; background: none; }
.arrow { color: var(--muted); }
.move { display: flex; flex-wrap: wrap; gap: 6px 8px; align-items: center; margin-top: 10px; }
.move button { padding: 4px 10px; border: 1px solid var(--border); border-radius: 4px; background: var(--surface); color: inherit; font: inherit; cursor: pointer; }
.move button:disabled { opacity: .5; cursor: default; }
.move .hint { flex: 1 1 200px; margin: 0; }
.note { margin: 6px 0 0; font-size: 12px; }
.typos { margin: 4px 0 8px; font-size: 12px; color: var(--muted); }
.typos p { margin: 4px 0; }
@media (max-width: 480px) { .add { grid-template-columns: 1fr; } }
</style>
