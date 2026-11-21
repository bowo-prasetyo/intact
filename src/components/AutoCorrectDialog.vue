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
// is what to replace with, as Word offers it. Saved when OK is chosen (see editor/autoCorrectList.js).
import { computed, nextTick, onMounted, ref } from 'vue';
import ModalDialog from './ModalDialog.vue';
import { t } from '../i18n.js';
import { TYPOS, entryProblem } from '../editor/autoCorrectList.js';

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
.typos { margin: 4px 0 8px; font-size: 12px; color: var(--muted); }
.typos p { margin: 4px 0; }
@media (max-width: 480px) { .add { grid-template-columns: 1fr; } }
</style>
