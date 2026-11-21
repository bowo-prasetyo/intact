<template>
  <ModalDialog :title="t('Thesaurus')" @close="$emit('close')" @submit="look(query)">
    <div class="search">
      <label>{{ t('Look up') }}
        <input ref="input" v-model="query" type="search" autocomplete="off" spellcheck="false" />
      </label>
      <label>{{ t('Language') }}
        <select v-model="language" @change="look(query)">
          <option value="en">English</option>
          <option value="id">Bahasa Indonesia</option>
        </select>
      </label>
      <button type="submit" class="go">{{ t('Look up') }}</button>
    </div>

    <p v-if="state === 'loading'" class="hint" role="status">{{ t('Opening the thesaurus…') }}</p>
    <p v-else-if="state === 'failed'" class="error" role="status">{{ t('The thesaurus couldn’t be opened. Reload the page and try again.') }}</p>
    <template v-else-if="result">
      <p v-if="!result.meanings.length" class="hint" role="status">{{ t('No meanings found for “{word}”.', { word: looked }) }} {{ language === 'id' ? t('The Indonesian thesaurus is small: it has only words checked by hand.') : '' }}</p>
      <template v-else>
        <p class="hint" role="status">{{ replacing ? t('Click a word to put it in place of “{word}”.', { word: replacing }) : t('Click a word to insert it at the caret.') }}{{ result.inflected ? ` ${t('Words for “{base}” are shown as “{word}” is: {example}.', { base: baseOf(result), word: looked, example: exampleOf(result) })}` : '' }}</p>
        <section v-for="(m, i) in result.meanings.slice(0, LIMIT)" :key="i" class="meaning">
          <h3>{{ m.words[0] || m.similar[0] || m.opposite[0] }} <span class="pos">({{ t(m.label) }}{{ m.form !== looked.toLowerCase() ? ` · ${m.form}` : '' }})</span></h3>
          <div v-if="m.words.length" class="words">
            <button v-for="w in m.words" :key="w" type="button" @click="$emit('insert', w)">{{ w }}</button>
          </div>
          <div v-if="m.similar.length" class="words related">
            <span class="kind">{{ t('Similar') }}</span>
            <button v-for="w in m.similar.slice(0, 12)" :key="w" type="button" @click="$emit('insert', w)">{{ w }}</button>
          </div>
          <div v-if="m.opposite.length" class="words related">
            <span class="kind">{{ t('Opposite') }}</span>
            <button v-for="w in m.opposite" :key="w" type="button" @click="$emit('insert', w)">{{ w }}</button>
          </div>
        </section>
        <p v-if="result.meanings.length > LIMIT" class="hint">{{ t('{n} more meanings: look up a word from those above to narrow it down.', { n: result.meanings.length - LIMIT }) }}</p>
      </template>
    </template>
    <p class="credit">{{ language === 'id' ? t('Words from Wordnet Bahasa (MIT licence).') : t('Words from WordNet, Princeton University.') }} {{ t('On this device: nothing is sent anywhere.') }}</p>
  </ModalDialog>
</template>

<script setup>
// Word's Thesaurus (Review › Thesaurus, Shift+F7): the word at the caret (or selected) looked up,
// its meanings with the words that say each, and similar and opposite ones; a click puts a word in
// its place (see EditorView's replaceWithSynonym). Any other word can be looked up. The thesaurus
// is loaded from the app's own files the first time (see editor/thesaurus.js).
import { nextTick, onMounted, ref } from 'vue';
import ModalDialog from './ModalDialog.vue';
import { t } from '../i18n.js';
import { loadThesaurus } from '../editor/thesaurus.js';

const props = defineProps({
  word: { type: String, default: '' }, // the word to look up first, and to replace
  lang: { type: String, default: 'en' }, // 'en' or 'id'
  participle: { type: Boolean, default: false }, // the word replaced follows have or be: an -ed form is a past participle
});
defineEmits(['close', 'insert']);

const LIMIT = 12; // meanings shown
const query = ref(props.word);
const language = ref(props.lang);
const replacing = props.word;
const state = ref('');
const result = ref(null);
const looked = ref('');
const input = ref(null);

/** The base of the first meaning shown in another form ("approve" for "approves"). */
const baseOf = (r) => r.meanings.find((m) => m.kind !== 'base').form;
/** One word shown in another form, as an example: "sanctions". */
const exampleOf = (r) => {
  const m = r.meanings.find((x) => x.kind !== 'base');
  return m.words[0] || m.similar[0] || m.opposite[0];
};

async function look(word) {
  const w = word.trim();
  if (!w) return;
  state.value = 'loading';
  try {
    const thesaurus = await loadThesaurus(language.value);
    // Only the word in the text has words before it to tell a participle by.
    result.value = thesaurus.lookup(w, { participle: props.participle && w.toLowerCase() === replacing.toLowerCase() });
    looked.value = w;
    state.value = '';
  } catch {
    state.value = 'failed';
  }
}

onMounted(async () => {
  await nextTick();
  input.value?.focus();
  input.value?.select();
  if (query.value) look(query.value);
});
</script>

<style scoped>
.search { display: grid; grid-template-columns: minmax(0, 1fr) auto auto; gap: 8px; align-items: end; }
.search label { margin: 0; min-width: 0; }
.search input { width: 100%; box-sizing: border-box; }
.go { padding: 6px 12px; border: 1px solid var(--border); border-radius: 4px; background: var(--surface); color: inherit; font: inherit; cursor: pointer; }
.meaning { margin: 10px 0; padding-top: 8px; border-top: 1px solid var(--border); }
h3 { margin: 0 0 6px; font-size: 13px; }
.pos { color: var(--muted); font-weight: normal; }
.words { display: flex; flex-wrap: wrap; gap: 4px 6px; align-items: center; margin-top: 4px; }
.words button { padding: 2px 8px; border: 1px solid var(--border); border-radius: 12px; background: var(--surface); color: inherit; font: inherit; cursor: pointer; }
.words button:hover, .words button:focus-visible { background: var(--hover); border-color: var(--accent); }
.kind { color: var(--muted); font-size: 12px; margin-right: 2px; }
.credit { margin: 12px 0 0; color: var(--muted); font-size: 11px; }
@media (max-width: 480px) { .search { grid-template-columns: 1fr; } }
</style>
