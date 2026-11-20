<template>
  <ModalDialog :title="t('Word count')" @close="$emit('close')" @submit="save">
    <dl class="counts">
      <template v-if="pages"><dt>{{ t('Pages') }}</dt><dd>≈ {{ pages }}</dd></template>
      <dt>{{ t('Words') }}</dt><dd>{{ n(words) }}</dd>
      <dt>{{ t('Characters (no spaces)') }}</dt><dd>{{ n(characters.withoutSpaces) }}</dd>
      <dt>{{ t('Characters (with spaces)') }}</dt><dd>{{ n(characters.withSpaces) }}</dd>
      <dt>{{ t('Paragraphs') }}</dt><dd>{{ n(stats.paragraphs) }}</dd>
      <dt>{{ t('Sentences') }}</dt><dd>{{ n(stats.sentences) }}</dd>
    </dl>
    <p class="hint">{{ t('Headers, footers and notes aren’t counted, as in Word.') }}</p>

    <fieldset>
      <legend>{{ t('Length') }}</legend>
      <div class="two">
        <label>{{ t('Target, in words') }}
          <input v-model="wordsTarget" type="number" min="1" step="1" inputmode="numeric" :placeholder="t('None')" />
        </label>
        <label>{{ t('Headline limit, in characters') }}
          <input v-model="headlineLimit" type="number" min="1" step="1" inputmode="numeric" :placeholder="t('None')" />
        </label>
      </div>
      <p class="hint">{{ toGo }}</p>
      <p class="hint">{{ headline ? t('Headline: {chars} characters, with spaces – “{text}”', { chars: headline.chars, text: short(headline.text) }) : t('No headline: the limit counts the text’s first paragraph when it is a heading (Title or Heading 1, Ctrl+Alt+1).') }}</p>
      <p class="hint">{{ t('The status bar shows them as you write. They stay with this document while it is open here (and in its unsaved copy), never in the file.') }}</p>
    </fieldset>

    <fieldset>
      <legend>{{ t('Readability') }}</legend>
      <dl class="counts">
        <dt>{{ t('Sentences per paragraph') }}</dt><dd>{{ stats.sentencesPerParagraph }}</dd>
        <dt>{{ t('Words per sentence') }}</dt><dd>{{ stats.wordsPerSentence }}</dd>
        <dt>{{ t('Characters per word') }}</dt><dd>{{ stats.charactersPerWord }}</dd>
        <template v-if="stats.scores">
          <dt>{{ t('Passive sentences') }}</dt><dd>{{ stats.scores.passive }}%</dd>
          <dt>{{ t('Flesch Reading Ease') }}</dt><dd>{{ stats.scores.ease }} <span class="note">({{ t(easeLabel(stats.scores.ease)) }})</span></dd>
          <dt>{{ t('Flesch-Kincaid Grade Level') }}</dt><dd>{{ stats.scores.grade }}</dd>
        </template>
      </dl>
      <p v-if="!stats.scores" class="hint">{{ t('The passive voice and the Flesch scores are worked out for English text, as in Word.') }}</p>
      <p v-else class="hint">{{ t('Reading Ease runs from 0 to 100: higher is easier (60–70 is plain English). The grade is the years of school a reader needs.') }}</p>
      <div class="long">
        <span>{{ stats.long.length ? t(stats.long.length === 1 ? '1 long sentence (more than {limit} words)' : '{n} long sentences (more than {limit} words)', { n: stats.long.length, limit: LONG_SENTENCE }) : t('No sentences of more than {limit} words', { limit: LONG_SENTENCE }) }}</span>
        <button v-if="stats.long.length" type="button" @click="$emit('show-long')">{{ t('Show') }}</button>
      </div>
    </fieldset>

    <template #actions>
      <button type="submit" class="confirm">{{ t('OK') }}</button>
    </template>
  </ModalDialog>
</template>

<script setup>
// Word's Word Count dialog (a click on the status bar's count): the document's counts, a length to
// write to (a target in words, and a limit for the headline in characters) and Word's Readability
// Statistics, with the long sentences to show one by one. Nothing here changes the document.
import { computed, ref } from 'vue';
import ModalDialog from './ModalDialog.vue';
import { t } from '../i18n.js';
import { easeLabel, LONG_SENTENCE } from '../editor/readability.js';

const props = defineProps({
  words: { type: Number, required: true },
  characters: { type: Object, required: true }, // { withSpaces, withoutSpaces }
  pages: { type: Number, default: 0 }, // about how many, or 0 where they aren't counted
  stats: { type: Object, required: true }, // from readability()
  target: { type: Object, required: true }, // { words, headline }: numbers or null
  headline: { type: Object, default: null }, // { text, chars }
});
const emit = defineEmits(['close', 'save', 'show-long']);

const wordsTarget = ref(props.target.words ?? '');
const headlineLimit = ref(props.target.headline ?? '');
const n = (x) => x.toLocaleString();
const short = (s) => (s.length > 60 ? `${s.slice(0, 60)}…` : s);
const whole = (v) => (Number.isInteger(Number(v)) && Number(v) > 0 ? Number(v) : null);

const toGo = computed(() => {
  const target = whole(wordsTarget.value);
  if (!target) return t('No target: set one to see how far you are from it.');
  const d = target - props.words;
  if (d > 0) return t('{n} words to go', { n: n(d) });
  if (d < 0) return t('{n} words over: cut to fit', { n: n(-d) });
  return t('On target');
});

function save() {
  emit('save', { words: whole(wordsTarget.value), headline: whole(headlineLimit.value) });
}
</script>

<style scoped>
.counts { display: grid; grid-template-columns: 1fr auto; gap: 4px 16px; margin: 8px 0; }
.counts dt { color: var(--muted); }
.counts dd { margin: 0; text-align: right; font-variant-numeric: tabular-nums; }
.note { color: var(--muted); }
.long { display: flex; align-items: center; justify-content: space-between; gap: 12px; margin-top: 6px; }
.long button { padding: 3px 10px; border: 1px solid var(--border); border-radius: 4px; background: var(--surface); color: inherit; font: inherit; cursor: pointer; }
.long button:hover { background: var(--hover); }
</style>
