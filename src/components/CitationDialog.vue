<template>
  <ModalDialog title="Citation" @close="$emit('close')" @submit="insert">
    <label>Citation style
      <select v-model="styleChosen" aria-label="Citation style">
        <option value="APA">APA (7th edition)</option>
        <option value="MLA">MLA (9th edition)</option>
      </select>
    </label>
    <fieldset v-if="sources.length">
      <legend>Cite a source in this document</legend>
      <label v-for="s in sources" :key="s.tag" class="check source">
        <input v-model="chosen" type="radio" name="source" :value="s.tag" />
        <span>{{ describe(s) }}</span>
      </label>
      <label class="check source"><input v-model="chosen" type="radio" name="source" value="" /> <span>A new source…</span></label>
    </fieldset>
    <fieldset v-if="!chosen">
      <legend>New source</legend>
      <label>Type
        <select v-model="draft.type" aria-label="Type of source">
          <option v-for="(label, value) in SOURCE_TYPES" :key="value" :value="value">{{ label }}</option>
        </select>
      </label>
      <label>Authors <input v-model="draft.authors" type="text" autocomplete="off" aria-label="Authors" placeholder="Smith, Jane; Lee, Kim" /></label>
      <p class="hint">Last name, first name; another author after a semicolon. Leave empty for an organisation’s work and name it as the publisher or website.</p>
      <label>Title <input v-model="draft.title" type="text" autocomplete="off" aria-label="Title" /></label>
      <div class="two">
        <label>Year <input v-model="draft.year" type="text" inputmode="numeric" autocomplete="off" aria-label="Year" /></label>
        <label v-if="draft.type === 'Book'">Publisher <input v-model="draft.publisher" type="text" autocomplete="off" aria-label="Publisher" /></label>
        <label v-else-if="draft.type === 'JournalArticle'">Journal <input v-model="draft.journal" type="text" autocomplete="off" aria-label="Journal" /></label>
        <label v-else>Website <input v-model="draft.site" type="text" autocomplete="off" aria-label="Website name" /></label>
      </div>
      <div v-if="draft.type === 'JournalArticle'" class="two">
        <label>Volume <input v-model="draft.volume" type="text" autocomplete="off" aria-label="Volume" /></label>
        <label>Issue <input v-model="draft.issue" type="text" autocomplete="off" aria-label="Issue" /></label>
        <label>Pages <input v-model="draft.pages" type="text" autocomplete="off" aria-label="Pages" placeholder="45-67" /></label>
      </div>
      <label v-if="draft.type !== 'Book'">Web address (URL or DOI) <input v-model="draft.url" type="text" autocomplete="off" aria-label="Web address" /></label>
    </fieldset>
    <label>Page (optional) <input v-model="page" type="text" autocomplete="off" aria-label="Page cited" placeholder="12" /></label>
    <p class="preview" aria-live="polite">{{ preview }}</p>
    <p v-if="error" class="error" role="alert">{{ error }}</p>
    <p class="hint">Kept in the document’s sources, as Word keeps them: Word’s Manage Sources lists them. Insert a bibliography to list the sources cited; Update (F9) brings it up to date.</p>
    <template #actions>
      <button type="submit" class="confirm">Insert</button>
    </template>
  </ModalDialog>
</template>

<script setup>
// Insert Citation, as on Word's References tab: cite a source the document has, or a new one (a
// book, a journal article or a website), in APA or MLA style, with the page cited if wanted.
import { ref, reactive, computed } from 'vue';
import ModalDialog from './ModalDialog.vue';
import { SOURCE_TYPES, formatCitation } from '../docx/references.js';

const props = defineProps({ sources: { type: Array, default: () => [] }, style: { type: String, default: 'APA' } });
// insert({ tag } | { source }, page, style): the source cited (one the document has, or a new one).
const emit = defineEmits(['close', 'insert']);
const styleChosen = ref(props.style === 'MLA' ? 'MLA' : 'APA');
const chosen = ref(props.sources.length ? props.sources[props.sources.length - 1].tag : '');
const page = ref('');
const error = ref('');
const draft = reactive({ type: 'Book', authors: '', title: '', year: '', publisher: '', journal: '', site: '', volume: '', issue: '', pages: '', url: '' });

/** The authors typed ("Smith, Jane; Lee, Kim") as [{ last, first }]. */
const authorsOf = (text) => text.split(';').map((a) => a.trim()).filter(Boolean).map((a) => {
  const [last, ...first] = a.split(',');
  if (first.length) return { last: last.trim(), first: first.join(',').trim() };
  const words = a.split(/\s+/);
  return words.length > 1 ? { last: words.pop(), first: words.join(' ') } : { last: a, first: '' };
});

const newSource = computed(() => ({
  type: draft.type, authors: authorsOf(draft.authors), corporate: '', title: draft.title.trim(), year: draft.year.trim(),
  publisher: draft.type === 'Book' ? draft.publisher.trim() : '', city: '', journal: draft.type === 'JournalArticle' ? draft.journal.trim() : '',
  volume: draft.type === 'JournalArticle' ? draft.volume.trim() : '', issue: draft.type === 'JournalArticle' ? draft.issue.trim() : '',
  pages: draft.type === 'JournalArticle' ? draft.pages.trim() : '', site: draft.type === 'InternetSite' ? draft.site.trim() : '',
  url: draft.type === 'Book' ? '' : draft.url.trim(), accessed: '',
}));
const current = computed(() => (chosen.value ? props.sources.find((s) => s.tag === chosen.value) : newSource.value));
const preview = computed(() => (current.value && (current.value.title || current.value.authors.length) ? formatCitation(current.value, styleChosen.value, page.value.trim()) : ''));

const describe = (s) => `${s.authors.map((a) => a.last).join(', ') || s.corporate || ''}${s.year ? ` (${s.year})` : ''}${s.title ? ` – ${s.title}` : ''}`;

function insert() {
  error.value = '';
  if (chosen.value) {
    emit('insert', { tag: chosen.value, page: page.value.trim(), style: styleChosen.value });
    return;
  }
  const s = newSource.value;
  if (!s.title) {
    error.value = 'Give the source a title';
    return;
  }
  emit('insert', { source: s, page: page.value.trim(), style: styleChosen.value });
}
</script>

<style scoped>
.source span { color: var(--text); }
.preview { min-height: 18px; margin: 8px 0; font-size: 14px; color: var(--text); }
</style>
