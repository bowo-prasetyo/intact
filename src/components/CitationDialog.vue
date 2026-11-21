<template>
  <ModalDialog title="Citation" @close="$emit('close')" @submit="insert">
    <label>Citation style
      <select v-model="styleChosen" aria-label="Citation style">
        <option v-for="(s, key) in CITATION_STYLES" :key="key" :value="key">{{ s.name }}</option>
      </select>
    </label>
    <label v-if="sources.length + library.length > 8">Find a source
      <input v-model="filter" type="search" autocomplete="off" aria-label="Find a source" placeholder="Author, title or year" />
    </label>
    <fieldset v-if="sources.length">
      <legend>Cite sources in this document</legend>
      <label v-for="s in shown(sources)" :key="s.tag" class="check source">
        <input v-model="chosen" type="checkbox" :value="s.tag" />
        <span>{{ describe(s) }}</span>
      </label>
    </fieldset>
    <fieldset v-if="library.length">
      <legend>Imported from a file</legend>
      <label v-for="item in shown(library.map((s, i) => ({ ...s, key: `lib:${i}` })))" :key="item.key" class="check source">
        <input v-model="chosen" type="checkbox" :value="item.key" />
        <span>{{ describe(item) }}</span>
      </label>
    </fieldset>
    <label class="check source"><input v-model="chosen" type="checkbox" value="new" /> <span>A new source…</span></label>
    <fieldset v-if="chosen.includes('new')">
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
      <label v-if="draft.type === 'Book'">City <input v-model="draft.city" type="text" autocomplete="off" aria-label="City" /></label>
      <label v-if="draft.type !== 'Book'">Web address (URL or DOI) <input v-model="draft.url" type="text" autocomplete="off" aria-label="Web address" /></label>
    </fieldset>
    <label>Page (optional) <input v-model="page" type="text" autocomplete="off" aria-label="Page cited" placeholder="12" /></label>
    <p v-if="chosen.length > 1" class="hint">The page goes with the first source ticked. Sources are cited in the order they were ticked.</p>
    <p class="preview" aria-live="polite">{{ preview }}</p>
    <p v-if="error" class="error" role="alert">{{ error }}</p>
    <p v-if="imported" class="hint" role="status">{{ imported }}</p>
    <p class="import">
      <button type="button" @click="fileInput.click()">Import from a reference manager…</button>
      <input ref="fileInput" type="file" accept=".bib,.ris,.txt,application/x-bibtex,application/x-research-info-systems" hidden aria-label="BibTeX or RIS file" @change="importFile" />
    </p>
    <p class="hint">A BibTeX (.bib) or RIS (.ris) file exported from Zotero, Mendeley, EndNote or a database, read on this device. Sources are kept in the document’s sources, as Word keeps them (Word’s Manage Sources lists them), only once cited. Insert a bibliography to list them; Update (F9) brings it up to date.</p>
    <template #actions>
      <button type="submit" class="confirm">Insert</button>
    </template>
  </ModalDialog>
</template>

<script setup>
// Insert Citation, as on Word's References tab: cite sources the document has, sources imported
// from a reference manager's file (BibTeX or RIS), or a new one (a book, a journal article or a
// website), one or several together, in APA, MLA, Chicago, Harvard or IEEE style, with the page
// cited if wanted.
import { ref, reactive, computed } from 'vue';
import ModalDialog from './ModalDialog.vue';
import { SOURCE_TYPES, CITATION_STYLES, citationStyle, formatCitation } from '../docx/references.js';
import { parseReferences } from '../docx/bibImport.js';

const props = defineProps({
  sources: { type: Array, default: () => [] }, // the document's sources
  library: { type: Array, default: () => [] }, // sources imported from a file, not yet in the document
  style: { type: String, default: 'APA' },
  numbers: { type: Map, default: () => new Map() }, // IEEE's numbers for the sources cited so far
});
// insert({ picks: [{ tag } | { source, libraryIndex? }], page, style }), in the order cited;
// import(sources): sources read from a file, for the library.
const emit = defineEmits(['close', 'insert', 'import']);
const styleChosen = ref(citationStyle(props.style));
const chosen = ref(props.sources.length ? [props.sources[props.sources.length - 1].tag] : ['new']);
const page = ref('');
const error = ref('');
const imported = ref('');
const filter = ref('');
const fileInput = ref(null);
const draft = reactive({ type: 'Book', authors: '', title: '', year: '', publisher: '', city: '', journal: '', site: '', volume: '', issue: '', pages: '', url: '' });

/** The authors typed ("Smith, Jane; Lee, Kim") as [{ last, first }]. */
const authorsOf = (text) => text.split(';').map((a) => a.trim()).filter(Boolean).map((a) => {
  const [last, ...first] = a.split(',');
  if (first.length) return { last: last.trim(), first: first.join(',').trim() };
  const words = a.split(/\s+/);
  return words.length > 1 ? { last: words.pop(), first: words.join(' ') } : { last: a, first: '' };
});

const newSource = computed(() => ({
  type: draft.type, authors: authorsOf(draft.authors), corporate: '', title: draft.title.trim(), year: draft.year.trim(),
  publisher: draft.type === 'Book' ? draft.publisher.trim() : '', city: draft.type === 'Book' ? draft.city.trim() : '',
  journal: draft.type === 'JournalArticle' ? draft.journal.trim() : '',
  volume: draft.type === 'JournalArticle' ? draft.volume.trim() : '', issue: draft.type === 'JournalArticle' ? draft.issue.trim() : '',
  pages: draft.type === 'JournalArticle' ? draft.pages.trim() : '', site: draft.type === 'InternetSite' ? draft.site.trim() : '',
  url: draft.type === 'Book' ? '' : draft.url.trim(), accessed: '',
}));

/** The source a key in `chosen` stands for: a document's tag, 'lib:3' or 'new'. */
function sourceOf(key) {
  if (key === 'new') return newSource.value;
  if (key.startsWith('lib:')) return props.library[Number(key.slice(4))];
  return props.sources.find((s) => s.tag === key);
}

const preview = computed(() => {
  const list = chosen.value.map(sourceOf).filter((s) => s && (s.title || s.authors.length));
  if (!list.length) return '';
  // IEEE: the numbers so far, and the next ones for sources not yet cited.
  const numbers = new Map(props.numbers);
  const keyed = list.map((s, i) => ({ ...s, tag: s.tag || `\u0000${i}` }));
  for (const s of keyed) if (!numbers.has(s.tag)) numbers.set(s.tag, numbers.size + 1);
  return formatCitation(keyed, styleChosen.value, page.value.trim(), numbers);
});

const describe = (s) => `${s.authors.map((a) => a.last).join(', ') || s.corporate || ''}${s.year ? ` (${s.year})` : ''}${s.title ? ` – ${s.title}` : ''}`;
/** The items whose description has every word of the filter, and those ticked. */
const shown = (items) => {
  const words = filter.value.toLowerCase().split(/\s+/).filter(Boolean);
  if (!words.length) return items;
  return items.filter((s) => chosen.value.includes(s.key || s.tag) || words.every((w) => describe(s).toLowerCase().includes(w)));
};

const sameWork = (a, b) => a.title.trim().toLowerCase() === b.title.trim().toLowerCase() && (a.year || '') === (b.year || '');

/** Reads a BibTeX or RIS file on this device and offers its sources (not those the document or list has). */
async function importFile(event) {
  const file = event.target.files?.[0];
  event.target.value = '';
  if (!file) return;
  error.value = '';
  imported.value = '';
  let found;
  try {
    found = parseReferences(await file.text());
  } catch {
    found = [];
  }
  if (!found.length) {
    error.value = `No sources found in ${file.name}: choose a BibTeX (.bib) or RIS (.ris) file`;
    return;
  }
  const known = [...props.sources, ...props.library];
  const fresh = found.filter((s) => !known.some((k) => sameWork(k, s)));
  const skipped = found.length - fresh.length;
  if (fresh.length) emit('import', fresh);
  const n = fresh.length;
  imported.value = `${n} ${n === 1 ? 'source' : 'sources'} imported from ${file.name}${skipped ? ` (${skipped} already here)` : ''}: tick those to cite. Only sources cited go into the document.`;
}

function insert() {
  error.value = '';
  if (!chosen.value.length) {
    error.value = 'Tick a source to cite';
    return;
  }
  if (chosen.value.includes('new') && !newSource.value.title) {
    error.value = 'Give the source a title';
    return;
  }
  const picks = chosen.value.map((key) => {
    if (key === 'new') return { source: newSource.value };
    if (key.startsWith('lib:')) return { source: props.library[Number(key.slice(4))], libraryIndex: Number(key.slice(4)) };
    return { tag: key };
  });
  emit('insert', { picks, page: page.value.trim(), style: styleChosen.value });
}
</script>

<style scoped>
.source span { color: var(--text); }
.preview { min-height: 18px; margin: 8px 0; font-size: 14px; color: var(--text); }
.import { margin: 8px 0 0; }
</style>
