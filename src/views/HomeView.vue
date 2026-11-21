<template>
  <!-- A file dropped anywhere on the page opens, not just on the Open tile (see onDrop). -->
  <div
    class="home"
    data-file-drop
    @dragenter="onDragEnter"
    @dragover="onDragOver"
    @dragleave="onDragLeave"
    @drop="onDrop"
  >
    <AppHeader show-language />
    <main class="home-main">
      <h1>{{ t('Change what you need.') }}<br /> {{ t('Keep the rest') }} <span class="brand-word">{{ t('intact') }}</span>.</h1>
      <p class="lead">{{ t('Open a') }} <code>.docx</code>{{ t(', fix it and download it again, or start a new one, ready for Word, right here in your browser.') }}</p>
      <ul class="promises" :aria-label="t('What Intact promises')">
        <li v-for="p in promises" :key="p.title">
          <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M3.5 8.5l3 3 6-7" /></svg>
          <span><strong>{{ t(p.title) }}</strong> – {{ t(p.text) }}</span>
        </li>
      </ul>

      <section v-if="drafts.length" class="drafts" aria-labelledby="drafts-title">
        <h2 id="drafts-title">{{ t('Unsaved changes on this device') }}</h2>
        <p class="drafts-note">{{ t('Kept in this browser only, until you download or close the document, or for {days} days.', { days: maxDays }) }}</p>
        <ul>
          <li v-for="d in drafts" :key="d.id">
            <span class="draft-name">{{ d.name }}</span>
            <span class="draft-when">{{ ago(d.savedAt) }}</span>
            <span v-if="d.protected" class="draft-when">{{ t('encrypted with its password') }}</span>
            <!-- Being edited in another tab: restoring it here would make a second copy. -->
            <span v-if="openElsewhere.has(d.id)" class="draft-open" :title="t('This document is open in another tab of this browser: switch to that tab to go on with it')">{{ t('Open in another tab') }}</span>
            <template v-else>
              <button type="button" class="restore" :disabled="busy" :aria-label="t('Restore {name}', { name: d.name })" @click="restore(d)">{{ t('Restore') }}</button>
              <button type="button" class="discard" :disabled="busy" :aria-label="t('Discard the changes to {name}', { name: d.name })" @click="discard(d)">{{ t('Discard') }}</button>
            </template>
          </li>
        </ul>
      </section>

      <div class="actions">
        <button class="tile" type="button" :disabled="busy" @click="createNew">
          <svg class="tile-icon" viewBox="0 0 24 24" width="32" height="32" aria-hidden="true"><path d="M6.5 3h7l5 5v11.5A1.5 1.5 0 0 1 17 21H6.5A1.5 1.5 0 0 1 5 19.5v-15A1.5 1.5 0 0 1 6.5 3z" /><path d="M13.5 3v5h5M12 11v6M9 14h6" /></svg>
          <span class="tile-title">{{ t('Blank document') }}</span>
          <span class="tile-sub">{{ t('Write a new document') }}</span>
        </button>

        <!-- A real button, so the file picker can be reached with the keyboard; the input stays hidden. -->
        <button
          class="tile drop"
          :class="{ over: dragOver }"
          type="button"
          :disabled="busy"
          aria-labelledby="open-title"
          aria-describedby="open-hint"
          @click="browse"
        >
          <svg class="tile-icon" viewBox="0 0 24 24" width="32" height="32" aria-hidden="true"><path d="M3 7.5V18a1.5 1.5 0 0 0 1.5 1.5h15A1.5 1.5 0 0 0 21 18V9a1.5 1.5 0 0 0-1.5-1.5h-7.2L10.5 5h-6A1.5 1.5 0 0 0 3 6.5z" /><path d="M12 16.5v-6M9.5 13 12 10.5l2.5 2.5" /></svg>
          <span id="open-title" class="tile-title">{{ dragOver ? t('Drop to open') : t('Open a .docx file') }}</span>
          <span id="open-hint" class="tile-sub">{{ t('Click to browse, or drop a file anywhere here') }}</span>
        </button>
        <input ref="fileInput" type="file" :accept="ACCEPT" hidden :disabled="busy" @change="onPick" />
      </div>

      <!-- A document encrypted with a password: asked for here, checked and decrypted on this device. -->
      <form v-if="pw.open" class="password" aria-labelledby="pw-title" @submit.prevent="submitPassword" @keydown.esc="cancelPassword">
        <h2 id="pw-title">{{ t('Protected with a password') }}</h2>
        <p id="pw-what" class="password-note">{{ pw.text }}</p>
        <label for="doc-password">{{ t('Password for {name}', { name: pw.name }) }}</label>
        <div class="password-row">
          <input id="doc-password" ref="pwInput" v-model="pw.value" type="password" autocomplete="off" spellcheck="false" required :disabled="busy" :aria-invalid="pw.error ? 'true' : 'false'" aria-describedby="pw-what pw-error" />
          <button type="submit" class="restore" :disabled="busy || !pw.value">{{ t('Open') }}</button>
          <button type="button" :disabled="busy" @click="cancelPassword">{{ t('Cancel') }}</button>
        </div>
        <p id="pw-error" class="password-error" role="alert">{{ pw.error }}</p>
      </form>

      <!-- Starter documents: Word documents made here, with [placeholders] to fill in. In Indonesian,
           the official documents (naskah dinas) come first, in Indonesian; the others are in English. -->
      <section v-for="group in starterGroups" :key="group.id" class="starters" :aria-labelledby="`starters-title-${group.id}`">
        <h2 :id="`starters-title-${group.id}`">{{ t(group.title) }}</h2>
        <ul>
          <li v-for="s in group.starters" :key="s.id">
            <button type="button" class="starter" :disabled="busy" :aria-label="`${t(s.title)}: ${t(s.description)}`" @click="createFrom(s.id)">
              <svg class="starter-icon" viewBox="0 0 24 24" width="28" height="28" aria-hidden="true"><path v-for="(d, i) in STARTER_ICONS[s.id]" :key="i" :d="d" /></svg>
              <span class="starter-title">{{ t(s.title) }}</span>
              <span class="starter-sub">{{ t(s.description) }}</span>
            </button>
          </li>
        </ul>
        <p class="starters-note">{{ group.note }}</p>
      </section>

      <p v-if="busy" class="status" role="status">{{ pw.open ? t('Checking the password…') : t('Opening…') }}</p>
      <p v-if="error" class="error" role="alert">{{ error }}</p>
    </main>
  </div>
</template>

<script setup>
import { ref, reactive, computed, nextTick, onMounted, onBeforeUnmount } from 'vue';
import { t, lang } from '../i18n.js';
import { useRouter, useRoute } from 'vue-router';
import AppHeader from '../components/AppHeader.vue';
import { openFile, newDocument, newFromStarter, openDraft } from '../store.js';
import { STARTERS, localPaper } from '../docx/starters.js';
import { listDrafts, deleteDraft, MAX_AGE } from '../drafts.js';
import { receiveFile, draftsOpenElsewhere } from '../tabs.js';
import { canOpenWithHandle, pickFileToOpen, handleOfDropped } from '../editor/fileHandles.js';

const router = useRouter();
const route = useRoute();
const busy = ref(false);
const error = ref('');
const dragOver = ref(false);
const fileInput = ref(null);

// Drafts of documents left with unsaved changes (a reload, a crash, a closed tab): see drafts.js.
// Those being edited in other tabs are shown as open there (see tabs.js), looked at again when
// this tab is shown, as the others may have opened or closed one meanwhile.
const drafts = ref([]);
const openElsewhere = ref(new Set());
const maxDays = Math.round(MAX_AGE / 86400000);
async function refreshDrafts() {
  const [list, open] = await Promise.all([listDrafts(), draftsOpenElsewhere()]);
  drafts.value = list;
  openElsewhere.value = open;
}
function onVisible() {
  if (document.visibilityState === 'visible') refreshDrafts();
}
onMounted(() => {
  refreshDrafts();
  document.addEventListener('visibilitychange', onVisible);
  window.addEventListener('focus', refreshDrafts);
});
onBeforeUnmount(() => {
  document.removeEventListener('visibilitychange', onVisible);
  window.removeEventListener('focus', refreshDrafts);
});

// A document handed to this tab by the one that opened it (a file dropped on its page: see
// tabs.js): asked for by the token in this tab's address, and opened.
onMounted(async () => {
  const id = route.query.receive;
  if (typeof id !== 'string' || !id) return;
  router.replace({ name: 'home' }); // a reload doesn't ask again
  busy.value = true;
  const file = await receiveFile(id);
  busy.value = false;
  if (file) load(file);
  else error.value = t('The document couldn’t be brought over from the other tab (it may have been closed). Open it here instead.');
});
const restore = (d) => { cancelPassword(); run((password) => openDraft(d, password), d.name); };
async function discard(d) {
  await deleteDraft(d.id);
  drafts.value = drafts.value.filter((x) => x.id !== d.id);
}
function ago(at) {
  const relative = new Intl.RelativeTimeFormat(lang.value === 'id' ? 'id' : undefined, { numeric: 'auto' });
  const minutes = Math.round((at - Date.now()) / 60000);
  if (minutes > -1) return t('just now');
  if (minutes > -60) return relative.format(minutes, 'minute');
  if (minutes > -60 * 24) return relative.format(Math.round(minutes / 60), 'hour');
  return relative.format(Math.round(minutes / 1440), 'day');
}

const ACCEPT = [
  '.docx', '.dotx', '.docm', '.dotm',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.template',
  'application/vnd.ms-word.document.macroEnabled.12',
  'application/vnd.ms-word.template.macroEnabled.12',
].join(',');

// The promises from VISION.md. Keep them true: a claim here that a release breaks is a bug.
const promises = [
  { title: 'Private', text: 'never uploaded' },
  { title: 'Faithful', text: 'only your edits change' },
  { title: 'Instant', text: 'nothing to install' },
];

// A password asked for: { open, name, text, value, error, task (opens the document with a password) }.
const pw = reactive({ open: false, name: '', text: '', value: '', error: '', task: null });
const pwInput = ref(null);

/**
 * Runs `task` (password => opens a document) and goes to the editor. A document encrypted with a
 * password asks for it, and opens once it is right; the password is used for this and nothing else.
 */
async function run(task, name = '') {
  busy.value = true;
  error.value = '';
  try {
    await task(pw.open ? pw.value : null);
    Object.assign(pw, { open: false, value: '', error: '', task: null });
    router.push({ name: 'edit' });
  } catch (e) {
    if (e.name === 'PasswordError' && (e.reason === 'needed' || e.reason === 'wrong')) {
      if (!pw.open) Object.assign(pw, { open: true, name, text: t('It is encrypted. Enter its password to open it: the password is checked on this device and never sent anywhere, and the document stays encrypted with it when you download it.'), value: '', error: '', task });
      else pw.error = t('That password isn’t right. Passwords are case-sensitive: check Caps Lock.');
      busy.value = false; // the field is enabled again before it takes the focus
      await nextTick();
      pwInput.value?.scrollIntoView({ block: 'center' });
      pwInput.value?.focus();
      pwInput.value?.select();
      return;
    }
    console.error(e);
    Object.assign(pw, { open: false, value: '', error: '', task: null });
    // A file Intact recognises but can't open (encrypted, rights-managed, .doc) says so as it is.
    error.value = ['UnreadableFileError', 'PasswordError'].includes(e.name) ? e.message : t('Could not open the document: {error}', { error: e.message });
  } finally {
    busy.value = false;
  }
}

function submitPassword() {
  if (pw.value && pw.task && !busy.value) run(pw.task, pw.name);
}
function cancelPassword() {
  Object.assign(pw, { open: false, value: '', error: '', task: null });
}

function load(file, handle = null) {
  if (!file) return;
  if (!/\.(docx|dotx|docm|dotm)$/i.test(file.name)) {
    error.value = t('Please choose a Word file: .docx, or a .dotx template, or a .docm or .dotm with macros (older .doc files are not supported).');
    return;
  }
  cancelPassword();
  run((password) => openFile(file, handle, password), file.name);
}

const onPick = (e) => { load(e.target.files[0]); e.target.value = ''; };

/**
 * Open a .docx file: with the browser's own file picker where it can give a handle to the file, so
 * Save writes back to it (see editor/fileHandles.js); else the file input, and Download.
 */
async function browse() {
  if (!canOpenWithHandle()) {
    fileInput.value.click();
    return;
  }
  try {
    const picked = await pickFileToOpen();
    if (picked) load(picked.file, picked.handle);
  } catch (e) {
    console.error(e);
    fileInput.value.click(); // the picker refused (a policy, a sandbox): the input still works
  }
}

// Dragging files over the page shows where they go: the Open tile lights up. dragenter and dragleave
// fire for every child crossed, so count them rather than trust the last one.
const hasFiles = (e) => [...(e.dataTransfer?.types || [])].includes('Files');
let dragDepth = 0;
function onDragEnter(e) {
  if (!hasFiles(e)) return;
  dragDepth += 1;
  dragOver.value = true;
}
function onDragOver(e) {
  if (!hasFiles(e)) return;
  e.preventDefault(); // allows the drop here
  e.dataTransfer.dropEffect = busy.value ? 'none' : 'copy';
}
function onDragLeave(e) {
  if (!hasFiles(e)) return;
  dragDepth = Math.max(0, dragDepth - 1);
  if (!dragDepth) dragOver.value = false;
}
function onDrop(e) {
  if (!hasFiles(e)) return;
  e.preventDefault(); // the browser would otherwise open or download the file itself
  dragDepth = 0;
  dragOver.value = false;
  if (busy.value) return;
  const file = e.dataTransfer.files[0];
  // Asked for now, while the drop is handled: the file's handle, so Save writes back to it.
  const handle = handleOfDropped([...e.dataTransfer.items].find((i) => i.kind === 'file'));
  handle.then((h) => load(file, h));
}
const createNew = () => { cancelPassword(); run(newDocument); };
const createFrom = (id) => { cancelPassword(); run(() => newFromStarter(id)); };
const paper = localPaper();
// The starters, in groups: in Indonesian, the official documents first, then those for teachers in
// Indonesian, then the English ones; else the English ones, then the English ones for teachers.
const starterGroups = computed(() => {
  const english = { id: 'en', title: lang.value === 'id' ? 'Other templates (in English)' : 'Or start from a template', starters: STARTERS.filter((s) => !s.lang && !s.group), note: t('{paper} paper, as used where you are. Fill in the [placeholders] once, for everywhere they appear.', { paper }) };
  if (lang.value !== 'id') {
    return [english, { id: 'teach', title: 'For teachers', starters: STARTERS.filter((s) => !s.lang && s.group === 'teach'), note: t('Tests number their questions and choices as Word does; the letter to parents is ready for Mail merge from a class list.') }];
  }
  return [
    { id: 'id', title: 'Official documents (Tata Naskah Dinas)', starters: STARTERS.filter((s) => s.lang === 'id' && !s.group), note: t('A4 paper, as the rules for official documents ask; F4 and Folio are in Page setup. Fill in the [placeholders] once, for everywhere they appear.') },
    { id: 'guru', title: 'For teachers (Kurikulum Merdeka)', starters: STARTERS.filter((s) => s.lang === 'id' && s.group === 'teach'), note: t('A4 paper. The test numbers its questions and choices as Word does; the letter to parents is ready for Mail merge from a class list.') },
    english,
  ];
});

// Line drawings for the starters, on a page outline (24 × 24, stroked).
const PAGE = 'M6.5 3h11A1.5 1.5 0 0 1 19 4.5v15a1.5 1.5 0 0 1-1.5 1.5h-11A1.5 1.5 0 0 1 5 19.5v-15A1.5 1.5 0 0 1 6.5 3z';
const STARTER_ICONS = {
  letter: ['M3.5 6.5h17v11h-17z', 'M3.5 6.5l8.5 6.5 8.5-6.5'],
  memo: [PAGE, 'M8 7.5h3M8 10h8M8 12.5h8M8 15h5'],
  minutes: [PAGE, 'M8 8h.01M8 12h.01M8 16h.01M10.5 8h5.5M10.5 12h5.5M10.5 16h5.5'],
  quote: [PAGE, 'M8 8h8M8 11.5h4M14 11.5h2M8 14.5h4M14 14.5h2M13 17.5h3'],
  invoice: ['M6 3h12v18l-2-1.5-2 1.5-2-1.5-2 1.5-2-1.5L6 21z', 'M9 8h6M9 11.5h6M9 15h3'],
  cv: [PAGE, 'M12 10.2a2 2 0 1 0 0-4 2 2 0 0 0 0 4zM8.5 14.5c.6-1.6 2-2.5 3.5-2.5s2.9.9 3.5 2.5M8 17.5h8'],
  report: [PAGE, 'M8.5 17v-3M12 17v-6M15.5 17V8.5'],
  // A board paper: the decision asked for in its box, then the text.
  'board-paper': [PAGE, 'M7.5 6.5h9v3.5h-9zM9 8.2h6M8 12.5h8M8 14.5h8M8 16.5h5'],
  manuscript: [PAGE, 'M9.5 7.5h5M8 11h8M8 13.5h8M8 16h8M8 18.5h5'],
  // The naskah dinas: a letterhead's line over the text; a memo's labels; an invitation's calendar; two signatures.
  'surat-dinas': [PAGE, 'M8 6.5h8M7.5 8.5h9M8 11h4M8 13h8M8 15h8M13 18h3'],
  'nota-dinas': [PAGE, 'M9 6.5h6M8 9h2M11 9h5M8 11h2M11 11h5M8 13h2M11 13h5M7.5 15h9M8 17.5h6'],
  'surat-undangan': [PAGE, 'M8 7h8M8.5 10.5h7v6h-7zM8.5 12.5h7M10.5 9.5v2M13.5 9.5v2'],
  'berita-acara': [PAGE, 'M9 6.5h6M8 9h8M8 11h8M8 13h6M8 17h3M13 17h3'],
  // A decree's title over its labelled considerations; an assignment's letterhead and its person's
  // details; a circular's numbered points; minutes' details and two signatures; a report's lettered
  // parts; a staff study's parts under its labels; a power of attorney's two signatures.
  keputusan: [PAGE, 'M9.5 6.5h5M10 8.5h4M8 11h2.5M11.5 11h4.5M8 13h2.5M11.5 13h4.5M8 15h2.5M11.5 15h4.5M13 18h3'],
  'surat-tugas': [PAGE, 'M8 6.5h8M7.5 8.5h9M10 10.5h4M8 13h2M11 13h5M8 15h2M11 15h5M13 18h3'],
  'surat-edaran': [PAGE, 'M8 6.5h8M7.5 8.5h9M8 11h.01M10 11h6M8 13.5h.01M10 13.5h6M8 16h.01M10 16h6'],
  notula: [PAGE, 'M9.5 6.5h5M8 9h2M11 9h5M8 11h2M11 11h5M8 13.5h8M8 15.5h8M8.5 18h3M13 18h3'],
  laporan: [PAGE, 'M9.5 6.5h5M8 9h.01M9.5 9h6.5M9.5 11h6.5M8 13.5h.01M9.5 13.5h6.5M9.5 15.5h6.5M13 18h3'],
  'telaahan-staf': [PAGE, 'M9.5 6.5h5M8 8.5h8M8 11h.01M9.5 11h6.5M8 13h.01M9.5 13h6.5M8 15h.01M9.5 15h6.5M13 18h3'],
  'surat-kuasa': [PAGE, 'M8 6.5h8M7.5 8.5h9M8 11h8M8 13h8M8 15h5M8.5 18h3M13 18h3'],
  // For teachers: a test's two columns of choices; a worksheet's tick; a rubric's grid; a lesson's
  // stages; a letter home with a child; and the Indonesian ones, as their English counterparts.
  test: [PAGE, 'M8 7h8M8 9.5h.01M8.5 11h2M8.5 12.5h2M13 9.5h.01M13.5 11h2M13.5 12.5h2M8 15.5h8M8 17.5h8'],
  worksheet: [PAGE, 'M8 7.5h8M8 10.5l1.2 1.2 2.3-2.4M13 11h3M8 14.5h8M8 17h5'],
  rubric: ['M3.5 6h17v12h-17z', 'M3.5 9.5h17M3.5 13.5h17M8 6v12M12 6v12M16 6v12'],
  'lesson-plan': [PAGE, 'M8 7.5h8M8 10.5h2M11.5 10.5h4.5M8 13.5h2M11.5 13.5h4.5M8 16.5h2M11.5 16.5h4.5'],
  'parent-letter': ['M3.5 6.5h17v11h-17z', 'M3.5 6.5l8.5 6.5 8.5-6.5', 'M17.5 3.5a1.5 1.5 0 1 0 0 .01'],
  soal: [PAGE, 'M8 7h8M8 9.5h.01M8.5 11h2M8.5 12.5h2M13 9.5h.01M13.5 11h2M13.5 12.5h2M8 15.5h8M8 17.5h8'],
  'kisi-kisi': ['M3.5 6h17v12h-17z', 'M3.5 9h17M6 6v12M10 6v12M14 6v12M17.5 6v12'],
  'modul-ajar': [PAGE, 'M9 6.5h6M8 9h3M8 11h8M8 13h8M8 15h3M8 17h8'],
  rubrik: ['M3.5 6h17v12h-17z', 'M3.5 9.5h17M3.5 13.5h17M8 6v12M12 6v12M16 6v12'],
  'surat-orang-tua': [PAGE, 'M8 6.5h8M7.5 8.5h9M8 11h5M8 13h8M8 15h8M8.5 18h3M13 18h3'],
};
</script>

<style scoped>
.draft-open { color: var(--muted); font-size: .9em; font-style: italic; }
.home { min-height: 100vh; display: flex; flex-direction: column; }
.home-main { max-width: 760px; margin: 0 auto; padding: 48px 16px; text-align: center; }
h1 { font-size: clamp(1.6rem, 4.5vw, 2.5rem); line-height: 1.2; letter-spacing: -0.01em; margin: 0 0 14px; text-wrap: balance; }
.brand-word { color: var(--accent); }
.lead { color: var(--muted); font-size: 1.05rem; margin: 0 0 16px; text-wrap: balance; }
.promises {
  display: flex; flex-wrap: wrap; justify-content: center; gap: 8px 22px;
  list-style: none; padding: 0; margin: 0 0 36px; font-size: .95rem;
}
.promises li { display: flex; align-items: center; gap: 6px; color: var(--muted); }
.promises strong { color: var(--text); font-weight: 600; }
.promises svg { flex: none; fill: none; stroke: var(--accent); stroke-width: 2; stroke-linecap: round; stroke-linejoin: round; }
.actions { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 20px; }
.tile {
  display: flex; flex-direction: column; align-items: center; gap: 6px; padding: 32px 20px;
  border: 1px solid var(--border); border-radius: 12px; background: var(--surface);
  cursor: pointer; font: inherit; color: inherit; transition: border-color .15s, box-shadow .15s;
}
.tile:hover, .drop.over { border-color: var(--accent); box-shadow: 0 4px 16px rgb(0 0 0 / 8%); }
.drop { border-style: dashed; }
.tile:focus-visible { outline: 2px solid var(--accent); outline-offset: 2px; }
.tile:disabled { cursor: progress; opacity: .7; }
.tile-icon { flex: none; fill: none; stroke: var(--accent); stroke-width: 1.6; stroke-linecap: round; stroke-linejoin: round; }
.tile-title { font-weight: 600; font-size: 1.1rem; }
.tile-sub { color: var(--muted); font-size: .9rem; }
.drafts {
  margin: 0 0 24px; padding: 14px 16px; text-align: left;
  border: 1px solid var(--border); border-left: 4px solid var(--accent); border-radius: 8px; background: var(--surface);
}
.drafts h2 { font-size: 1rem; margin: 0 0 2px; }
.drafts-note { margin: 0 0 8px; color: var(--muted); font-size: .9rem; }
.drafts ul { list-style: none; margin: 0; padding: 0; }
.drafts li { display: flex; flex-wrap: wrap; align-items: center; gap: 4px 12px; padding: 6px 0; border-top: 1px solid var(--border); }
.drafts li:first-child { border-top: 0; }
.draft-name { font-weight: 600; overflow-wrap: anywhere; }
.draft-when { color: var(--muted); font-size: .9rem; margin-right: auto; }
.drafts button { font: inherit; border-radius: 6px; padding: 5px 14px; cursor: pointer; border: 1px solid var(--border); background: var(--surface); color: inherit; }
.drafts button.restore { background: var(--accent); border-color: var(--accent); color: var(--surface); font-weight: 600; }
.drafts button:hover:not(:disabled) { filter: brightness(1.08); }
.drafts button:focus-visible { outline: 2px solid var(--accent); outline-offset: 2px; }
.starters { margin-top: 40px; text-align: left; }
.starters h2 { font-size: 1.05rem; margin: 0 0 12px; text-align: center; }
.starters ul { list-style: none; margin: 0; padding: 0; display: grid; grid-template-columns: repeat(auto-fill, minmax(160px, 1fr)); gap: 12px; }
.starter {
  width: 100%; height: 100%; display: flex; flex-direction: column; align-items: flex-start; gap: 4px; padding: 14px;
  border: 1px solid var(--border); border-radius: 10px; background: var(--surface); color: inherit; font: inherit; text-align: left;
  cursor: pointer; transition: border-color .15s, box-shadow .15s;
}
.starter:hover { border-color: var(--accent); box-shadow: 0 4px 16px rgb(0 0 0 / 8%); }
.starter:focus-visible { outline: 2px solid var(--accent); outline-offset: 2px; }
.starter:disabled { cursor: progress; opacity: .7; }
.starter-icon { flex: none; margin-bottom: 4px; fill: none; stroke: var(--accent); stroke-width: 1.5; stroke-linecap: round; stroke-linejoin: round; }
.starter-title { font-weight: 600; }
.starter-sub { color: var(--muted); font-size: .85rem; line-height: 1.35; }
.starters-note { margin: 12px 0 0; color: var(--muted); font-size: .85rem; text-align: center; }
.status { margin-top: 24px; color: var(--muted); }
.error { margin-top: 24px; color: var(--danger); }
.password {
  margin-top: 16px; text-align: left; padding: 14px 16px; border: 1px solid var(--border); border-radius: 8px; background: var(--surface);
}
.password h2 { font-size: 1rem; margin: 0 0 4px; }
.password-note { margin: 0 0 10px; color: var(--muted); font-size: .9rem; }
.password label { display: block; font-weight: 600; margin-bottom: 4px; overflow-wrap: anywhere; }
.password-row { display: flex; flex-wrap: wrap; gap: 8px; }
.password input { flex: 1 1 200px; min-width: 0; font: inherit; padding: 6px 8px; border: 1px solid var(--border); border-radius: 6px; background: var(--surface); color: inherit; }
.password input[aria-invalid="true"] { border-color: var(--danger); }
.password button { font: inherit; border-radius: 6px; padding: 5px 14px; cursor: pointer; border: 1px solid var(--border); background: var(--surface); color: inherit; }
.password button.restore { background: var(--accent); border-color: var(--accent); color: var(--surface); font-weight: 600; }
.password button:disabled { opacity: .6; cursor: default; }
.password-error { margin: 6px 0 0; color: var(--danger); min-height: 1em; }
.password-error:empty { display: none; }
</style>
