<template>
  <div class="findbar" role="search" @keydown.esc.prevent="close">
    <div class="row">
      <input
        ref="findInput"
        v-model="query"
        class="field"
        type="text"
        placeholder="Find"
        aria-label="Find"
        spellcheck="false"
        @keydown.enter.prevent="$event.shiftKey ? go(-1) : go(1)"
      />
      <span class="status" aria-live="polite">{{ status }}</span>
      <button type="button" title="Previous (Shift+Enter)" aria-label="Previous match" :disabled="!editableMatches.length" @click="go(-1)">↑</button>
      <button type="button" title="Next (Enter)" aria-label="Next match" :disabled="!editableMatches.length" @click="go(1)">↓</button>
      <label class="opt"><input v-model="matchCase" type="checkbox" /> Match case</label>
      <label class="opt"><input v-model="wholeWord" type="checkbox" /> Whole words</label>
      <button type="button" class="toggle" :aria-expanded="showReplace" title="Replace (Ctrl+H)" @click="toggleReplace">Replace…</button>
      <button type="button" class="close" title="Close (Esc)" aria-label="Close find" @click="close">✕</button>
    </div>
    <div v-if="showReplace" class="row">
      <input
        ref="replaceInput"
        v-model="replacement"
        class="field"
        type="text"
        placeholder="Replace with"
        aria-label="Replace with"
        spellcheck="false"
        @keydown.enter.prevent="replaceOne"
      />
      <button type="button" :disabled="!current" @click="replaceOne">Replace</button>
      <button type="button" :disabled="!editableMatches.length" @click="replaceAll">Replace all</button>
    </div>
  </div>
</template>

<script setup>
// Find and replace bar. Matches are highlighted with the CSS Custom Highlight API, which marks text
// without changing the DOM, so highlighting never affects undo or what is saved.
import { ref, computed, watch, nextTick, onMounted, onBeforeUnmount } from 'vue';
import { findMatches, matchRange, replaceMatch, compareToPoint } from '../editor/find.js';

const props = defineProps({
  root: { type: Object, required: true }, // the editor page element
  /** Runs a document change as one undoable edit; the change returns the text nodes it touched. */
  edit: { type: Function, required: true },
});
const emit = defineEmits(['close']);

const query = ref('');
const replacement = ref('');
const matchCase = ref(false);
const wholeWord = ref(false);
const showReplace = ref(false);
const findInput = ref(null);
const replaceInput = ref(null);
const matches = ref([]);
const currentIndex = ref(-1);
const note = ref('');
// The current match is the first one at or after this point. A live Range, so it survives the
// text-node merging that undo snapshots do.
let anchor = null;

function setAnchor(node, offset) {
  anchor = document.createRange();
  anchor.setStart(node, offset);
}

const editableMatches = computed(() => matches.value.filter((m) => m.editable));
const current = computed(() => editableMatches.value[currentIndex.value] || null);
const readOnlyCount = computed(() => matches.value.length - editableMatches.value.length);

const status = computed(() => {
  if (note.value) return note.value;
  if (!query.value) return '';
  const n = editableMatches.value.length;
  const main = n ? `${currentIndex.value + 1} of ${n}` : 'No results';
  return readOnlyCount.value ? `${main} · ${readOnlyCount.value} in read-only content` : main;
});

const highlights = typeof CSS !== 'undefined' && CSS.highlights && typeof Highlight !== 'undefined';

function paint() {
  if (!highlights) return;
  const all = new Highlight();
  for (const m of editableMatches.value) all.add(matchRange(m));
  const now = current.value ? new Highlight(matchRange(current.value)) : new Highlight();
  all.priority = 2; // above comment highlights
  now.priority = 3;
  CSS.highlights.set('find-match', all);
  CSS.highlights.set('find-current', now);
}

function clearPaint() {
  if (!highlights) return;
  CSS.highlights.delete('find-match');
  CSS.highlights.delete('find-current');
}

/** Re-runs the search, keeping the current match at the anchor point. */
function refresh() {
  matches.value = findMatches(props.root, query.value, { matchCase: matchCase.value, wholeWord: wholeWord.value });
  const list = editableMatches.value;
  if (!list.length) currentIndex.value = -1;
  else if (anchor?.startContainer.isConnected && props.root.contains(anchor.startContainer)) {
    const i = list.findIndex((m) => compareToPoint(m, anchor.startContainer, anchor.startOffset) >= 0);
    currentIndex.value = i === -1 ? 0 : i;
  } else currentIndex.value = Math.min(Math.max(currentIndex.value, 0), list.length - 1);
  paint();
}

function setAnchorToCurrent() {
  const m = current.value;
  if (!m) return;
  const r = matchRange(m);
  setAnchor(r.startContainer, r.startOffset);
}

function reveal() {
  const m = current.value;
  if (!m) return;
  const rect = matchRange(m).getBoundingClientRect();
  const canvas = props.root.closest('.canvas');
  if (!canvas) return;
  const box = canvas.getBoundingClientRect();
  if (rect.top < box.top + 8 || rect.bottom > box.bottom - 8) {
    canvas.scrollTop += rect.top - box.top - box.height / 3;
  }
}

function go(step) {
  note.value = '';
  refresh();
  const n = editableMatches.value.length;
  if (!n) return;
  currentIndex.value = (currentIndex.value + step + n) % n;
  setAnchorToCurrent();
  paint();
  reveal();
}

function replaceOne() {
  note.value = '';
  refresh();
  const m = current.value;
  if (!m) return;
  props.edit(() => {
    const after = replaceMatch(m, replacement.value);
    setAnchor(after.node, after.offset); // continue with the next match after the replaced text
    return [after.node];
  });
  refresh();
  reveal();
}

function replaceAll() {
  refresh();
  const list = editableMatches.value;
  if (!list.length) return;
  props.edit(() => {
    const touched = [];
    // Last to first, so that earlier matches' offsets stay valid.
    for (let i = list.length - 1; i >= 0; i--) touched.push(replaceMatch(list[i], replacement.value).node);
    return touched;
  });
  anchor = null;
  refresh();
  note.value = `Replaced ${list.length} ${list.length === 1 ? 'match' : 'matches'}`;
}

function toggleReplace() {
  showReplace.value = !showReplace.value;
  if (showReplace.value) nextTick(() => replaceInput.value?.focus());
}

/** Opens (or re-focuses) the bar. `from` is where to start searching; `text` pre-fills the query. */
function open({ replace = false, text = '', from = null } = {}) {
  if (text) query.value = text;
  if (replace) showReplace.value = true;
  if (from) setAnchor(from.node, from.offset);
  note.value = '';
  refresh();
  reveal();
  nextTick(() => {
    const input = replace && query.value ? replaceInput.value : findInput.value;
    input?.focus();
    input?.select();
  });
}

function close() {
  clearPaint();
  emit('close', current.value ? matchRange(current.value) : null);
}

watch([query, matchCase, wholeWord], () => {
  note.value = '';
  refresh();
  reveal();
});

// Keep matches in step with the document: typing, undo/redo, replacing.
let timer = 0;
let observer = null;
onMounted(() => {
  observer = new MutationObserver(() => {
    clearTimeout(timer);
    timer = setTimeout(refresh, 150);
  });
  observer.observe(props.root, { characterData: true, childList: true, subtree: true });
});
onBeforeUnmount(() => {
  observer?.disconnect();
  clearTimeout(timer);
  clearPaint();
});

defineExpose({ open, refresh });
</script>

<style scoped>
.findbar {
  display: flex; flex-direction: column; gap: 6px; padding: 6px 12px;
  background: var(--surface); border-bottom: 1px solid var(--border); font-size: 14px;
}
.row { display: flex; flex-wrap: wrap; align-items: center; gap: 6px; }
.field {
  flex: 0 1 260px; min-width: 140px; height: 30px; padding: 0 8px; font: inherit; color: inherit;
  background: var(--surface); border: 1px solid var(--border); border-radius: 4px;
}
.field:focus { outline: 2px solid var(--accent); outline-offset: -1px; }
.status { min-width: 80px; color: var(--muted); font-size: 13px; }
button {
  height: 30px; min-width: 30px; padding: 0 10px; font: inherit; color: inherit; cursor: pointer;
  background: transparent; border: 1px solid var(--border); border-radius: 4px;
}
button:hover:not(:disabled) { background: var(--hover); }
button:disabled { opacity: .4; cursor: default; }
.toggle[aria-expanded="true"] { background: var(--active); }
.close { border-color: transparent; margin-left: auto; }
.opt { display: inline-flex; align-items: center; gap: 4px; color: var(--muted); font-size: 13px; user-select: none; }
</style>

<style>
::highlight(find-match) { background-color: #ffe58f; color: inherit; }
::highlight(find-current) { background-color: #ff9f43; color: #000; }
</style>
