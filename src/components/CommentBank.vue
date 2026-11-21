<template>
  <div class="bank">
    <div class="bank-bar">
      <button type="button" class="link" :aria-expanded="String(open)" aria-controls="comment-bank" @click="toggle">
        {{ open ? '▾' : '▸' }} {{ t('Saved comments') }} <span class="count">{{ list.length }}</span>
      </button>
      <button type="button" class="link" :disabled="!current.trim()" :title="t('Keep this comment to reuse it in other comments and documents, in this browser')" @click="saveCurrent">{{ t('Save for reuse') }}</button>
    </div>
    <div v-if="open" id="comment-bank" class="bank-body">
      <input v-if="list.length > 6" v-model="filter" type="search" autocomplete="off" :aria-label="t('Find a saved comment')" :placeholder="t('Find a saved comment')" />
      <ul v-if="list.length" :aria-label="t('Saved comments')">
        <li v-for="item in shown" :key="item.text">
          <button type="button" class="use" :title="item.i < 9 ? t('Put it in the comment (Alt+{n})', { n: item.i + 1 }) : t('Put it in the comment')" @click="use(item.text)">
            <span v-if="item.i < 9" class="key" aria-hidden="true">{{ item.i + 1 }}</span>{{ item.text }}
          </button>
          <button type="button" class="remove" :aria-label="t('Remove saved comment: {text}', { text: item.text })" :title="t('Remove from saved comments')" @click="remove(item.text)">×</button>
        </li>
      </ul>
      <p v-else class="hint">{{ t('None yet: write a comment you give often, then Save for reuse. Alt+1 to Alt+9 put in the first nine.') }}</p>
      <div class="bank-file">
        <button type="button" class="link" :disabled="!list.length" @click="exportList">{{ t('Export…') }}</button>
        <button type="button" class="link" @click="fileInput.click()">{{ t('Import…') }}</button>
        <input ref="fileInput" type="file" accept=".csv,.tsv,.txt,text/csv,text/tab-separated-values,text/plain" hidden :aria-label="t('A CSV file of saved comments')" @change="importList" />
      </div>
    </div>
    <p v-if="said" class="said" role="status">{{ said }}</p>
  </div>
</template>

<script setup>
// The comment bank in the new-comment box: saved comments put into the comment being written with
// a click (or Alt+1 to Alt+9, see useNumber), the comment written saved for reuse, and the list
// exported or imported as a CSV file. Kept in this browser (see editor/commentBank.js).
import { computed, ref } from 'vue';
import { t } from '../i18n.js';
import { loadCommentBank, saveCommentBank, addToBank, bankToCsv, bankFromCsv, MAX_FILE } from '../editor/commentBank.js';
import { parseCsv, decodeCsv } from '../docx/merge.js';

const props = defineProps({ current: { type: String, default: '' } }); // the comment being written
// use(text): put a saved comment into the comment being written.
const emit = defineEmits(['use']);
const OPEN_KEY = 'intact.commentBankOpen';
const list = ref(loadCommentBank());
const open = ref((() => { try { return localStorage.getItem(OPEN_KEY) === '1'; } catch { return false; } })());
const filter = ref('');
const said = ref('');
const fileInput = ref(null);

const shown = computed(() => {
  const words = filter.value.toLowerCase().split(/\s+/).filter(Boolean);
  return list.value.map((text, i) => ({ text, i })).filter(({ text }) => words.every((w) => text.toLowerCase().includes(w)));
});

function toggle() {
  open.value = !open.value;
  try { localStorage.setItem(OPEN_KEY, open.value ? '1' : '0'); } catch { /* remembered for this page only */ }
}

function keep(next) {
  list.value = next;
  if (!saveCommentBank(next)) said.value = t('This browser keeps nothing: saved comments last until the page is closed.');
}

function saveCurrent() {
  const { list: next, added } = addToBank(list.value, [props.current]);
  if (!added) {
    said.value = t('Already saved');
    return;
  }
  keep(next);
  said.value = next.length <= 9 ? t('Saved: Alt+{n} puts it in a comment', { n: next.length }) : t('Saved for reuse');
}

function remove(text) {
  keep(list.value.filter((s) => s !== text));
  said.value = t('Removed from saved comments');
}

function use(text) {
  said.value = '';
  emit('use', text);
}

/** Alt+`n` (1 to 9): the n-th saved comment put in. False if there is none. */
function useNumber(n) {
  const text = list.value[n - 1];
  if (!text) return false;
  use(text);
  return true;
}
defineExpose({ useNumber });

/** Export…: the saved comments as a CSV file, in the downloads, on this device. */
function exportList() {
  const url = URL.createObjectURL(new Blob([bankToCsv(list.value)], { type: 'text/csv' }));
  const a = document.createElement('a');
  a.href = url;
  a.download = 'saved-comments.csv';
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10000);
  said.value = t(list.value.length === 1 ? 'Exported 1 saved comment to saved-comments.csv, in your downloads' : 'Exported {n} saved comments to saved-comments.csv, in your downloads', { n: list.value.length });
}

/** Import…: a CSV file's comments added after these (those already saved left out). */
async function importList(e) {
  const picked = e.target.files?.[0];
  e.target.value = '';
  if (!picked) return;
  if (picked.size > MAX_FILE) {
    said.value = t('{name} is too big for a list of comments (1 MB at most).', { name: picked.name });
    return;
  }
  const { list: next, added, skipped } = addToBank(list.value, bankFromCsv(decodeCsv(await picked.arrayBuffer()), parseCsv));
  keep(next);
  open.value = true;
  said.value = [
    t(added === 1 ? '1 comment imported' : '{n} comments imported', { n: added }),
    skipped && t(skipped === 1 ? '1 already saved' : '{n} already saved', { n: skipped }),
  ].filter(Boolean).join(', ');
}
</script>

<style scoped>
.bank { margin: 6px 0 0; font-size: 12px; }
.bank-bar, .bank-file { display: flex; flex-wrap: wrap; gap: 4px 12px; align-items: center; justify-content: space-between; }
.bank-file { justify-content: flex-start; margin-top: 6px; }
.link { border: 0; background: none; padding: 2px 0; color: var(--accent); font: inherit; cursor: pointer; }
.link:disabled { color: var(--muted); cursor: default; }
.count { color: var(--muted); }
.bank-body { margin-top: 4px; }
.bank-body input[type=search] { width: 100%; box-sizing: border-box; margin-bottom: 4px; padding: 4px 6px; border: 1px solid var(--border); border-radius: 4px; background: var(--surface); color: var(--text); font: inherit; }
ul { list-style: none; margin: 0; padding: 0; max-height: 180px; overflow: auto; border: 1px solid var(--border); border-radius: 4px; }
li { display: flex; align-items: flex-start; border-bottom: 1px solid var(--border); }
li:last-child { border-bottom: 0; }
.use { flex: 1; min-width: 0; padding: 5px 6px; border: 0; background: none; color: var(--text); font: inherit; text-align: left; white-space: pre-wrap; overflow-wrap: anywhere; cursor: pointer; }
.use:hover, .use:focus-visible { background: var(--hover); }
.key { display: inline-block; min-width: 14px; margin-right: 6px; padding: 0 3px; border: 1px solid var(--border); border-radius: 3px; color: var(--muted); font-size: 11px; text-align: center; }
.remove { flex: none; width: 24px; height: 24px; border: 0; background: none; color: var(--muted); font-size: 15px; cursor: pointer; }
.remove:hover { color: var(--danger); }
.hint { margin: 4px 0; color: var(--muted); }
.said { margin: 4px 0 0; color: var(--muted); }
</style>
