<template>
  <div class="editor-shell">
    <AppHeader>
      <div class="file">
        <input v-model="baseName" class="filename" aria-label="File name" spellcheck="false" :style="{ width: `${Math.max(baseName.length, 4) + 2}ch` }" />
        <span class="ext">.docx</span>
        <span v-if="dirty" class="dirty" title="Unsaved changes">●</span>
      </div>
      <button class="primary" type="button" :disabled="saving" title="Download (Ctrl+S)" @click="save">
        {{ saving ? 'Saving…' : 'Download' }}
      </button>
      <button class="ghost" type="button" @click="close">Close</button>
    </AppHeader>

    <div class="toolbar" role="toolbar" aria-label="Formatting" @mousedown="keepFocus">
      <button type="button" title="Undo (Ctrl+Z)" @click="exec('undo')">↶</button>
      <button type="button" title="Redo (Ctrl+Y)" @click="exec('redo')">↷</button>
      <span class="sep" />
      <select :value="state.block" aria-label="Paragraph style" @change="applyBlockStyle($event.target.value)">
        <option value="normal">Normal</option>
        <option value="title">Title</option>
        <option value="h1">Heading 1</option>
        <option value="h2">Heading 2</option>
        <option value="h3">Heading 3</option>
        <option v-if="!['normal', 'title', 'h1', 'h2', 'h3'].includes(state.block)" :value="state.block" disabled>
          {{ state.block === 'other' ? 'Other style' : state.block.replace('h', 'Heading ') }}
        </option>
      </select>
      <span class="sep" />
      <button type="button" title="Bold (Ctrl+B)" :class="{ on: state.bold }" @click="exec('bold')"><b>B</b></button>
      <button type="button" title="Italic (Ctrl+I)" :class="{ on: state.italic }" @click="exec('italic')"><i>I</i></button>
      <button type="button" title="Underline (Ctrl+U)" :class="{ on: state.underline }" @click="exec('underline')"><u>U</u></button>
      <button type="button" title="Strikethrough" :class="{ on: state.strike }" @click="exec('strikeThrough')"><s>S</s></button>
      <span class="sep" />
      <button v-for="a in aligns" :key="a.cmd" type="button" :title="a.title" :class="{ on: state.align === a.value }" @click="exec(a.cmd)">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path :d="a.icon" stroke="currentColor" stroke-width="1.6" fill="none" /></svg>
      </button>
    </div>

    <div class="canvas" @mousedown.self="focusEnd">
      <div
        ref="page"
        class="docx-page"
        contenteditable="true"
        spellcheck="true"
        :style="pageStyle"
        @input="onInput"
        @keydown="onKeydown"
        @paste="onPaste"
        @drop="onDrop"
        @click="onClick"
      />
    </div>

    <footer class="statusbar">
      <span>{{ words }} words</span>
      <span v-if="message">{{ message }}</span>
    </footer>
  </div>
</template>

<script setup>
import { ref, reactive, computed, onMounted, onBeforeUnmount } from 'vue';
import { useRouter, onBeforeRouteLeave } from 'vue-router';
import AppHeader from '../components/AppHeader.vue';
import { currentDoc, setDocument } from '../store.js';

const router = useRouter();
const doc = currentDoc.value;
const page = ref(null);
const dirty = ref(false);
const saving = ref(false);
const message = ref('');
const words = ref(0);
const baseName = ref(doc.name.replace(/\.docx$/i, ''));
const state = reactive({ bold: false, italic: false, underline: false, strike: false, block: 'normal', align: 'left' });
let savedRange = null;
let styleEl = null;

const BLOCK_SELECTOR = 'p,h1,h2,h3,h4,h5,h6,li';
const aligns = [
  { cmd: 'justifyLeft', value: 'left', title: 'Align left', icon: 'M2 3h12M2 6.5h8M2 10h12M2 13.5h8' },
  { cmd: 'justifyCenter', value: 'center', title: 'Center', icon: 'M2 3h12M4 6.5h8M2 10h12M4 13.5h8' },
  { cmd: 'justifyRight', value: 'right', title: 'Align right', icon: 'M2 3h12M6 6.5h8M2 10h12M6 13.5h8' },
  { cmd: 'justifyFull', value: 'justify', title: 'Justify', icon: 'M2 3h12M2 6.5h12M2 10h12M2 13.5h12' },
];

const pageStyle = computed(() => {
  const p = doc.page;
  return {
    width: `${p.width}px`,
    minHeight: `${p.minHeight}px`,
    padding: `${p.top}px ${p.right}px ${p.bottom}px ${p.left}px`,
  };
});

function refreshCss() {
  styleEl.textContent = doc.css('.docx-page');
}

onMounted(() => {
  styleEl = document.createElement('style');
  document.head.append(styleEl);
  refreshCss();
  page.value.replaceChildren(doc.render());
  document.execCommand('defaultParagraphSeparator', false, 'p');
  document.execCommand('styleWithCSS', false, false);
  document.addEventListener('selectionchange', onSelectionChange);
  window.addEventListener('beforeunload', onBeforeUnload);
  countWords();
  page.value.focus();
  placeCaret(page.value.querySelector(BLOCK_SELECTOR) || page.value, true);
});

onBeforeUnmount(() => {
  document.removeEventListener('selectionchange', onSelectionChange);
  window.removeEventListener('beforeunload', onBeforeUnload);
  styleEl?.remove();
});

onBeforeRouteLeave(() => {
  if (dirty.value && !window.confirm('You have unsaved changes. Leave without downloading?')) return false;
  return true;
});

function onBeforeUnload(e) {
  if (dirty.value) e.preventDefault();
}

// ------------------------------------------------------------------ selection & toolbar state

function selectionInPage() {
  const sel = window.getSelection();
  if (!sel.rangeCount) return null;
  const range = sel.getRangeAt(0);
  return page.value?.contains(range.commonAncestorContainer) ? range : null;
}

function closestBlock(node) {
  const el = node.nodeType === 1 ? node : node.parentElement;
  const block = el?.closest(BLOCK_SELECTOR);
  return block && page.value.contains(block) ? block : null;
}

function onSelectionChange() {
  const range = selectionInPage();
  if (!range) return;
  savedRange = range.cloneRange();
  state.bold = document.queryCommandState('bold');
  state.italic = document.queryCommandState('italic');
  state.underline = document.queryCommandState('underline');
  state.strike = document.queryCommandState('strikeThrough');
  const block = closestBlock(range.startContainer);
  state.block = block ? doc.blockKind(block.dataset.style) : 'normal';
  const align = block ? getComputedStyle(block).textAlign : 'left';
  state.align = { start: 'left', end: 'right' }[align] || align;
}

function restoreSelection() {
  if (selectionInPage()) return;
  page.value.focus();
  if (savedRange) {
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(savedRange);
  }
}

function keepFocus(e) {
  // Keep the caret in the document when clicking toolbar buttons (but let <select> open).
  if (e.target.closest('button')) e.preventDefault();
}

function placeCaret(el, atStart = false) {
  const range = document.createRange();
  range.selectNodeContents(el);
  range.collapse(atStart);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
}

function focusEnd() {
  page.value.focus();
  const blocks = page.value.querySelectorAll(':scope > ' + BLOCK_SELECTOR.split(',').join(', :scope > '));
  placeCaret(blocks[blocks.length - 1] || page.value);
}

// ------------------------------------------------------------------ editing commands

function exec(cmd, value = null) {
  restoreSelection();
  document.execCommand(cmd, false, value);
  onSelectionChange();
  markDirty();
}

function selectedBlocks(range) {
  const all = [...page.value.querySelectorAll(BLOCK_SELECTOR)].filter((b) => range.intersectsNode(b));
  if (all.length) return all;
  const one = closestBlock(range.startContainer);
  return one ? [one] : [];
}

function applyBlockStyle(kind) {
  restoreSelection();
  const range = selectionInPage();
  if (!range) return;
  const styleId = doc.styleIdFor(kind);
  refreshCss(); // a missing builtin style may just have been added
  const tag = kind === 'normal' ? 'p' : kind === 'title' ? 'h1' : kind;
  const { startContainer, startOffset, endContainer, endOffset } = range;
  const replaced = new Map();
  for (const block of selectedBlocks(range)) {
    let el = block;
    if (block.tagName.toLowerCase() !== tag) {
      el = document.createElement(tag);
      for (const attr of block.attributes) el.setAttribute(attr.name, attr.value);
      while (block.firstChild) el.append(block.firstChild);
      block.replaceWith(el);
      replaced.set(block, el);
    }
    if (styleId) el.dataset.style = styleId;
    else delete el.dataset.style;
  }
  const r = document.createRange();
  r.setStart(replaced.get(startContainer) || startContainer, startOffset);
  r.setEnd(replaced.get(endContainer) || endContainer, endOffset);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(r);
  onSelectionChange();
  markDirty();
}

// ------------------------------------------------------------------ events

let wordTimer = 0;
function countWords() {
  const text = page.value?.innerText || '';
  words.value = (text.match(/[\p{L}\p{N}]+(?:['’-][\p{L}\p{N}]+)*/gu) || []).length;
}

function markDirty() {
  dirty.value = true;
  message.value = '';
  clearTimeout(wordTimer);
  wordTimer = setTimeout(countWords, 300);
}

function onInput() {
  // Never leave the page completely empty: keep at least one paragraph to type into.
  if (!page.value.firstElementChild) {
    const p = document.createElement('p');
    p.append(document.createElement('br'));
    page.value.append(p);
    placeCaret(p, true);
  }
  markDirty();
}

function onKeydown(e) {
  const mod = e.ctrlKey || e.metaKey;
  if (mod && e.key.toLowerCase() === 's') {
    e.preventDefault();
    save();
  } else if (e.key === 'Tab' && !mod && !e.altKey) {
    e.preventDefault();
    document.execCommand('insertText', false, '\t');
  }
}

function onPaste(e) {
  // Paste as plain text: foreign HTML can't be mapped faithfully to Word markup.
  e.preventDefault();
  const text = e.clipboardData.getData('text/plain').replace(/\r\n?/g, '\n');
  text.split('\n').forEach((line, i) => {
    if (i) document.execCommand('insertParagraph');
    if (line) document.execCommand('insertText', false, line);
  });
}

function onDrop(e) {
  if (e.dataTransfer?.files?.length) e.preventDefault();
}

function onClick(e) {
  const a = e.target.closest('a[href]');
  if (a && (e.ctrlKey || e.metaKey) && /^(https?:|mailto:)/i.test(a.getAttribute('href'))) {
    e.preventDefault();
    window.open(a.href, '_blank', 'noopener');
  }
}

// ------------------------------------------------------------------ file actions

async function save() {
  if (saving.value) return;
  saving.value = true;
  try {
    const name = `${baseName.value.trim() || 'document'}.docx`;
    const blob = await doc.save(page.value);
    doc.name = name;
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    document.body.append(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
    dirty.value = false;
    message.value = `Saved ${name}`;
  } catch (e) {
    console.error(e);
    message.value = `Save failed: ${e.message}`;
  } finally {
    saving.value = false;
  }
}

async function close() {
  const ok = await router.push({ name: 'home' });
  if (ok === undefined) setDocument(null); // navigation succeeded
}
</script>

<style scoped>
.editor-shell { height: 100vh; height: 100dvh; display: flex; flex-direction: column; }
.file { display: flex; align-items: center; gap: 2px; min-width: 0; flex: 1; }
.filename {
  min-width: 0; max-width: 100%; font: inherit; color: #fff; background: transparent;
  border: 1px solid transparent; border-radius: 4px; padding: 4px 6px;
}
.filename:hover, .filename:focus { border-color: rgb(255 255 255 / 50%); outline: none; background: rgb(255 255 255 / 10%); }
.ext { opacity: .7; }
.dirty { color: #ffd166; margin-left: 6px; }
.primary { background: #fff; color: var(--brand); border: 0; border-radius: 4px; padding: 6px 14px; font-weight: 600; cursor: pointer; }
.primary:disabled { opacity: .6; }
@media (max-width: 560px) { .ext, :deep(.nav) { display: none; } }
.ghost { background: transparent; color: #fff; border: 1px solid rgb(255 255 255 / 50%); border-radius: 4px; padding: 5px 12px; cursor: pointer; }

.toolbar {
  display: flex; flex-wrap: wrap; align-items: center; gap: 2px; padding: 6px 12px;
  background: var(--surface); border-bottom: 1px solid var(--border);
}
.toolbar button {
  min-width: 32px; height: 32px; display: inline-flex; align-items: center; justify-content: center;
  border: 1px solid transparent; border-radius: 4px; background: transparent; color: inherit; font-size: 15px; cursor: pointer;
}
.toolbar button:hover { background: var(--hover); }
.toolbar button.on { background: var(--active); border-color: var(--border); }
.toolbar select { height: 32px; min-width: 130px; border: 1px solid var(--border); border-radius: 4px; background: var(--surface); color: inherit; padding: 0 6px; }
.sep { width: 1px; height: 22px; background: var(--border); margin: 0 6px; }

.canvas { flex: 1; overflow: auto; background: var(--canvas); padding: 24px 16px 48px; }
.docx-page {
  box-sizing: border-box; margin: 0 auto; max-width: 100%; background: #fff; color: #000;
  box-shadow: 0 1px 3px rgb(0 0 0 / 20%), 0 4px 16px rgb(0 0 0 / 8%); outline: none;
  white-space: pre-wrap; overflow-wrap: break-word; tab-size: 4; caret-color: #000;
}
.statusbar {
  display: flex; gap: 24px; padding: 4px 16px; font-size: 12px; color: var(--muted);
  background: var(--surface); border-top: 1px solid var(--border);
}

@media (max-width: 860px) {
  .canvas { padding: 0; }
  .docx-page { padding: 24px 16px !important; min-height: 100% !important; width: 100% !important; box-shadow: none; }
}
</style>

<style>
/* Unscoped: applies to the generated document HTML. */
.docx-page :is(h1, h2, h3, h4, h5, h6) { font-size: inherit; font-weight: inherit; }
.docx-page :is(p, h1, h2, h3, h4, h5, h6) { position: relative; }
.docx-page .list::before {
  content: attr(data-marker); position: absolute; transform: translateX(calc(-100% - 0.5em));
  text-indent: 0; white-space: nowrap;
}
.docx-page a { color: #0563c1; }
.docx-page img { max-width: 100%; vertical-align: bottom; }
.docx-page .raw-hidden { display: none; }
.docx-page .raw-block, .docx-page .raw-inline:not(:has(img)) {
  background: #f1f3f6; color: #555; border-radius: 3px; cursor: default; user-select: none;
}
.docx-page .raw-block { padding: 4px 8px; margin: 4px 0; border-left: 3px solid #c3c9d4; white-space: pre-wrap; }
.docx-page .raw-inline.deleted { text-decoration: line-through; color: #b3261e; background: transparent; }
.docx-page .raw-inline.page-break { display: block; height: 0; border-top: 1px dashed #aab; margin: 8px 0; background: none; }
.docx-page .raw-inline.note-ref { vertical-align: super; font-size: smaller; }
.docx-page table.docx-table { border-collapse: collapse; width: 100%; margin: 4px 0; }
.docx-page table.docx-table td { border: 1px solid #bfbfbf; padding: 2px 6px; vertical-align: top; outline: none; }
.docx-page table.docx-table td:focus-within { box-shadow: inset 0 0 0 2px #2b579a55; }
.docx-page table.docx-table td.vmerge { border-top-color: transparent; }
.docx-page table.docx-table td.vmerge > * { visibility: hidden; }
</style>
