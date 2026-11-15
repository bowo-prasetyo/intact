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
      <button type="button" title="Undo (Ctrl+Z)" @click="undo">↶</button>
      <button type="button" title="Redo (Ctrl+Y)" @click="redo">↷</button>
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
      <span class="sep" />
      <button type="button" title="Bullets" :class="{ on: state.list === 'bullet' }" :aria-pressed="state.list === 'bullet'" @click="toggleList('bullet')">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M6.5 4h8M6.5 8h8M6.5 12h8M2.5 4h.01M2.5 8h.01M2.5 12h.01" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" fill="none" /></svg>
      </button>
      <button type="button" title="Numbering" :class="{ on: state.list === 'number' }" :aria-pressed="state.list === 'number'" @click="toggleList('number')">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M6.5 4h8M6.5 8h8M6.5 12h8M2 2.5h1.2v3M1.6 5.5h2M1.6 9.6q0-1.6 1.6-1 .9.6-1.6 3h2" stroke="currentColor" stroke-width="1.1" fill="none" /></svg>
      </button>
      <button type="button" title="Decrease list level (Shift+Tab)" :disabled="!state.list" @click="changeLevel(-1)">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M7 4h7M7 8h7M7 12h7M4.5 5.5L2 8l2.5 2.5" stroke="currentColor" stroke-width="1.6" fill="none" /></svg>
      </button>
      <button type="button" title="Increase list level (Tab)" :disabled="!state.list" @click="changeLevel(1)">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M7 4h7M7 8h7M7 12h7M2 5.5L4.5 8 2 10.5" stroke="currentColor" stroke-width="1.6" fill="none" /></svg>
      </button>
      <span class="sep" />
      <button type="button" title="Find and replace (Ctrl+F, Ctrl+H)" :class="{ on: findOpen }" :aria-pressed="findOpen" @click="findOpen ? closeFind() : openFind(false)">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><circle cx="6.5" cy="6.5" r="4.2" stroke="currentColor" stroke-width="1.6" fill="none" /><path d="M9.6 9.6L14 14" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" /></svg>
      </button>
      <button v-if="threads.length" type="button" class="comments-toggle" :title="`Comments (${threads.length})`" :class="{ on: commentsOpen }" :aria-pressed="commentsOpen" @click="commentsOpen = !commentsOpen">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M2.5 3h11v7.5H7L4 13v-2.5H2.5z" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round" fill="none" /></svg>
        <span class="badge">{{ threads.length }}</span>
      </button>
    </div>

    <FindBar v-if="findOpen" ref="findBar" :root="page" :roots="searchRoots" :edit="applyEdit" @close="onFindClosed" @reveal="revealRange" />

    <div class="workspace">
      <div class="canvas" @mousedown.self="focusEnd">
        <div class="docx-sheet" :style="sheetStyle" @mousedown.self="focusEnd">
          <div v-if="doc.headers.length" class="docx-hf header" :style="headerStyle" :title="hfTitle('header')" @click="onHeaderFooterClick">
            <div v-if="doc.headers.length > 1" class="hf-tabs" role="tablist" aria-label="Headers">
              <button v-for="(v, i) in doc.headers" :key="v.rid" type="button" role="tab" :aria-selected="hf.header === i" @click="hf.header = i">{{ v.label }}</button>
            </div>
            <div ref="headerEl" class="hf-variants" />
          </div>
          <div
            ref="page"
            class="docx-page"
            contenteditable="true"
            spellcheck="true"
            :style="pageStyle"
            @beforeinput="onBeforeInput"
            @input="onInput"
            @keydown="onKeydown"
            @paste="onPaste"
            @drop="onDrop"
            @mousedown="onPageMousedown"
            @click="onClick"
          />
          <div v-if="doc.footers.length" class="docx-hf footer" :style="footerStyle" :title="hfTitle('footer')" @click="onHeaderFooterClick">
            <div ref="footerEl" class="hf-variants" />
            <div v-if="doc.footers.length > 1" class="hf-tabs" role="tablist" aria-label="Footers">
              <button v-for="(v, i) in doc.footers" :key="v.rid" type="button" role="tab" :aria-selected="hf.footer === i" @click="hf.footer = i">{{ v.label }}</button>
            </div>
          </div>
        </div>
      </div>
      <CommentsPanel v-if="threads.length" v-show="commentsOpen" :threads="threads" :active="activeComment" @select="selectComment" @close="commentsOpen = false" />
    </div>

    <div v-if="chooser.open" ref="chooserEl" class="chooser" :role="chooser.kind === 'date' ? 'dialog' : 'listbox'" aria-label="Choose a value" :style="{ left: `${chooser.x}px`, top: `${chooser.y}px` }" @keydown.esc="closeChooser(true)" @keydown="onChooserKeydown">
      <input v-if="chooser.kind === 'date'" ref="dateInput" type="date" :value="chooser.date" aria-label="Date" @keydown="onDateKeydown" @change="onDateChange" />
      <template v-else>
        <button v-for="c in chooser.choices" :key="c.value" type="button" role="option" :aria-selected="c.text === chooser.current" @click="choose(c)">{{ c.text }}</button>
        <p v-if="!chooser.choices.length" class="empty">No choices</p>
      </template>
    </div>

    <footer class="statusbar">
      <span>{{ words }} words</span>
      <span role="status" aria-live="polite">{{ message }}</span>
    </footer>
  </div>
</template>

<script setup>
import { ref, reactive, computed, watch, nextTick, onMounted, onBeforeUnmount } from 'vue';
import { useRouter, onBeforeRouteLeave } from 'vue-router';
import AppHeader from '../components/AppHeader.vue';
import FindBar from '../components/FindBar.vue';
import CommentsPanel from '../components/CommentsPanel.vue';
import { currentDoc, setDocument } from '../store.js';
import { createHistory } from '../editor/history.js';
import { isEmptyBlock, isAtBlockStart, adjacentCell, selectCell } from '../editor/structure.js';
import { controlAt, controlBox, setControlText, clearPlaceholder, fieldAtSelection } from '../editor/controls.js';
import { parseClipboard, runNodes, insertAtCaret } from '../editor/paste.js';
import {
  removesContent, rangesRemovedBy, trackedChangeAt, refusalMessage, commentMarkers, restoreCommentMarkers,
  onlyHiddenMarkers, beyondHiddenMarkers,
} from '../editor/protect.js';
import { readOnlyAt, readOnlyMessage, isEditingKey, caretBeside } from '../editor/readonly.js';

const router = useRouter();
const doc = currentDoc.value;
const page = ref(null);
const dirty = ref(false);
const saving = ref(false);
const message = ref('');
const words = ref(0);
const baseName = ref(doc.name.replace(/\.docx$/i, ''));
const state = reactive({ bold: false, italic: false, underline: false, strike: false, block: 'normal', align: 'left', list: null });
const chooser = reactive({ open: false, kind: '', x: 0, y: 0, choices: [], current: '', date: '' });
const chooserEl = ref(null);
const threads = ref([]);
const commentsOpen = ref(false);
const activeComment = ref(null);
const findOpen = ref(false);
const findBar = ref(null);
const dateInput = ref(null);
let chooserTarget = null;
let savedRange = null;
let styleEl = null;
let undoStack = null;

const BLOCK_SELECTOR = 'p,h1,h2,h3,h4,h5,h6,li';
const aligns = [
  { cmd: 'justifyLeft', value: 'left', title: 'Align left', icon: 'M2 3h12M2 6.5h8M2 10h12M2 13.5h8' },
  { cmd: 'justifyCenter', value: 'center', title: 'Center', icon: 'M2 3h12M4 6.5h8M2 10h12M4 13.5h8' },
  { cmd: 'justifyRight', value: 'right', title: 'Align right', icon: 'M2 3h12M6 6.5h8M2 10h12M6 13.5h8' },
  { cmd: 'justifyFull', value: 'justify', title: 'Justify', icon: 'M2 3h12M2 6.5h12M2 10h12M2 13.5h12' },
];

// The sheet is the page: header (in the top margin), the editable body, footer (bottom margin). As in
// Word, the body starts at the top margin or below the header, whichever is lower.
const headerEl = ref(null);
const footerEl = ref(null);
const sheetStyle = computed(() => ({ width: `${doc.page.width}px`, minHeight: `${doc.page.minHeight}px` }));
const pageStyle = computed(() => {
  const p = doc.page;
  return { padding: `${doc.headers.length ? 0 : p.top}px ${p.right}px ${doc.footers.length ? 0 : p.bottom}px ${p.left}px` };
});
const headerStyle = computed(() => {
  const p = doc.page;
  return { padding: `${p.header}px ${p.right}px 0 ${p.left}px`, minHeight: `${Math.max(0, p.top - p.header)}px` };
});
const footerStyle = computed(() => {
  const p = doc.page;
  return { padding: `0 ${p.right}px ${p.footer}px ${p.left}px`, minHeight: `${Math.max(0, p.bottom - p.footer)}px` };
});

// Which header / footer is shown, when the document has several (first page, even pages, sections).
const hf = reactive({ header: 0, footer: 0 });

function hfTitle(which) {
  const variant = (which === 'header' ? doc.headers : doc.footers)[hf[which]];
  const label = variant?.label || '';
  const name = label === 'All pages' || !label ? which : `${label} ${which}`;
  return `${name[0].toUpperCase()}${name.slice(1)} – shown for reference; edit it in Word`;
}

/** Renders every header (or footer) variant once; only the selected one is visible. */
function renderVariants(el, which) {
  if (!el) return;
  const variants = which === 'header' ? doc.headers : doc.footers;
  const nodes = variants.map((v, i) => {
    const div = document.createElement('div');
    div.className = 'hf-variant';
    div.dataset.which = which;
    div.dataset.variant = String(i);
    div.hidden = i !== hf[which];
    const content = doc.renderHeaderFooter(v);
    if (content) div.append(...content.childNodes);
    return div;
  });
  el.replaceChildren(...nodes);
  el.parentElement.classList.toggle('empty', variants.length === 1 && !nodes[0].childNodes.length);
}

watch(() => ({ ...hf }), () => {
  for (const el of [headerEl.value, footerEl.value]) {
    for (const v of el?.children || []) v.hidden = Number(v.dataset.variant) !== hf[v.dataset.which];
  }
});

function refreshCss() {
  styleEl.textContent = doc.css('.docx-sheet');
}

onMounted(() => {
  styleEl = document.createElement('style');
  document.head.append(styleEl);
  refreshCss();
  page.value.replaceChildren(doc.render());
  renderVariants(headerEl.value, 'header');
  renderVariants(footerEl.value, 'footer');
  threads.value = commentThreads();
  commentsOpen.value = threads.value.length > 0 && window.innerWidth >= 1100;
  paintComments();
  document.execCommand('defaultParagraphSeparator', false, 'p');
  document.execCommand('styleWithCSS', false, false);
  document.addEventListener('selectionchange', onSelectionChange);
  document.addEventListener('mousedown', onDocumentMousedown, true);
  window.addEventListener('keydown', onWindowKeydown, true);
  window.addEventListener('beforeunload', onBeforeUnload);
  countWords();
  page.value.focus();
  placeCaret(page.value.querySelector(BLOCK_SELECTOR) || page.value, true);
  undoStack = createHistory(page.value);
  undoStack.commit();
});

onBeforeUnmount(() => {
  document.removeEventListener('selectionchange', onSelectionChange);
  document.removeEventListener('mousedown', onDocumentMousedown, true);
  window.removeEventListener('keydown', onWindowKeydown, true);
  window.removeEventListener('beforeunload', onBeforeUnload);
  undoStack?.dispose();
  styleEl?.remove();
  if (canHighlight) {
    CSS.highlights.delete('comment');
    CSS.highlights.delete('comment-active');
    CSS.highlights.delete('field-near');
  }
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

/**
 * The paragraph a selection starts in. Firefox starts a Select All at the page itself (offset 0)
 * rather than inside the first paragraph, so look down into the node the offset points at.
 */
function startBlock(range) {
  const { startContainer: c, startOffset: o } = range;
  const at = c.nodeType === 1 ? c.childNodes[o] || c : c;
  return closestBlock(at) || (at.nodeType === 1 ? at.querySelector(BLOCK_SELECTOR) : null);
}

function closestBlock(node) {
  const el = node.nodeType === 1 ? node : node.parentElement;
  const block = el?.closest(BLOCK_SELECTOR);
  return block && page.value.contains(block) ? block : null;
}

function onSelectionChange() {
  const range = selectionInPage();
  if (!range) return;
  lastReadOnly = null;
  // Clicking into placeholder text selects all of it, so typing replaces it (as in Word).
  const control = range.collapsed ? controlAt(range.startContainer, page.value) : null;
  if (control?.classList.contains('showing-placeholder') && control.querySelector(':scope > .sdt-content')) {
    selectCell(controlBox(control));
    return;
  }
  savedRange = range.cloneRange();
  markNearField(range);
  state.bold = document.queryCommandState('bold');
  state.italic = document.queryCommandState('italic');
  state.underline = document.queryCommandState('underline');
  state.strike = document.queryCommandState('strikeThrough');
  const block = startBlock(range);
  state.block = block ? doc.blockKind(block.dataset.style) : 'normal';
  const align = block ? getComputedStyle(block).textAlign : 'left';
  state.align = { start: 'left', end: 'right' }[align] || align;
  state.list = block ? doc.listKind(doc.listOf(block)) : null;
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

const CHAR_FORMATS = new Set(['bold', 'italic', 'underline', 'strikeThrough']);

/** Plain-text fields (as in Word) take no character formatting. */
function inPlainField() {
  const range = selectionInPage();
  const box = range && controlAt(range.startContainer, page.value)?.querySelector(':scope > [data-plain]');
  if (box) message.value = 'Plain text fields can’t be formatted';
  return !!box;
}

function exec(cmd, value = null) {
  restoreSelection();
  if (CHAR_FORMATS.has(cmd) && inPlainField()) return;
  undoStack.flush();
  document.execCommand(cmd, false, value);
  undoStack.commit();
  onSelectionChange();
  markDirty();
}

function afterHistory(changed) {
  if (!changed) return;
  savedRange = null;
  refreshComments(); // undo/redo replaced the DOM the comment ranges pointed into
  onSelectionChange();
  markDirty();
}
const undo = () => { closeChooser(); afterHistory(undoStack.undo()); };
const redo = () => { closeChooser(); afterHistory(undoStack.redo()); };

function selectedBlocks(range) {
  const all = [...page.value.querySelectorAll(BLOCK_SELECTOR)].filter((b) => range.intersectsNode(b) && !b.closest('.sdt-locked'));
  if (all.length) return all;
  const one = closestBlock(range.startContainer);
  return one ? [one] : [];
}

function applyBlockStyle(kind) {
  restoreSelection();
  const range = selectionInPage();
  if (!range) return;
  undoStack.flush();
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
  doc.renumber(page.value);
  undoStack.commit();
  onSelectionChange();
  markDirty();
}

// ------------------------------------------------------------------ lists

/** Runs a change to list membership as one undoable step. */
function changeLists(change) {
  restoreSelection();
  const range = selectionInPage();
  if (!range) return;
  undoStack.flush();
  if (change(selectedBlocks(range)) === false) return;
  doc.renumber(page.value);
  undoStack.commit();
  onSelectionChange();
  markDirty();
}

/** The list next to `block` (before or after it), if it is of `kind`. */
function adjacentList(block, before, kind) {
  let el = before ? block.previousElementSibling : block.nextElementSibling;
  while (el?.classList.contains('raw-hidden')) el = before ? el.previousElementSibling : el.nextElementSibling;
  const list = el?.matches(BLOCK_SELECTOR) ? doc.listOf(el) : null;
  return list && doc.listKind(list) === kind ? list : null;
}

/**
 * Bullets / Numbering, as in Word: turns the selected paragraphs into a list of that kind (joining
 * a list of the same kind right before or after them), or takes them out of it if they already are.
 */
function toggleList(kind) {
  changeLists((blocks) => {
    if (!blocks.length) return false;
    const lists = blocks.map((b) => doc.listOf(b));
    if (lists.every((l) => doc.listKind(l) === kind)) {
      blocks.forEach((b) => doc.removeList(b));
      return true;
    }
    const join = adjacentList(blocks[0], true, kind) || adjacentList(blocks[blocks.length - 1], false, kind);
    const numId = join?.numId ?? doc.newList(kind);
    blocks.forEach((b, i) => {
      const ilvl = lists[i]?.ilvl ?? join?.ilvl ?? 0;
      doc.setList(b, numId, doc.hasListLevel(numId, ilvl) ? ilvl : 0);
    });
    return true;
  });
}

/** Moves the selected list items one level down (+1) or up (-1). */
function changeLevel(delta) {
  changeLists((blocks) => {
    const items = blocks.map((b) => [b, doc.listOf(b)]).filter(([, l]) => l && doc.hasListLevel(l.numId, l.ilvl + delta));
    if (!items.length) {
      if (delta > 0 && blocks.some((b) => doc.listOf(b))) message.value = 'This list has no further levels';
      return false;
    }
    items.forEach(([b, l]) => doc.setList(b, l.numId, l.ilvl + delta));
    return true;
  });
}

/** Enter on an empty list item moves it up a level, and out of the list from the top level. */
function endListItem(block) {
  const list = doc.listOf(block);
  changeLists(() => {
    if (list.ilvl > 0) doc.setList(block, list.numId, list.ilvl - 1);
    else doc.removeList(block);
  });
}

/** Tab / Shift+Tab: next or previous table cell, else list level, else a tab character. */
function onTab(e) {
  const range = selectionInPage();
  if (!range) return;
  const start = range.startContainer;
  const cell = (start.nodeType === 1 ? start : start.parentElement).closest('td');
  if (cell && page.value.contains(cell)) {
    const target = adjacentCell(cell, e.shiftKey ? -1 : 1);
    if (target) selectCell(target);
    else if (!e.shiftKey) message.value = 'Adding table rows isn’t supported yet';
    return;
  }
  const block = closestBlock(range.startContainer);
  const inList = range.collapsed
    ? block && doc.listOf(block) && isAtBlockStart(range, block)
    : selectedBlocks(range).some((b) => doc.listOf(b));
  if (inList) changeLevel(e.shiftKey ? -1 : 1);
  else if (!e.shiftKey) {
    const range = selectionInPage();
    if (range && !range.collapsed && refuseRemoving([range])) return;
    if (range && !range.collapsed) deleteSelection();
    document.execCommand('insertText', false, '\t');
  }
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
  restoreCommentMarkers(markersBefore, page.value);
  markersBefore = [];
  // Never leave the page completely empty: keep at least one paragraph to type into.
  if (!page.value.firstElementChild) {
    const p = document.createElement('p');
    p.append(document.createElement('br'));
    page.value.append(p);
    placeCaret(p, true);
  }
  const range = selectionInPage();
  const control = range && controlAt(range.startContainer, page.value);
  if (control) clearPlaceholder(control, (id) => doc.isPlaceholderStyle(id));
  const cell = range && (range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement)?.closest('td.showing-placeholder');
  if (cell) {
    cell.classList.remove('showing-placeholder');
    for (const s of cell.querySelectorAll('[data-rstyle]')) if (doc.isPlaceholderStyle(s.dataset.rstyle)) delete s.dataset.rstyle;
  }
  doc.renumber(page.value);
  undoStack.schedule();
  markDirty();
  refreshComments();
}

function onBeforeInput(e) {
  // Undo/redo from the browser's Edit or context menu: use our history, not the native stack.
  if (e.inputType === 'historyUndo' || e.inputType === 'historyRedo') {
    e.preventDefault();
    if (e.inputType === 'historyUndo') undo();
    else redo();
    return;
  }
  if (!removesContent(e)) return;
  // Typing, deleting or pasting while the selection is in read-only content: say what it is.
  const locked = readOnlySelection();
  if (locked) {
    e.preventDefault();
    explainReadOnly(locked, true);
    return;
  }
  const ranges = rangesRemovedBy(e);
  // Backspace or Delete next to hidden markers (comment ranges, bookmarks, field codes) would remove
  // just the invisible marker: delete what is beyond it instead.
  const sel = window.getSelection();
  const step = { deleteContentBackward: true, deleteContentForward: false }[e.inputType];
  if (step !== undefined && sel.isCollapsed && ranges.length === 1 && onlyHiddenMarkers(ranges[0])) {
    e.preventDefault();
    const target = beyondHiddenMarkers(ranges[0], step);
    if (!target || refuseRemoving([target])) return;
    markersBefore = commentMarkers(page.value);
    sel.removeAllRanges();
    sel.addRange(target);
    document.execCommand('delete'); // fires input, so onInput puts the markers back
    return;
  }
  // Typing over, deleting, cutting or dragging away a tracked change would drop it silently.
  if (refuseRemoving(ranges)) e.preventDefault();
  // Comment range ends deleted with the text are put back in onInput, at the caret: where the text
  // was deleted. (After a drag the caret is at the drop instead; the cleanup on save covers that.)
  else if (e.inputType !== 'deleteByDrag') markersBefore = commentMarkers(page.value);
}

let markersBefore = [];

/** Deletes the selection with execCommand (which fires no beforeinput), keeping comment ranges. */
function deleteSelection() {
  const before = commentMarkers(page.value);
  document.execCommand('delete');
  restoreCommentMarkers(before, page.value);
}

/**
 * Refuses an edit that would remove a tracked change, saying why: true if `ranges` (what the edit
 * would delete) touch one. Edits made with execCommand fire no beforeinput, so call this first.
 */
function refuseRemoving(ranges) {
  const el = trackedChangeAt(ranges, page.value);
  if (!el) return false;
  message.value = refusalMessage(el);
  flash(el);
  return true;
}

/** Flashes what an edit can't change. An animation, not a class, so undo snapshots don't change. */
function flash(el) {
  el.animate?.([
    { outline: '2px solid #b3261e', outlineOffset: '1px', backgroundColor: 'rgb(179 38 30 / 15%)' },
    { outline: '2px solid transparent', outlineOffset: '1px' },
  ], { duration: 1200, easing: 'ease-out' });
}

// ------------------------------------------------------------------ read-only content

// The read-only block or cell last clicked. The editor gives up focus then (the browser would make
// up a caret for the next key otherwise), and keys typed next say why nothing happens.
let lastReadOnly = null;
let pointerDown = null;

/** The read-only element the selection starts or ends in, or null. */
function readOnlySelection() {
  const range = selectionInPage();
  return range && (readOnlyAt(range.startContainer, page.value) || readOnlyAt(range.endContainer, page.value));
}

/** Says in the status bar what read-only `el` is and where to edit it; `refused` flashes it too. */
function explainReadOnly(el, refused = false) {
  message.value = readOnlyMessage(el, page.value);
  if (refused) flash(el);
}

function onPageMousedown(e) {
  pointerDown = { x: e.clientX, y: e.clientY, range: selectionInPage()?.cloneRange() || null };
}

/**
 * A click (not a drag) on read-only content says what it is. The browser leaves the caret where it
 * was, so typing would land there, out of sight: put it beside inline content, on the side clicked,
 * or clear it for a block or cell.
 */
function clickReadOnly(el, e) {
  const down = pointerDown;
  pointerDown = null;
  if (down && Math.hypot(e.clientX - down.x, e.clientY - down.y) > 4) return;
  explainReadOnly(el);
  const now = selectionInPage();
  const moved = now && down?.range && !(now.startContainer === down.range.startContainer && now.startOffset === down.range.startOffset
    && now.endContainer === down.range.endContainer && now.endOffset === down.range.endOffset);
  if (moved && !readOnlyAt(now.startContainer, page.value)) return; // the browser placed the caret itself
  const sel = window.getSelection();
  const beside = caretBeside(el, e.clientX, e.clientY);
  sel.removeAllRanges();
  lastReadOnly = beside ? null : el;
  if (!beside) {
    document.activeElement?.blur();
    return;
  }
  el.parentElement.closest('[contenteditable="true"]')?.focus({ preventScroll: true });
  sel.addRange(beside);
}

function onHeaderFooterClick(e) {
  if (!e.target.closest('button')) message.value = `Read-only: ${e.currentTarget.title}`;
}

/**
 * Typing while the selection is in a header, footer or comment (outside the editable page), or
 * after clicking a read-only block, does nothing; say why. True if the key was explained.
 */
function explainKeyOutsideText(e) {
  if (!isEditingKey(e) || document.activeElement?.matches('input, textarea, select, [contenteditable="true"]')) return false;
  if (lastReadOnly?.isConnected) {
    explainReadOnly(lastReadOnly, true);
    return true;
  }
  const node = window.getSelection().anchorNode;
  const el = node?.nodeType === 1 ? node : node?.parentElement;
  const hf = el?.closest('.docx-hf');
  if (hf) {
    message.value = `Read-only: ${hf.title}`;
    flash(hf);
    return true;
  }
  if (el?.closest('.comments') && e.key !== 'Enter') {
    message.value = 'Read-only: comments are shown for reference – reply to or resolve them in Word';
    return true;
  }
  return false;
}

function onKeydown(e) {
  const mod = e.ctrlKey || e.metaKey;
  const key = e.key.toLowerCase();
  const range = selectionInPage();
  if (isEditingKey(e) || (e.key === 'Tab' && !mod && !e.altKey)) {
    const locked = readOnlySelection();
    if (locked) {
      e.preventDefault();
      explainReadOnly(locked, true);
      return;
    }
  }
  const field = range && controlAt(range.startContainer, page.value)?.querySelector(':scope > .sdt-content');
  if ((e.altKey && !mod && e.key === 'ArrowDown') || (e.key === ' ' && !mod && !e.altKey && range && !range.collapsed)) {
    const near = fieldAtSelection(range, page.value);
    if (near) {
      e.preventDefault();
      activateField(near);
      return;
    }
  }
  if (mod && !e.altKey && ['b', 'i', 'u'].includes(key) && inPlainField()) {
    e.preventDefault();
    return;
  }
  if (e.key === 'Enter' && field?.tagName === 'SPAN') {
    // A field inside a line can't hold paragraphs; multi-line plain text fields take line breaks.
    e.preventDefault();
    if (e.shiftKey || field.dataset.plain === 'multi') document.execCommand('insertLineBreak');
    return;
  }
  if (mod && key === 's') {
    e.preventDefault();
    save();
  } else if (mod && !e.altKey && (key === 'z' || key === 'y')) {
    e.preventDefault();
    if (key === 'y' || e.shiftKey) redo();
    else undo();
  } else if (e.key === 'Tab' && !mod && !e.altKey) {
    e.preventDefault();
    onTab(e);
  } else if ((e.key === 'Enter' || e.key === 'Backspace') && !mod && !e.altKey && !e.shiftKey && !e.isComposing) {
    // Word: Enter on an empty list item, or Backspace at the start of one, ends the item's list.
    const range = selectionInPage();
    const block = range?.collapsed && closestBlock(range.startContainer);
    if (!block || !doc.listOf(block)) return;
    if (e.key === 'Enter' ? !isEmptyBlock(block) : !isAtBlockStart(range, block)) return;
    e.preventDefault();
    if (e.key === 'Enter') endListItem(block);
    else changeLists(() => doc.removeList(block));
  }
}

function onPaste(e) {
  e.preventDefault();
  const range = selectionInPage();
  const locked = readOnlySelection();
  if (locked) {
    explainReadOnly(locked, true);
    return;
  }
  if (!range || refuseRemoving([range])) return;
  undoStack.flush();
  const html = e.clipboardData.getData('text/html');
  let text = e.clipboardData.getData('text/plain').replace(/\r\n?/g, '\n');
  const field = controlAt(range.startContainer, page.value)?.querySelector(':scope > .sdt-content');
  if (field?.tagName === 'SPAN' || field?.dataset.plain) {
    // Fields inside a line, and plain-text fields, take text only; inside a line it stays on one
    // line (or becomes line breaks in multi-line fields).
    if (field.tagName === 'SPAN' && field.dataset.plain !== 'multi') text = text.replace(/\s*\n\s*/g, ' ');
    document.execCommand('insertText', false, text);
    undoStack.commit();
    onInput();
    return;
  }
  const { blocks, skippedImages } = parseClipboard(html, text);
  if (!blocks.length) return;
  if (!range.collapsed) deleteSelection();
  pasteBlocks(blocks);
  onInput();
  undoStack.commit();
  if (skippedImages) {
    message.value = `${skippedImages} ${skippedImages === 1 ? 'picture was' : 'pictures were'} not pasted – pictures can't be added here yet`;
  }
}

/**
 * Inserts parsed blocks with Word's "merge formatting": text takes the formatting where it lands plus
 * its own bold/italic/underline/strike. The first block joins the paragraph at the caret; each
 * further one starts a new paragraph that keeps the destination paragraph's settings, unless it is
 * a list item, which joins the destination list if it is of the same kind or else a new list.
 */
function pasteBlocks(blocks) {
  const startRange = selectionInPage();
  const dest = startRange && closestBlock(startRange.startContainer);
  const destWasEmpty = !!dest && isEmptyBlock(dest);
  const destNum = dest ? { numId: dest.dataset.numId, ilvl: dest.dataset.ilvl } : {};
  const destList = dest ? doc.listOf(dest) : null;
  const destKind = doc.listKind(destList);
  const newLists = {};

  const applyList = (el, block) => {
    if (block.list) {
      const { kind, level } = block.list;
      const numId = destList && destKind === kind ? destList.numId : (newLists[kind] ||= doc.newList(kind));
      doc.setList(el, numId, doc.hasListLevel(numId, level) ? level : 0);
    } else if (destNum.numId != null) {
      el.dataset.numId = destNum.numId;
      el.dataset.ilvl = destNum.ilvl ?? '0';
    } else {
      delete el.dataset.numId;
      delete el.dataset.ilvl;
    }
  };

  blocks.forEach((block, i) => {
    if (i > 0) document.execCommand('insertParagraph');
    const here = selectionInPage();
    const el = here && closestBlock(here.startContainer);
    if (el && (i > 0 || destWasEmpty)) applyList(el, block);
    insertAtCaret(runNodes(block.runs));
  });
  doc.renumber(page.value);
}

function onDrop(e) {
  if (e.dataTransfer?.files?.length) e.preventDefault();
}

function onClick(e) {
  const marker = e.target.closest('.comment-ref');
  if (marker) {
    commentsOpen.value = true;
    selectComment(marker.dataset.commentId, false);
    return;
  }
  const placeholderCell = e.target.closest('td.sdt-cell.showing-placeholder');
  if (placeholderCell) {
    selectCell(placeholderCell); // a field around a table cell: typing replaces its placeholder
    return;
  }
  const locked = readOnlyAt(e.target, page.value);
  if (locked) {
    clickReadOnly(locked, e);
    return;
  }
  const control = controlAt(e.target, page.value);
  if (control?.classList.contains('showing-placeholder') && e.target.closest('.sdt-content')) {
    selectCell(controlBox(control)); // the browser resets the selection on mouseup
    return;
  }
  if (control && !e.target.closest('.sdt-content')) {
    const kind = doc.control(control).kind;
    if (kind === 'checkbox') toggleCheckbox(control);
    else if (kind !== 'text') openChooser(control);
    return;
  }
  const a = e.target.closest('a[href]');
  if (a && (e.ctrlKey || e.metaKey) && /^(https?:|mailto:)/i.test(a.getAttribute('href'))) {
    e.preventDefault();
    window.open(a.href, '_blank', 'noopener');
  }
}

// ------------------------------------------------------------------ content controls

/** Applies a change to a control as one undoable step. */
function changeControl(el, change) {
  undoStack.flush();
  change();
  clearPlaceholder(el, (id) => doc.isPlaceholderStyle(id));
  undoStack.commit();
  markDirty();
}

function toggleCheckbox(el) {
  const { symbols } = doc.control(el);
  changeControl(el, () => {
    el.dataset.checked = el.dataset.checked === '1' ? '0' : '1';
    setControlText(el, el.dataset.checked === '1' ? symbols.on : symbols.off);
  });
}

function openChooser(el) {
  const info = doc.control(el);
  const box = el.getBoundingClientRect();
  chooserTarget = el;
  Object.assign(chooser, {
    open: true,
    kind: info.kind,
    x: Math.max(8, Math.min(box.left, window.innerWidth - 240)),
    y: box.bottom + 4,
    choices: info.choices,
    current: el.classList.contains('showing-placeholder') ? '' : controlBox(el).textContent,
    date: info.date?.iso || new Date().toISOString().slice(0, 10),
  });
  nextTick(() => (dateInput.value || chooserEl.value?.querySelector('[aria-selected="true"], button'))?.focus());
}

function closeChooser(refocus = false) {
  const el = chooserTarget;
  chooser.open = false;
  chooserTarget = null;
  if (refocus) {
    // Back to the document, with the caret just after the field.
    page.value.focus();
    if (el?.isConnected) {
      const range = document.createRange();
      range.setStartAfter(el);
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
    }
  }
}

/** ↑ / ↓ / Home / End move between the items of a drop-down list. */
function onChooserKeydown(e) {
  if (chooser.kind === 'date' || !['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(e.key)) return;
  e.preventDefault();
  const items = [...chooserEl.value.querySelectorAll('button')];
  const i = items.indexOf(document.activeElement);
  const next = { ArrowDown: Math.min(i + 1, items.length - 1), ArrowUp: Math.max(i - 1, 0), Home: 0, End: items.length - 1 }[e.key];
  items[next]?.focus();
}

// A date typed digit by digit is "complete" (and fires change) before the year is finished, so typed
// dates are taken on Enter; a date picked from the calendar is taken at once.
let lastDateKey = 0;
function onDateKeydown(e) {
  lastDateKey = Date.now();
  if (e.key === 'Enter') {
    e.preventDefault();
    chooseDate(e.target.value);
  }
}
function onDateChange(e) {
  if (Date.now() - lastDateKey > 400) chooseDate(e.target.value);
}

/** Alt+↓, or Space on a selected field: tick a check box, or open a list or date picker. */
function activateField(el) {
  if (doc.control(el).kind === 'checkbox') toggleCheckbox(el);
  else openChooser(el);
}

const canMarkFields = typeof CSS !== 'undefined' && CSS.highlights && typeof Highlight !== 'undefined';

/** Marks the field the caret is touching, so it's clear that Alt+↓ will act on it. */
function markNearField(range) {
  if (!canMarkFields) return;
  const field = fieldAtSelection(range, page.value);
  if (!field) {
    CSS.highlights.delete('field-near');
    return;
  }
  const r = document.createRange();
  r.selectNodeContents(field);
  CSS.highlights.set('field-near', new Highlight(r));
}

function onDocumentMousedown(e) {
  if (chooser.open && !chooserEl.value?.contains(e.target) && !chooserTarget?.contains(e.target)) closeChooser();
}

function choose(choice) {
  const el = chooserTarget;
  closeChooser(true);
  if (!el?.isConnected) return;
  changeControl(el, () => {
    setControlText(el, choice.text);
    el.dataset.value = choice.value;
  });
}

function chooseDate(iso) {
  const el = chooserTarget;
  if (!iso || !el?.isConnected) return;
  closeChooser(true);
  changeControl(el, () => {
    setControlText(el, doc.formatDate(el, iso));
    el.dataset.fullDate = iso;
  });
}

// ------------------------------------------------------------------ comments

/**
 * Top-level comments in the order they appear in the text, each with its replies, leaving out those
 * whose text was deleted (ids in `removed`).
 */
function commentThreads(removed = new Set()) {
  const comments = doc.comments.filter((c) => !removed.has(c.id));
  const ids = new Set(comments.map((c) => c.id));
  const order = (c) => doc.commentOrder.get(c.id) ?? Infinity;
  return comments
    .filter((c) => !c.parentId || !ids.has(c.parentId))
    .sort((a, b) => order(a) - order(b))
    .map((c) => ({ ...c, n: doc.commentOrder.get(c.id) || null, replies: comments.filter((r) => r.parentId === c.id) }));
}

// As in Word, deleting the text a comment is attached to deletes the comment. The panel follows the
// text, and the status bar says so, since Undo is the way back.
let removedComments = new Set();

function refreshComments() {
  if (!doc.comments.length) return;
  const removed = doc.removedComments(page.value);
  const added = [...removed].filter((id) => !removedComments.has(id));
  const changed = added.length || removed.size !== removedComments.size;
  removedComments = removed;
  if (changed) threads.value = commentThreads(removed);
  paintComments();
  if (added.length) message.value = commentRemovalMessage(added);
}

function commentRemovalMessage(ids) {
  const removed = doc.comments.filter((c) => ids.includes(c.id));
  const top = removed.filter((c) => !c.parentId || !ids.includes(c.parentId));
  const what = top.length === 1
    ? `the comment by ${top[0].author}${removed.length > 1 ? ` and ${removed.length - 1 === 1 ? 'its reply' : `its ${removed.length - 1} replies`}` : ''}`
    : `${removed.length} comments`;
  return `Deleted ${what} with the text ${removed.length === 1 ? 'it was' : 'they were'} attached to – Undo (Ctrl+Z) brings ${removed.length === 1 ? 'it' : 'them'} back`;
}

/** The commented text: from the range start marker to the end marker (or the reference mark). */
function commentRange(id) {
  const q = (sel) => page.value.querySelector(`[${sel}="${CSS.escape(id)}"]`);
  const start = q('data-comment-start');
  const end = q('data-comment-end');
  const ref = q('data-comment-id');
  if (!start && !ref) return null;
  const range = document.createRange();
  if (start) range.setStartAfter(start);
  else range.setStartBefore(ref);
  if (end) range.setEndBefore(end);
  else range.setEndAfter(ref);
  return range;
}

const canHighlight = typeof CSS !== 'undefined' && CSS.highlights && typeof Highlight !== 'undefined';

function paintComments() {
  if (!canHighlight) return;
  const all = new Highlight();
  const active = new Highlight();
  active.priority = 1;
  for (const t of threads.value) {
    const range = commentRange(t.id);
    if (!range) continue;
    all.add(range);
    if (t.id === activeComment.value) active.add(range);
  }
  CSS.highlights.set('comment', all);
  CSS.highlights.set('comment-active', active);
}

function selectComment(id, reveal = true) {
  activeComment.value = id;
  paintComments();
  if (reveal) {
    const range = commentRange(id);
    if (range) revealRect(range.getBoundingClientRect());
  } else {
    nextTick(() => document.querySelector('.comments .thread.active')?.scrollIntoView({ block: 'nearest' }));
  }
}

function revealRect(rect) {
  const canvas = page.value.closest('.canvas');
  const box = canvas.getBoundingClientRect();
  if (rect.top < box.top + 8 || rect.bottom > box.bottom - 8) canvas.scrollTop += rect.top - box.top - box.height / 3;
}

// ------------------------------------------------------------------ find and replace

/** Ctrl+F / Ctrl+H anywhere in the editor open Find (and Replace) instead of the browser's own. */
function onWindowKeydown(e) {
  if (explainKeyOutsideText(e)) return;
  const mod = e.ctrlKey || e.metaKey;
  const key = e.key.toLowerCase();
  if (!mod || e.altKey || e.shiftKey || (key !== 'f' && key !== 'h')) return;
  e.preventDefault();
  openFind(key === 'h');
}

async function openFind(replace) {
  const range = selectionInPage() || savedRange;
  const selected = range && !range.collapsed ? range.toString() : '';
  const text = selected && selected.length <= 200 && !/\n/.test(selected) ? selected : '';
  // Re-pressing Ctrl+F / Ctrl+H while Find is open keeps the current match.
  const from = range && !findOpen.value ? { node: range.startContainer, offset: range.startOffset } : null;
  findOpen.value = true;
  await nextTick();
  findBar.value?.open({ replace, text, from });
}

function closeFind() {
  findOpen.value = false;
  page.value.focus();
}

/** What Find searches, in order: headers, the document, footers, then comment texts. */
function searchRoots() {
  return [
    ...(headerEl.value?.children || []),
    page.value,
    ...(footerEl.value?.children || []),
    ...document.querySelectorAll('.comments .comment .text'),
  ];
}

/** Brings a match into view, switching to its header/footer tab or opening Comments if need be. */
async function revealRange(range) {
  const el = range.startContainer.parentElement;
  const variant = el?.closest('.hf-variant');
  if (variant?.hidden) hf[variant.dataset.which] = Number(variant.dataset.variant);
  if (el?.closest('.comments') && !commentsOpen.value) commentsOpen.value = true;
  await nextTick();
  const scroller = el?.closest('.canvas, .comments');
  if (!scroller) return;
  const rect = range.getBoundingClientRect();
  const box = scroller.getBoundingClientRect();
  if (rect.top < box.top + 8 || rect.bottom > box.bottom - 8) scroller.scrollTop += rect.top - box.top - box.height / 3;
}

/** Closing Find leaves the current match selected in the document, as in Word. */
function onFindClosed(range) {
  closeFind();
  if (range && page.value.contains(range.startContainer)) {
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
  }
}

/** Runs a change made outside the browser's editing (such as Replace) as one undoable edit. */
function applyEdit(change) {
  undoStack.flush();
  const touched = change() || [];
  for (const node of touched) {
    const control = controlAt(node, page.value);
    if (control) clearPlaceholder(control, (id) => doc.isPlaceholderStyle(id));
  }
  page.value.normalize();
  undoStack.commit();
  markDirty();
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
    const n = doc.commentsRemoved.size;
    message.value = `Saved ${name}${n ? ` – ${n} ${n === 1 ? 'comment' : 'comments'} removed with ${n === 1 ? 'its' : 'their'} text` : ''}`;
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
.toolbar button:hover:not(:disabled) { background: var(--hover); }
.toolbar button:disabled { opacity: .4; cursor: default; }
.toolbar button.on { background: var(--active); border-color: var(--border); }
.toolbar select { height: 32px; min-width: 130px; border: 1px solid var(--border); border-radius: 4px; background: var(--surface); color: inherit; padding: 0 6px; }
.sep { width: 1px; height: 22px; background: var(--border); margin: 0 6px; }

.workspace { flex: 1; min-height: 0; display: flex; }
.canvas { flex: 1; min-width: 0; overflow: auto; background: var(--canvas); padding: 24px 16px 48px; }
.docx-sheet {
  box-sizing: border-box; margin: 0 auto; max-width: 100%; background: #fff; color: #000;
  box-shadow: 0 1px 3px rgb(0 0 0 / 20%), 0 4px 16px rgb(0 0 0 / 8%);
  display: flex; flex-direction: column;
}
.docx-page {
  flex: 1; box-sizing: border-box; outline: none;
  white-space: pre-wrap; overflow-wrap: break-word; tab-size: 4; caret-color: #000;
}
.docx-hf { flex: none; white-space: pre-wrap; overflow-wrap: break-word; tab-size: 4; opacity: .55; cursor: default; }
.docx-hf.footer { display: flex; flex-direction: column; justify-content: flex-end; }
.docx-hf:hover { opacity: .8; }
.docx-hf.empty { padding-bottom: 0 !important; padding-top: 0 !important; }
.docx-hf { position: relative; }
.hf-tabs {
  position: absolute; left: 8px; right: 8px; display: flex; flex-wrap: wrap; gap: 4px;
  font: 11px/1.4 system-ui, sans-serif; white-space: normal;
}
.header .hf-tabs { top: 6px; }
.footer .hf-tabs { bottom: 6px; }
.hf-tabs button {
  padding: 1px 8px; border: 1px solid #c9d1de; border-radius: 9px; background: #fff; color: #44516a; cursor: pointer;
}
.hf-tabs button[aria-selected="true"] { background: #2b579a; border-color: #2b579a; color: #fff; }
.comments-toggle { gap: 3px; padding: 0 6px; }
.badge { font-size: 11px; color: var(--muted); }
.chooser {
  position: fixed; z-index: 10; min-width: 160px; max-width: 280px; max-height: 280px; overflow: auto;
  display: flex; flex-direction: column; padding: 4px; background: var(--surface); color: inherit;
  border: 1px solid var(--border); border-radius: 6px; box-shadow: 0 4px 16px rgb(0 0 0 / 15%);
}
.chooser button { text-align: left; padding: 6px 10px; border: 0; border-radius: 4px; background: transparent; color: inherit; font: inherit; cursor: pointer; }
.chooser button:hover, .chooser button:focus-visible, .chooser button[aria-selected="true"] { background: var(--hover); }
.chooser input { font: inherit; padding: 4px; }
.chooser .empty { margin: 6px 10px; color: var(--muted); }
.statusbar {
  display: flex; gap: 24px; padding: 4px 16px; font-size: 12px; color: var(--muted);
  background: var(--surface); border-top: 1px solid var(--border);
}

@media (max-width: 860px) {
  .canvas { padding: 0; }
  .docx-sheet { min-height: 100% !important; width: 100% !important; box-shadow: none; }
  .docx-page { padding: 24px 16px !important; }
  .docx-hf { padding: 12px 16px !important; min-height: 0 !important; }
  .docx-hf.header:has(.hf-tabs) { padding-top: 34px !important; }
  .docx-hf.footer:has(.hf-tabs) { padding-bottom: 34px !important; }
}
</style>

<style>
/* Unscoped: applies to the generated document HTML. */
.docx-sheet :is(h1, h2, h3, h4, h5, h6) { font-size: inherit; font-weight: inherit; }
.docx-sheet :is(p, h1, h2, h3, h4, h5, h6) { position: relative; }
.docx-sheet .cs-before { margin-top: 0 !important; }
.docx-sheet .cs-after { margin-bottom: 0 !important; }
.docx-sheet .list::before {
  content: attr(data-marker); position: absolute; transform: translateX(calc(-100% - 0.5em));
  text-indent: 0; white-space: nowrap;
}
.docx-sheet a { color: #0563c1; }
.docx-sheet img { max-width: 100%; vertical-align: bottom; }
/* Hidden markers (comment ranges, bookmarks, field codes) are empty: shown inline, not display:none,
   because the browser's editing drops unrendered nodes when it moves text, as in merging paragraphs. */
.docx-sheet .raw-hidden { display: none; }
.docx-sheet span.raw-hidden { display: inline; }
.docx-sheet .raw-block, .docx-sheet .raw-inline:not(:has(img)) {
  background: #f1f3f6; color: #555; border-radius: 3px; cursor: default; user-select: none;
}
.docx-sheet .raw-block { padding: 4px 8px; margin: 4px 0; border-left: 3px solid #c3c9d4; white-space: pre-wrap; }
.docx-sheet .raw-inline.deleted { text-decoration: line-through; color: #b3261e; background: transparent; }
.docx-sheet .raw-inline.inserted { text-decoration: underline; color: #1e6b34; background: transparent; }
.docx-sheet .raw-inline.page-break { display: block; height: 0; border-top: 1px dashed #aab; margin: 8px 0; background: none; }
.docx-sheet .raw-inline.note-ref { vertical-align: super; font-size: smaller; background: transparent; color: #2b579a; }
.docx-sheet .raw-inline.note-ref::before { content: attr(data-n); }
.docx-sheet .raw-inline.symbol { background: transparent; color: inherit; }
.docx-sheet .raw-inline.comment-ref { display: inline; cursor: pointer; user-select: none; background: transparent; padding: 0; }
.docx-sheet .comment-ref::before {
  content: attr(data-n); display: inline-block; min-width: 1.3em; margin: 0 1px; padding: 0 3px; border-radius: 7px;
  font: 600 10px/1.4 system-ui, sans-serif; text-align: center; vertical-align: super; background: #d0705c; color: #fff;
}
.docx-sheet .field { border-bottom: 1px dotted #9aa6b8; }
::highlight(comment) { background-color: #fbe3dc; }
::highlight(field-near) { background-color: #d6e2f5; }
::highlight(comment-active) { background-color: #f6b9a8; }
.docx-sheet .sdt { border-radius: 2px; transition: background-color .1s, box-shadow .1s; }
.docx-sheet div.sdt { box-shadow: 0 0 0 1px transparent; }
.docx-sheet .sdt:not(.sdt-locked):hover, .docx-sheet .sdt:not(.sdt-locked):focus-within { background: #e8eef8; box-shadow: 0 0 0 1px #9db3d9; }
.docx-sheet .sdt-content { outline: none; }
.docx-sheet span.sdt-content:empty { padding: 0 1ch; }
.docx-sheet .showing-placeholder { color: #767676; }
.docx-sheet :is(.sdt-checkbox, .sdt-dropdown, .sdt-date) { cursor: pointer; }
.docx-sheet :is(span.sdt-dropdown, span.sdt-combo, span.sdt-date)::after { content: '▾'; font-size: .8em; color: #5b6b85; margin-left: 2px; cursor: pointer; }
.docx-sheet .sdt-locked { cursor: default; user-select: none; background: #f1f3f6; }
.docx-sheet table.docx-table { border-collapse: collapse; width: 100%; margin: 4px 0; }
.docx-sheet table.docx-table td { border: 1px solid #bfbfbf; padding: 2px 6px; vertical-align: top; outline: none; }
.docx-sheet table.docx-table td:focus-within { box-shadow: inset 0 0 0 2px #2b579a55; }
.docx-sheet table.docx-table td.vmerge { border-top-color: transparent; }
.docx-sheet table.docx-table td.vmerge > * { visibility: hidden; }
</style>
