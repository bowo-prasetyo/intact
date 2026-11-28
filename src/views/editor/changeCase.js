// Change Case (Shift+F3, and its menu): sentence case, lower, UPPER, Capitalise Each Word and tOGGLE, as Word cycles them.
// Split out of EditorView.vue: see src/views/editor/README.md. The editor's shared state comes in as
// `ctx`, read when it is used, never copied at setup.
import { reactive, ref, nextTick } from 'vue';
import { t, lang } from '../../i18n.js';
import { nextCase, changeCase, caseLabel } from '../../editor/case.js';
import { portions } from '../../editor/inline.js';
import { replaceTracked, revisionDate } from '../../editor/track.js';

/**
 * `ctx`: { closestBlock, doc, explainLock, explainReadOnly, keepInView, markDirty, message, onSelectionChange, readOnlySelection, refocusAt, refreshChanges, refuseRemoving, restoreSelection, selectRange, selectionInPage, textLocked, trackedBy, tracking, undoStack }, as the editor has them.
 * Returns { caseButton, caseMenu, caseMenuEl, changeCaseTo, closeCaseMenu, toggleCaseMenu, wordAt }.
 */
export function useChangeCase(ctx) {
  const caseMenu = reactive({ open: false, x: 0, y: 0 });
  const caseMenuEl = ref(null);
  const caseButton = ref(null);
  let caseRange = null; // the text to change: the selection when the menu opened

  function toggleCaseMenu() {
    if (caseMenu.open) return closeCaseMenu();
    ctx.restoreSelection();
    const range = ctx.selectionInPage();
    if (!range) {
      ctx.message.value = 'Select the text to change the case of';
      return;
    }
    caseRange = range.cloneRange();
    const rect = caseButton.value.getBoundingClientRect();
    Object.assign(caseMenu, { open: true, x: Math.max(8, Math.min(rect.left, window.innerWidth - 220)), y: rect.bottom + 4 });
    nextTick(() => {
      ctx.keepInView(caseMenu, caseMenuEl.value);
      caseMenuEl.value?.querySelector('button')?.focus();
    });
  }

  function closeCaseMenu(refocusText = false) {
    caseMenu.open = false;
    if (refocusText) ctx.refocusAt(caseRange);
  }

  /** The word at a collapsed `range` (letters and digits around the caret), or null. */
  function wordAt(range) {
    const node = range.startContainer;
    if (node.nodeType !== 3) return null;
    const isWord = (ch) => /[\p{L}\p{N}'’]/u.test(ch || '');
    let start = range.startOffset;
    let end = range.startOffset;
    while (start > 0 && isWord(node.data[start - 1])) start--;
    while (end < node.length && isWord(node.data[end])) end++;
    if (start === end) return null;
    const word = document.createRange();
    word.setStart(node, start);
    word.setEnd(node, end);
    return word;
  }

  /**
   * Change Case, as Word's Aa menu and Shift+F3: the selected text (or the word at the caret) in
   * `kind`'s case (see editor/case.js), or, with `kind` null, the next in Shift+F3's cycle (UPPERCASE,
   * lowercase, Capitalize Each Word). Each run keeps its formatting. While tracking, each changed run
   * is a tracked deletion and insertion, as Word records a change of case. Deleted text isn't changed.
   */
  function changeCaseTo(kind) {
    if (caseMenu.open) closeCaseMenu(true);
    else ctx.restoreSelection();
    if (ctx.textLocked.value) return ctx.explainLock();
    let range = ctx.selectionInPage();
    if (!range) return;
    if (range.collapsed) {
      range = wordAt(range);
      if (!range) {
        ctx.message.value = 'Select the text to change the case of, or put the caret in a word';
        return;
      }
      ctx.selectRange(range);
    }
    const locked = ctx.readOnlySelection();
    if (locked) return ctx.explainReadOnly(locked, true);
    const pieces = portions(range).filter((p) => !p.node.parentElement.closest('[data-tc="del"]'));
    if (!pieces.length) return;
    const block = ctx.closestBlock(range.startContainer) || ctx.closestBlock(pieces[0].node);
    const before = document.createRange();
    if (block) {
      before.setStart(block, 0);
      before.setEnd(pieces[0].node, pieces[0].s);
    }
    const lang = pieces[0].node.parentElement.closest('[lang]')?.lang || ctx.doc.styles.language() || undefined;
    const texts = pieces.map((p) => p.node.data.slice(p.s, p.e));
    kind ||= nextCase(texts.join(''), lang);
    if (!kind) return;
    const changed = changeCase(texts, kind, block ? before.toString() : '', lang);
    if (changed.every((t, i) => t === texts[i])) {
      ctx.message.value = `The text is in ${caseLabel(kind)} already`;
      return;
    }
    const tracked = ctx.trackedBy(range.startContainer);
    if (tracked && ctx.refuseRemoving([range])) return;
    ctx.undoStack.flush();
    let first = null;
    let last = null;
    if (tracked) {
      // From the last piece back, so that the pieces before keep their offsets.
      const date = revisionDate();
      for (let i = pieces.length - 1; i >= 0; i--) {
        const { node, s, e } = pieces[i];
        if (changed[i] === texts[i]) {
          first = [node, s];
          last ||= [node, e];
          continue;
        }
        const r = document.createRange();
        r.setStart(node, s);
        r.setEnd(node, e);
        const after = replaceTracked(r, changed[i], { ...tracked, date });
        first = [after.node, after.offset - changed[i].length];
        last ||= [after.node, after.offset];
      }
      ctx.refreshChanges();
    } else {
      for (let i = pieces.length - 1; i >= 0; i--) {
        const { node, s, e } = pieces[i];
        node.data = node.data.slice(0, s) + changed[i] + node.data.slice(e);
        first = [node, s];
        last ||= [node, s + changed[i].length];
      }
    }
    const sel = document.createRange();
    sel.setStart(...first);
    sel.setEnd(...last);
    ctx.selectRange(sel); // still selected, as in Word: Shift+F3 again goes on to the next case
    ctx.undoStack.commit();
    ctx.markDirty();
    ctx.onSelectionChange();
    ctx.message.value = `${caseLabel(kind)}${tracked ? ' (tracked)' : ''}`;
  }

  return { caseButton, caseMenu, caseMenuEl, changeCaseTo, closeCaseMenu, toggleCaseMenu, wordAt };
}
