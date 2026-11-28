// AutoCorrect entries of your own and common typos replaced as you type, with Ctrl+Z straight after giving back what was typed, and the AutoCorrect dialog.
// Split out of EditorView.vue: see src/views/editor/README.md. The editor's shared state comes in as
// `ctx`, read when it is used, never copied at setup.
import { ref } from 'vue';
import { t, lang } from '../../i18n.js';
import { loadAutoCorrect, saveAutoCorrect, autoCorrectWord } from '../../editor/autoCorrectList.js';
import { smartQuote } from '../../editor/smartTyping.js';

/**
 * `ctx`: { autoFormat, closestBlock, doc, inPlainFieldQuiet, markDirty, message, page, pendingSplit, refocus, savedRange, selectRange, selectionInPage, trackedBy, tracking, typeText, undo, undoStack }, as the editor has them.
 * Returns { autoCorrectDialog, autoCorrectSettings, openAutoCorrect, closeAutoCorrect, saveAutoCorrectSettings, autoCorrectEnter, autoCorrectTyped }.
 */
export function useAutoCorrect(ctx) {
  // Word's AutoCorrect "Replace text as you type": your entries and common English typos, as the word
  // is ended by a space, punctuation or Enter (see editor/autoCorrectList.js). Kept in this browser.
  const autoCorrectSettings = ref(loadAutoCorrect());
  const autoCorrectDialog = ref(null); // { selected } while it is open
  const ENDS_WORD = /^[\s.,;:!?)\]}"'”’»]$/u;

  function openAutoCorrect() {
    const range = ctx.selectionInPage() || ctx.savedRange;
    const selected = range && !range.collapsed && ctx.page.value.contains(range.commonAncestorContainer) ? range.toString().replace(/\s+/g, ' ').trim().slice(0, 255) : '';
    autoCorrectDialog.value = { selected };
  }

  function closeAutoCorrect() {
    autoCorrectDialog.value = null;
    ctx.refocus();
  }

  function saveAutoCorrectSettings(settings) {
    autoCorrectSettings.value = settings;
    const kept = saveAutoCorrect(settings);
    autoCorrectDialog.value = null;
    ctx.refocus();
    const n = settings.entries.length;
    ctx.message.value = !settings.on ? t('AutoCorrect is off: text is typed as it is')
      : `${t(n === 1 ? 'AutoCorrect: 1 entry of yours' : 'AutoCorrect: {n} entries of yours', { n })}${kept ? '' : t(' – for this session only: the browser doesn’t keep them')}`;
  }

  /** The AutoCorrect replacement for the word just before the caret (`range`, collapsed), or null. */
  function correctionAt(range, typed = ' ') {
    const node = range?.startContainer;
    if (!range?.collapsed || node?.nodeType !== 3 || ctx.inPlainFieldQuiet()) return null;
    if (node.parentElement.closest('a, .raw-inline, [data-raw], .field, [data-tc="del"]')) return null;
    const block = ctx.closestBlock(node);
    if (!block) return null;
    const lead = document.createRange();
    lead.setStart(block, 0);
    lead.setEnd(node, range.startOffset);
    const tag = node.parentElement.closest('[lang]')?.lang || ctx.doc.styles.language() || (lang.value === 'id' ? 'id' : 'en');
    const hit = autoCorrectWord(lead.toString(), autoCorrectSettings.value, /^en\b/i.test(tag), typed);
    // The word typed in one piece of text, as it is.
    if (!hit || range.startOffset < hit.from.length || node.data.slice(range.startOffset - hit.from.length, range.startOffset) !== hit.from) return null;
    // What goes after it: a quote typed is a curly one where AutoFormat makes them.
    const quote = ctx.autoFormat.value && smartQuote(lead.toString(), typed, tag);
    return { ...hit, after: quote || typed };
  }

  /** Replaces the word of `hit` (see correctionAt): `after` characters typed after it stay after the caret. */
  function applyCorrection(hit, after = '') {
    const now = ctx.selectionInPage();
    const at = now?.startContainer;
    const typed = hit.from + after;
    // Found again where it is now: typing the character after it may have moved it to another node.
    if (!now?.collapsed || at?.nodeType !== 3 || at.data.slice(Math.max(0, now.startOffset - typed.length), now.startOffset) !== typed) return false;
    const word = document.createRange();
    word.setStart(at, now.startOffset - typed.length);
    word.setEnd(at, now.startOffset - after.length);
    ctx.selectRange(word);
    ctx.typeText(hit.to);
    // The caret goes back after what was typed after the word.
    const sel = window.getSelection();
    for (let i = 0; i < after.length; i++) sel.modify('move', 'forward', 'character');
    ctx.undoStack.commit();
    ctx.markDirty();
    return true;
  }

  /**
   * A space or punctuation ending a word with an entry: the character is typed first, as an undo step
   * of its own, then the word replaced, so Ctrl+Z straight after gives back what was typed, as in Word.
   */
  function autoCorrectTyped(e) {
    if (!ENDS_WORD.test(e.data)) return false;
    const hit = correctionAt(ctx.selectionInPage(), e.data);
    if (!hit) return false;
    e.preventDefault();
    ctx.typeText(hit.after);
    ctx.undoStack.commit();
    applyCorrection(hit, hit.after);
    return true;
  }

  /** Enter after a word with an entry: the word is replaced (an undo step of its own), then the paragraph is split. */
  function autoCorrectEnter(e) {
    const hit = correctionAt(ctx.selectionInPage(), '\n');
    if (!hit) return false;
    e.preventDefault();
    ctx.undoStack.flush();
    applyCorrection(hit);
    // Then Enter, as it would have been: tracked while tracking.
    const opts = ctx.trackedBy(ctx.selectionInPage()?.startContainer);
    if (opts) ctx.pendingSplit = opts;
    document.execCommand('insertParagraph'); // fires input
    return true;
  }

  return { autoCorrectDialog, autoCorrectSettings, openAutoCorrect, closeAutoCorrect, saveAutoCorrectSettings, autoCorrectEnter, autoCorrectTyped };
}
