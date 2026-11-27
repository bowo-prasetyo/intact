// Redact: the selection blacked out for good, never tracked, and where a redacted text is still in
// the file, for Before you release.
// Split out of EditorView.vue: see src/views/editor/README.md. The editor's shared state comes in as
// `ctx`, read when it is used, never copied at setup.
import { t } from '../../i18n.js';
import { findMatches, matchRange } from '../../editor/find.js';
import { portions } from '../../editor/inline.js';
import { readOnlyAt } from '../../editor/readonly.js';

/**
 * `ctx`: { changes, countWords, doc, explainLock, explainReadOnly, flash, markDirty, message, page, restoreSelection, savedRange, selectionInPage, textLocked, threads, tracking, undoStack }, as the editor has them.
 * Returns { redactSelection, redactLeftovers, redactedLeftovers, redactedPassages }.
 */
export function useRedaction(ctx) {
  // Redact: the selected text replaced with █, one for each character, as a page is blacked out before
  // it is released (a source's name, an address in a leaked document). Never tracked, even while
  // tracking: a tracked deletion keeps the text in the file, for anyone to read in Word. A link in it
  // goes, address and all (see doc.forgetLink). What was redacted is remembered in this session only
  // (never in a draft or the file), so "Before you release" can say where it still is.
  const REDACTION = '█';
  const redactedTexts = new Set();
  const DELETIONS = '[data-tc="del"], .raw-inline.deleted, [data-tracked]';

  /** A tracked change in `range` (anyone's), whose earlier wording would stay in the file, or null. */
  function trackedIn(range) {
    return [...ctx.page.value.querySelectorAll(DELETIONS)].find((el) => range.intersectsNode(el)) || null;
  }

  /** Read-only content with text inside `range` (a field, an equation, kept content), which can't be redacted here. */
  function lockedTextIn(range) {
    if (readOnlyAt(range.startContainer, ctx.page.value)) return readOnlyAt(range.startContainer, ctx.page.value);
    if (readOnlyAt(range.endContainer, ctx.page.value)) return readOnlyAt(range.endContainer, ctx.page.value);
    return [...ctx.page.value.querySelectorAll('[contenteditable="false"]')].find((el) => range.intersectsNode(el)
      && !el.contains(range.commonAncestorContainer) && !el.matches('.note-ref, .comment-ref, .picture, .picture *')
      && el.textContent.trim()) || null;
  }

  /** Blacks out the editable text of `range` (see portions); returns how many characters. */
  function redactRange(range) {
    const pieces = portions(range);
    const links = new Set();
    let n = 0;
    for (let i = pieces.length - 1; i >= 0; i--) {
      const { node, s, e } = pieces[i];
      node.data = node.data.slice(0, s) + node.data.slice(s, e).replace(/[^\t\n\r]/gu, REDACTION) + node.data.slice(e);
      n += e - s;
      const a = node.parentElement.closest('a[href]');
      if (a && ctx.page.value.contains(a)) links.add(a);
    }
    for (const a of links) {
      ctx.doc.forgetLink(a);
      a.replaceWith(...a.childNodes);
    }
    return n;
  }

  /** Redact (Review tab, right-click menu): blacks out the selected text for good. */
  function redactSelection() {
    if (ctx.textLocked.value) return ctx.explainLock();
    ctx.restoreSelection();
    const range = ctx.selectionInPage() || ctx.savedRange;
    if (!range || range.collapsed || !ctx.page.value.contains(range.commonAncestorContainer)) {
      ctx.message.value = t('Select the text to redact first, such as a source’s name');
      ctx.page.value.focus();
      return;
    }
    const tracked = trackedIn(range);
    if (tracked) {
      ctx.message.value = t('The selection holds tracked changes: accept or reject them first (Review › Accept), so no earlier wording stays in the file');
      ctx.flash(tracked);
      return;
    }
    const locked = lockedTextIn(range);
    if (locked) return ctx.explainReadOnly(locked, true);
    const text = range.toString().replace(/\s+/g, ' ').trim();
    const pictures = [...ctx.page.value.querySelectorAll('.picture')].some((el) => range.intersectsNode(el));
    ctx.undoStack.flush();
    const end = [range.endContainer, range.endOffset];
    const n = redactRange(range);
    if (!n) return;
    ctx.page.value.normalize();
    ctx.undoStack.commit();
    ctx.markDirty();
    if (text.replaceAll(REDACTION, '').trim().length >= 2) redactedTexts.add(text);
    try { getSelection().collapse(...end); } catch { /* the end moved: the caret stays where it is */ }
    const left = redactedLeftovers().find((l) => l.text === text);
    const more = left ? left.redactable + left.tracked + left.locked : 0;
    ctx.message.value = [
      t('Redacted{tracked} – Undo (Ctrl+Z) brings it back', { tracked: ctx.tracking.value ? t(' (not tracked: a tracked deletion would keep the text in the file)') : '' }),
      pictures && t('pictures in it stay: delete them'),
      more && t('“{text}” is in the document {n} more times: Before you release (Review) redacts them all', { text: text.length > 40 ? `${text.slice(0, 40)}…` : text, n: more }),
    ].filter(Boolean).join(' · ');
  }

  /**
   * Where what was redacted this session still is: [{ text, redactable, tracked, locked, comments,
   * properties }], counts of the places in the text (and its headers, footers and notes) that can be
   * redacted here, in tracked changes, in read-only content; comments and properties that hold it.
   */
  function redactedLeftovers() {
    const comments = ctx.threads.value.flatMap((th) => [th, ...th.replies]).map((c) => (c.paragraphs || []).join('\n'));
    const p = ctx.doc.properties || {};
    const values = [...(p.core || []), ...(p.app || []), ...(p.custom || [])].map((x) => String(x.value ?? ''));
    const out = [];
    for (const text of redactedTexts) {
      const where = { text, redactable: 0, tracked: 0, locked: 0, comments: 0, properties: 0 };
      for (const m of findMatches(ctx.page.value, text)) {
        const el = m.segment.nodes[0].node.parentElement;
        if (el.closest(DELETIONS)) where.tracked++;
        else if (m.editable && !readOnlyAt(el, ctx.page.value)) where.redactable++;
        else where.locked++;
      }
      const lower = text.toLowerCase();
      where.comments = comments.filter((c) => c.toLowerCase().includes(lower)).length;
      where.properties = values.filter((v) => v.toLowerCase().includes(lower)).length;
      if (where.redactable || where.tracked || where.locked || where.comments || where.properties) out.push(where);
    }
    return out;
  }

  /** Before you release › Redact them all: every place a redacted text is still in editable text. */
  function redactLeftovers() {
    ctx.undoStack.flush();
    let n = 0;
    for (const text of redactedTexts) {
      const ranges = findMatches(ctx.page.value, text)
        .filter((m) => m.editable && !m.segment.nodes[0].node.parentElement.closest(DELETIONS) && !readOnlyAt(m.segment.nodes[0].node.parentElement, ctx.page.value))
        .map(matchRange);
      // From the last back, so that the earlier ones keep their offsets.
      for (const range of ranges.reverse()) n += redactRange(range) ? 1 : 0;
    }
    ctx.page.value.normalize();
    ctx.undoStack.commit();
    ctx.markDirty();
    ctx.countWords();
    ctx.message.value = t('Redacted {n} more places{tracked} – Undo (Ctrl+Z) brings them back', { n, tracked: ctx.tracking.value ? t(' (not tracked)') : '' });
  }

  /** Redacted passages in the page: runs of █ (also from an earlier session, or another tool). */
  function redactedPassages() {
    return (ctx.page.value.textContent.match(/█+/g) || []).length;
  }

  return { redactSelection, redactLeftovers, redactedLeftovers, redactedPassages };
}
