// Keeps deletions from quietly taking what the user can't see or edit with them.
//  * Notes (data-protected): the footnotes and endnotes below the text, each note and its number,
//    can't be deleted with text around them; a note goes when its reference in the text does.
//  * Tracked changes: deleting a revision would accept or reject it without a trace, so an edit
//    that would remove anything labelled data-tracked (see DocxDocument.markTracked) is refused.
//  * Comments: a comment lives as long as its mark (data-comment-ref). When a deletion takes one end
//    of a surviving comment's range with the surrounding text, that marker is put back where the
//    text was deleted, as Word shrinks the range rather than losing it.
//  * Hidden markers (comment ranges, bookmarks, field codes): Backspace or Delete next to one deletes
//    the character beyond it, not the invisible marker.

// How far a collapsed delete reaches, for browsers that report no target range.
const STEPS = {
  deleteContentBackward: ['backward', 'character'],
  deleteContentForward: ['forward', 'character'],
  deleteWordBackward: ['backward', 'word'],
  deleteWordForward: ['forward', 'word'],
  deleteSoftLineBackward: ['backward', 'lineboundary'],
  deleteSoftLineForward: ['forward', 'lineboundary'],
  deleteHardLineBackward: ['backward', 'paragraphboundary'],
  deleteHardLineForward: ['forward', 'paragraphboundary'],
};

function toRange(r) {
  if (r instanceof Range) return r;
  try {
    const range = document.createRange();
    range.setStart(r.startContainer, r.startOffset);
    range.setEnd(r.endContainer, r.endOffset);
    return range;
  } catch {
    return null; // a static range whose nodes have moved
  }
}

/** True if a beforeinput event of this type removes the content of its target ranges. */
export function removesContent(e) {
  return e.inputType.startsWith('delete') || e.inputType.startsWith('insert');
}

/**
 * The ranges a beforeinput event would remove: its target ranges, or else the selection (extended
 * by the step a collapsed delete takes, as the browser would). Collapsed ranges remove nothing.
 */
export function rangesRemovedBy(e) {
  const targets = (e.getTargetRanges?.() || []).map(toRange).filter((r) => r && !r.collapsed);
  if (targets.length) return targets;
  const sel = window.getSelection();
  if (!sel.rangeCount) return [];
  const range = sel.getRangeAt(0);
  if (!range.collapsed) return [range.cloneRange()];
  const step = STEPS[e.inputType];
  if (!step || typeof sel.modify !== 'function') return [];
  const { anchorNode, anchorOffset } = sel;
  sel.modify('extend', ...step);
  const reach = sel.rangeCount ? sel.getRangeAt(0).cloneRange() : null;
  sel.collapse(anchorNode, anchorOffset);
  return reach && !reach.collapsed ? [reach] : [];
}

/**
 * The first tracked element under `root` that removing `ranges` would take away, or null. An element
 * is at risk when a range overlaps it without lying wholly inside it: that deletes an inline revision
 * or a table holding revisions, or merges a paragraph whose mark is tracked into its neighbour.
 */
export function trackedChangeAt(ranges, root) {
  // Also what can't go with the text around it (data-protected): notes and their numbers.
  const tracked = root.querySelectorAll('[data-tracked], [data-protected]');
  if (!tracked.length) return null;
  for (const r of ranges) {
    const range = toRange(r);
    if (!range || range.collapsed) continue;
    for (const el of tracked) {
      if (!range.intersectsNode(el)) continue;
      if (el.contains(range.startContainer) && el.contains(range.endContainer)) continue;
      return el;
    }
  }
  return null;
}

/** The status-bar explanation for a refused edit. */
export function refusalMessage(el) {
  if (el.dataset.protected && !el.dataset.tracked) {
    if (el.classList.contains('note-mark')) return 'Not deleted: the note’s number stays – Word numbers notes. To remove the note, delete its number in the text.';
    return `Not deleted: that would remove ${el.dataset.protected}. Select within one note, or delete a note’s number in the text to remove it.`;
  }
  // Revisions shown on their own (and paragraph breaks) are accepted or rejected here; those inside
  // tables' structure or locked content only in Word.
  const here = el.matches('.inserted, .deleted, [data-rev-mark]');
  return `Not deleted: that would remove ${el.dataset.tracked}. Accept or reject it first${here ? ' (on the toolbar)' : ' in Word'}, or select only the text around it.`;
}

/** The comment range markers under `root`: take this before an edit, for restoreCommentMarkers. */
export function commentMarkers(root) {
  return [...root.querySelectorAll('[data-comment-start], [data-comment-end]')];
}

/**
 * Puts back, at the caret (where the text was deleted), the markers in `before` that an edit removed
 * although their comment's mark is still there. Browsers delete these empty markers along with a
 * neighbouring character, not only when they are selected, so what was lost is found afterwards.
 */
export function restoreCommentMarkers(before, root) {
  if (!before.length) return;
  const marks = new Set([...root.querySelectorAll('[data-comment-ref]')].map((el) => el.dataset.commentRef));
  const lost = before.filter((el) => marks.has(el.dataset.commentStart ?? el.dataset.commentEnd)
    && !root.querySelector(`[data-raw="${el.dataset.raw}"]`));
  const sel = window.getSelection();
  if (!lost.length || !sel.rangeCount || !root.contains(sel.getRangeAt(0).startContainer)) return;
  const range = sel.getRangeAt(0);
  range.collapse(true);
  const frag = document.createDocumentFragment();
  frag.append(...lost);
  range.insertNode(frag);
  range.setStartAfter(lost[lost.length - 1]);
  range.collapse(true);
  sel.removeAllRanges();
  sel.addRange(range);
}

const BLOCKS = 'p, h1, h2, h3, h4, h5, h6, li';

/** True if `range` holds nothing but hidden markers (comment ranges, bookmarks, field codes). */
export function onlyHiddenMarkers(range) {
  const frag = toRange(range)?.cloneContents();
  return !!frag?.firstElementChild && !frag.textContent && !frag.querySelector(':not(.raw-hidden)');
}

/** A paragraph's deletable units in order: text nodes, line breaks and atoms (non-editable content). */
function units(el, out = []) {
  for (const n of el.childNodes) {
    if (n.nodeType === 3) { if (n.data) out.push(n); continue; }
    if (n.nodeType !== 1 || n.classList.contains('raw-hidden')) continue;
    if (n.tagName === 'BR' || n.getAttribute('contenteditable') === 'false') out.push(n);
    else units(n, out);
  }
  return out;
}

/**
 * What a Backspace (`backward`) or Delete at `range` should remove instead of the hidden markers it
 * would hit: the character (or line break, or non-editable item) beyond them in the same paragraph,
 * or else the break to the neighbouring paragraph of the same editing host, merging the two. Null if
 * there is nothing there to delete.
 */
export function beyondHiddenMarkers(range, backward) {
  const at = toRange(range);
  const from = at.startContainer.nodeType === 1 ? at.startContainer : at.startContainer.parentElement;
  const block = from.closest(BLOCKS);
  const host = from.closest('[contenteditable="true"]');
  if (!host) return null;
  // Inside a field within a line, the field is the paragraph: its text only, never a merge.
  const scope = block && host.contains(block) ? block : host;
  const r = document.createRange();
  const list = units(scope);
  // A unit is before the markers if it ends at or before the range's start, after if it starts at
  // or after its end.
  const endsBefore = (u) => (u.nodeType === 3 ? at.comparePoint(u, u.data.length) <= 0 : at.comparePoint(u.parentNode, [...u.parentNode.childNodes].indexOf(u) + 1) <= 0);
  const startsAfter = (u) => (u.nodeType === 3 ? at.comparePoint(u, 0) >= 0 : at.comparePoint(u.parentNode, [...u.parentNode.childNodes].indexOf(u)) >= 0);
  const unit = backward ? list.filter(endsBefore).pop() : list.find(startsAfter);
  if (unit?.nodeType === 3) {
    // One character, keeping surrogate pairs together.
    const { data } = unit;
    const pair = backward ? /[\uDC00-\uDFFF]/.test(data[data.length - 1]) : /[\uD800-\uDBFF]/.test(data[0]);
    const size = pair && data.length > 1 ? 2 : 1;
    r.setStart(unit, backward ? data.length - size : 0);
    r.setEnd(unit, backward ? data.length : size);
    return r;
  }
  if (unit) {
    r.selectNode(unit);
    return r;
  }
  if (scope !== block) return null;
  const blocks = [...host.querySelectorAll(BLOCKS)].filter((b) => b.closest('[contenteditable="true"]') === host);
  const other = blocks[blocks.indexOf(block) + (backward ? -1 : 1)];
  if (!other) return null;
  const [first, second] = backward ? [other, block] : [block, other];
  r.setStart(first, first.childNodes.length);
  r.setEnd(second, 0);
  return r;
}
