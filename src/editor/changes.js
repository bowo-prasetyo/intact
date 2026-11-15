// Finding tracked changes in the editor, for Accept, Reject, Previous and Next: inline revisions
// (someone's insertion or deletion, or one of this session) and paragraph breaks that are tracked.
// Revisions inside content saved verbatim (locked fields and cells, preserved blocks) and table-row
// revisions can only be accepted or rejected in Word, and aren't listed.

const BLOCKS = 'p, h1, h2, h3, h4, h5, h6, li';

const isInline = (el) => !!el.dataset.tc || (el.classList.contains('raw-inline') && (el.classList.contains('inserted') || el.classList.contains('deleted')));
const isMark = (el) => el.matches(BLOCKS) && (!!el.dataset.tcMark || !!el.dataset.revMark);

/** True unless `el` sits in content that is saved verbatim (so a change there wouldn't be kept). */
function resolvable(el) {
  const holder = el.parentElement?.closest('[data-raw], td[contenteditable="false"]');
  return !holder || holder.tagName === 'TABLE';
}

/**
 * The tracked changes under `root`, in document order: [{ el, mark }], where a paragraph break
 * (mark: true) comes after its paragraph's text.
 */
export function changeItems(root) {
  const out = [];
  const walk = (node) => {
    for (const el of node.children) {
      if (isInline(el)) {
        if (resolvable(el)) out.push({ el, mark: false });
        continue;
      }
      walk(el);
      if (isMark(el) && resolvable(el)) out.push({ el, mark: true });
    }
  };
  walk(root);
  return out;
}

/** The changes that can only be accepted or rejected in Word: those in verbatim content, and tables'. */
export function wordOnlyChanges(root) {
  let n = 0;
  for (const el of root.querySelectorAll('[data-tracked]')) {
    if (isInline(el) || isMark(el)) { if (!resolvable(el)) n++; } else if (el.closest('[data-raw]') === el) n++;
  }
  return n;
}

/** Where a change is: the element for an inline one, the end of the paragraph for a break. */
export function itemRange(item) {
  const range = document.createRange();
  if (item.mark) {
    range.selectNodeContents(item.el);
    range.collapse(false);
  } else {
    range.selectNode(item.el);
  }
  return range;
}

/** The node just before (dir -1) or after (dir 1) a caret, skipping empty text and hidden markers. */
function neighbour(node, offset, dir, block) {
  let n;
  if (node.nodeType === 3) {
    if (dir < 0 ? offset > 0 : offset < node.length) return null; // a character is next to it
    n = node;
    while (n && n !== block && !(dir < 0 ? n.previousSibling : n.nextSibling)) n = n.parentNode;
    n = n && n !== block ? (dir < 0 ? n.previousSibling : n.nextSibling) : null;
  } else {
    n = dir < 0 ? node.childNodes[offset - 1] : node.childNodes[offset];
    if (!n && node !== block) return neighbour(node.parentNode, [...node.parentNode.childNodes].indexOf(node) + (dir < 0 ? 0 : 1), dir, block);
  }
  while (n && ((n.nodeType === 3 && !n.data) || n.classList?.contains('raw-hidden'))) n = dir < 0 ? n.previousSibling : n.nextSibling;
  // Into formatting around it, to a change at its edge.
  while (n?.nodeType === 1 && !isInline(n) && n.getAttribute('contenteditable') !== 'false' && n.firstChild) n = dir < 0 ? n.lastChild : n.firstChild;
  return n?.nodeType === 1 ? n : null;
}

/**
 * The changes a selection is on: every change it covers; or, for a caret, the change it is in or
 * next to (as when a change was just clicked), or the paragraph break it stands before.
 */
export function changesAt(range, items) {
  if (!range) return [];
  if (!range.collapsed) {
    // A paragraph break is covered when the selection runs past the end of its paragraph.
    return items.filter((item) => (item.mark
      ? range.comparePoint(item.el, item.el.childNodes.length) === 0
      : range.intersectsNode(item.el)));
  }
  const { startContainer: node, startOffset: offset } = range;
  const el = node.nodeType === 1 ? node : node.parentElement;
  const inside = items.find((item) => !item.mark && item.el.contains(el));
  if (inside) return [inside];
  const block = el.closest(BLOCKS);
  const before = neighbour(node, offset, -1, block);
  const after = neighbour(node, offset, 1, block);
  const next = items.find((item) => !item.mark && (item.el === before || item.el === after));
  if (next) return [next];
  // At the end of a paragraph whose break is a change.
  const mark = block && items.find((item) => item.mark && item.el === block);
  if (mark) {
    const rest = document.createRange();
    rest.setStart(node, offset);
    rest.setEnd(block, block.childNodes.length);
    if (!rest.toString() && !rest.cloneContents().querySelector('img, .raw-inline:not(.raw-hidden)')) return [mark];
  }
  return [];
}

/**
 * The change after (dir 1) or before (dir -1) a selection, wrapping round at the end; null if
 * there are none. A change the selection is exactly on is skipped.
 */
export function adjacentChange(range, items, dir) {
  if (!items.length) return null;
  const on = (item) => {
    const r = itemRange(item);
    return range && r.compareBoundaryPoints(Range.START_TO_START, range) === 0 && r.compareBoundaryPoints(Range.END_TO_END, range) === 0;
  };
  if (!range) return dir > 0 ? items[0] : items[items.length - 1];
  const list = dir > 0 ? items : [...items].reverse();
  const found = list.find((item) => {
    if (on(item)) return false;
    const r = itemRange(item);
    // END_TO_START compares this range's start with the other's end; START_TO_END, its end with the other's start.
    return dir > 0
      ? r.compareBoundaryPoints(Range.END_TO_START, range) >= 0 // starts at or after the selection's end
      : r.compareBoundaryPoints(Range.START_TO_END, range) <= 0; // ends at or before its start
  });
  return found || list.find((item) => !on(item)) || list[0];
}
