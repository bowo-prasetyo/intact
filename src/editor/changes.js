// Finding tracked changes in the editor, for Accept, Reject, Previous and Next. An item is
// { el, kind }:
//   inline  someone's insertion or deletion (a span), or one of this session (data-tc);
//   run     a tracked formatting change of a run (data-fmt);
//   mark    a tracked paragraph break (the paragraph; data-tc-mark or data-rev-mark);
//   para    a tracked change to a paragraph's formatting (data-fmt-para);
//   row     a tracked table row (data-rev-row); the changes in a deleted row are part of it.
// Changes inside content saved verbatim (locked fields and cells, preserved blocks) can only be
// accepted or rejected in Word, and aren't listed (DocxDocument.wordOnlyChanges counts them).

const BLOCKS = 'p, h1, h2, h3, h4, h5, h6, li';

const isInline = (el) => !!el.dataset.tc || (el.classList.contains('raw-inline') && (el.classList.contains('inserted') || el.classList.contains('deleted')));
const isMark = (el) => el.matches(BLOCKS) && (!!el.dataset.tcMark || !!el.dataset.revMark);

/** True unless `el` sits in content that is saved verbatim (so a change there wouldn't be kept). */
function resolvable(el) {
  const holder = el.parentElement?.closest('[data-raw], td[contenteditable="false"]');
  return !holder || holder.tagName === 'TABLE';
}

/** The tracked changes under `root`, in document order (a paragraph's own changes after its text). */
export function changeItems(root) {
  const out = [];
  const add = (el, kind) => { if (resolvable(el)) out.push({ el, kind, mark: kind === 'mark' }); };
  const walk = (node) => {
    for (const el of node.children) {
      if (isInline(el)) {
        add(el, 'inline');
        continue;
      }
      // A deleted row is one change: accepting it removes the row, rejecting it brings its text back.
      if (el.tagName === 'TR' && el.dataset.revRow === 'del') {
        add(el, 'row');
        continue;
      }
      if (el.dataset.fmt === 'run') add(el, 'run');
      walk(el);
      if (el.matches(BLOCKS)) {
        if (el.dataset.fmtPara) add(el, 'para');
        if (isMark(el)) add(el, 'mark');
      }
      if (el.tagName === 'TR' && el.dataset.revRow) add(el, 'row');
    }
  };
  walk(root);
  return out;
}

/** Where a change is, to select it: the span; a run's text; a paragraph's end (its break) or text; a row's cells. */
export function itemRange(item) {
  const range = document.createRange();
  const { el, kind } = item;
  if (kind === 'mark') {
    range.selectNodeContents(el);
    range.collapse(false);
  } else if (kind === 'run' || kind === 'para') {
    range.selectNodeContents(el);
  } else if (kind === 'row') {
    const cells = [...el.cells];
    range.setStart(cells[0], 0);
    range.setEnd(cells[cells.length - 1], cells[cells.length - 1].childNodes.length);
  } else {
    range.selectNode(el);
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
    if (!n && node !== block && node.parentNode) return neighbour(node.parentNode, [...node.parentNode.childNodes].indexOf(node) + (dir < 0 ? 0 : 1), dir, block);
  }
  while (n && ((n.nodeType === 3 && !n.data) || n.classList?.contains('raw-hidden'))) n = dir < 0 ? n.previousSibling : n.nextSibling;
  // Into formatting around it, to a change at its edge.
  while (n?.nodeType === 1 && !isInline(n) && n.getAttribute('contenteditable') !== 'false' && n.firstChild) n = dir < 0 ? n.lastChild : n.firstChild;
  return n?.nodeType === 1 ? n : null;
}

/**
 * The changes a selection is on: every change it covers (a paragraph's own changes when it runs past
 * the paragraph's end); or, for a caret, the change it is in or next to (as when one was just
 * clicked), else a formatting change around it, the break it stands before, its paragraph's
 * formatting change, or its row.
 */
export function changesAt(range, items) {
  if (!range) return [];
  if (!range.collapsed) {
    return items.filter(({ el, kind }) => (kind === 'mark' || kind === 'para'
      ? range.comparePoint(el, el.childNodes.length) === 0
      : range.intersectsNode(el)));
  }
  const { startContainer: node, startOffset: offset } = range;
  const el = node.nodeType === 1 ? node : node.parentElement;
  const find = (test) => items.find(test);
  const inside = find((i) => i.kind === 'inline' && i.el.contains(el));
  if (inside) return [inside];
  const block = el.closest(BLOCKS);
  const before = neighbour(node, offset, -1, block);
  const after = neighbour(node, offset, 1, block);
  const next = find((i) => i.kind === 'inline' && (i.el === before || i.el === after));
  if (next) return [next];
  const run = find((i) => i.kind === 'run' && i.el.contains(el));
  if (run) return [run];
  const mark = block && find((i) => i.kind === 'mark' && i.el === block);
  if (mark) {
    const rest = document.createRange();
    rest.setStart(node, offset);
    rest.setEnd(block, block.childNodes.length);
    if (!rest.toString() && !rest.cloneContents().querySelector('img, .raw-inline:not(.raw-hidden)')) return [mark];
  }
  const para = block && find((i) => i.kind === 'para' && i.el === block);
  if (para) return [para];
  const row = find((i) => i.kind === 'row' && i.el.contains(el));
  return row ? [row] : [];
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
