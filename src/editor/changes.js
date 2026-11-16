// Finding tracked changes in the editor, for Accept, Reject, Previous and Next. An item is
// { el, kind }:
//   inline  someone's insertion or deletion (a span), or one of this session (data-tc);
//   run     a tracked formatting change of a run (data-fmt);
//   mark    a tracked paragraph break (the paragraph; data-tc-mark or data-rev-mark);
//   para    a tracked change to a paragraph's formatting (data-fmt-para);
//   row     a tracked table row (data-rev-row); the changes in a deleted row are part of it;
//   cells   a tracked column: the cells inserted or deleted together (data-rev-cell, one data-rev-key),
//           listed once, at its first cell; the changes in deleted cells are part of it;
//   merge   a tracked merge or split of cells down a column (data-rev-merge, one data-rev-merge-key),
//           listed once, at its first cell, with the text it moved;
//   tprops, rprops, cprops   a tracked change to a table's, row's or cell's properties
//           (data-fmt-table, data-fmt-row, data-fmt-cell).
// Changes inside content saved verbatim (locked fields and cells, preserved blocks) aren't
// listed one by one: each such piece (a locked field or cell, a field) is one item, kind 'locked',
// whose changes are all accepted or rejected together.

const BLOCKS = 'p, h1, h2, h3, h4, h5, h6, li';

const isInline = (el) => !!el.dataset.tc || (el.classList.contains('raw-inline') && (el.classList.contains('inserted') || el.classList.contains('deleted')));
const isMark = (el) => el.matches(BLOCKS) && (!!el.dataset.tcMark || !!el.dataset.revMark);

/** True for this session's text moved by a tracked merge: it is accepted or rejected with the merge. */
function partOfMerge(el, kind) {
  if (kind !== 'inline' && kind !== 'mark') return false;
  const cell = el.closest('td');
  if (!cell?.dataset.revMerge) return false;
  const key = kind === 'mark' ? `${el.dataset.tcMarkAuthor}|${el.dataset.tcMarkDate}` : `${el.dataset.tcAuthor}|${el.dataset.tcDate}`;
  return key === cell.dataset.revMergeKey;
}

/** The cells of the tracked merge `td` belongs to: same table, same data-rev-merge-key. */
export function mergeGroup(td) {
  const table = td.closest('table');
  return [...table.querySelectorAll(':scope > tbody > tr > td, :scope > tr > td')].filter((c) => c.dataset.revMergeKey === td.dataset.revMergeKey);
}

/** The cells of the tracked column `td` belongs to: same table, same data-rev-key. */
export function cellGroup(td) {
  const table = td.closest('table');
  return [...table.querySelectorAll(':scope > tbody > tr > td, :scope > tr > td')].filter((c) => c.dataset.revKey === td.dataset.revKey);
}

/** True unless `el` sits in content that is saved verbatim (so a change there wouldn't be kept). */
function resolvable(el) {
  const holder = el.parentElement?.closest('[data-raw], td[contenteditable="false"]');
  return !holder || holder.tagName === 'TABLE';
}


/** The tracked changes under `root`, in document order (a paragraph's own changes after its text). */
export function changeItems(root) {
  const out = [];
  const seen = new Set(); // tracked columns listed: table + key
  const add = (el, kind) => { if (resolvable(el) && !partOfMerge(el, kind)) out.push({ el, kind, mark: kind === 'mark' }); };
  const walk = (node) => {
    for (const el of node.children) {
      // Content kept as it is, holding tracked changes: one item for all of them.
      if (el.dataset.raw != null && el.dataset.tracked && el.tagName !== 'TABLE' && !isInline(el)) {
        add(el, 'locked');
        continue;
      }
      if (el.tagName === 'TD' && el.contentEditable === 'false' && el.querySelector('[data-tracked]')) {
        add(el, 'locked');
        continue;
      }
      if (el.tagName === 'TABLE' && el.dataset.fmtTable) add(el, 'tprops');
      if (el.tagName === 'TD' && el.dataset.fmtCell) add(el, 'cprops');
      if (el.tagName === 'TD' && el.dataset.revMerge) {
        const table = el.closest('table');
        if (![...seen].some(([t, k]) => t === table && k === `merge|${el.dataset.revMergeKey}`)) {
          seen.add([table, `merge|${el.dataset.revMergeKey}`]);
          add(el, 'merge');
        }
      }
      if (el.tagName === 'TD' && el.dataset.revCell) {
        const key = el.dataset.revKey;
        const table = el.closest('table');
        if (![...seen].some(([t, k]) => t === table && k === key)) {
          seen.add([table, key]);
          add(el, 'cells');
        }
        if (el.dataset.revCell === 'del') continue;
      }
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
      if (el.tagName === 'TR' && el.dataset.fmtRow) add(el, 'rprops');
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
  } else if (kind === 'cells' || kind === 'merge') {
    const cells = kind === 'cells' ? cellGroup(el) : mergeGroup(el);
    const last = cells[cells.length - 1];
    range.setStart(cells[0], 0);
    range.setEnd(last, last.childNodes.length);
  } else if (kind === 'locked') {
    if (el.tagName === 'TD') range.selectNodeContents(el);
    else range.selectNode(el);
  } else if (kind === 'cprops') {
    range.selectNodeContents(el);
  } else if (kind === 'tprops') {
    range.setStartBefore(el); // the table's own change comes before what is in it
    range.collapse(true);
  } else if (kind === 'rprops') {
    const last = el.cells[el.cells.length - 1]; // the row's own change comes after its cells
    range.selectNodeContents(last);
    range.collapse(false);
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
      : kind === 'cells' || kind === 'merge' ? (kind === 'cells' ? cellGroup(el) : mergeGroup(el)).some((td) => range.intersectsNode(td)) : range.intersectsNode(el)));
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
  const locked = find((i) => i.kind === 'locked' && (i.el.contains(el) || i.el === before || i.el === after));
  if (locked) return [locked];
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
  const merged = el.closest('td[data-rev-merge]');
  const merge = merged && find((i) => i.kind === 'merge' && mergeGroup(i.el).includes(merged));
  if (merge) return [merge];
  const cell = el.closest('td[data-rev-cell]');
  const column = cell && find((i) => i.kind === 'cells' && cellGroup(i.el).includes(cell));
  if (column) return [column];
  const row = find((i) => i.kind === 'row' && i.el.contains(el));
  if (row) return [row];
  const cellProps = find((i) => i.kind === 'cprops' && i.el.contains(el));
  if (cellProps) return [cellProps];
  const rowProps = find((i) => i.kind === 'rprops' && i.el.contains(el));
  if (rowProps) return [rowProps];
  const table = el.closest('table');
  const tableProps = table && find((i) => i.kind === 'tprops' && i.el === table);
  return tableProps ? [tableProps] : [];
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
