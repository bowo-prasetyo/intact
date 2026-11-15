// Small DOM helpers for Word-like keyboard behaviour in lists and tables.

/** True if the block has no text and nothing else visible (images, preserved content). */
export function isEmptyBlock(block) {
  return !block.textContent && !block.querySelector('img, .raw-inline:not(.raw-hidden)');
}

/** True if nothing visible comes before the collapsed `range` inside `block`. */
export function isAtBlockStart(range, block) {
  const before = document.createRange();
  before.setStart(block, 0);
  before.setEnd(range.startContainer, range.startOffset);
  const frag = before.cloneContents();
  return !frag.textContent && !frag.querySelector('img, .raw-inline:not(.raw-hidden)');
}

/** The editable cell `step` cells after (or before) `cell` in reading order, skipping merged cells. */
export function adjacentCell(cell, step) {
  const table = cell.closest('table');
  const cells = [...table.rows].flatMap((r) => [...r.cells]).filter((c) => c === cell || !c.classList.contains('vmerge'));
  return cells[cells.indexOf(cell) + step] || null;
}

/** Selects the text of a cell, as Word does when Tab moves into it. */
export function selectCell(cell) {
  const blocks = cell.querySelectorAll(':scope > :is(p, h1, h2, h3, h4, h5, h6)');
  const first = blocks[0] || cell;
  const last = blocks[blocks.length - 1] || cell;
  const range = document.createRange();
  range.setStart(first, 0);
  if (!isEmptyBlock(cell)) range.setEnd(last, last.childNodes.length);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
}
