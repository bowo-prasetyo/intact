// Read-only content in plain language. Clicking such content, or typing while the selection is in
// it, says what it is and where to edit it, instead of doing nothing (or, worse, typing wherever
// the caret was before: browsers don't move the caret when you click non-editable content).

// Tables are not editable themselves, but their cells are: they never count as the read-only part.
const STRUCTURE = new Set(['TABLE', 'TBODY', 'THEAD', 'TFOOT', 'TR']);

/**
 * The read-only element of editing host `root` that `node` is in, or null: the outermost
 * non-editable element below the nearest editable one. Interactive fields (check boxes, drop-down
 * lists, dates and the frames of text fields) and hidden markers don't count.
 */
export function readOnlyAt(node, root) {
  let el = node?.nodeType === 1 ? node : node?.parentElement;
  if (!el || !root.contains(el)) return null;
  let found = null;
  for (; el && el !== root; el = el.parentElement) {
    const ce = el.getAttribute('contenteditable');
    if (ce === 'true') break;
    if (ce === 'false' && !STRUCTURE.has(el.tagName)) found = el;
  }
  if (!found || found.classList.contains('raw-hidden')) return null;
  if (found.classList.contains('sdt') && !found.classList.contains('sdt-locked')) return null;
  return found;
}

/** What a read-only element is and where to change it, for the status bar. */
export function readOnlyMessage(el, root) {
  let titled = el;
  while (titled && titled !== root && !titled.title) titled = titled.parentElement;
  const title = titled && titled !== root ? titled.title : '';
  if (!title) return 'Read-only: this part can’t be edited here – edit it in Word';
  return `Read-only: ${title}${/\bWord\b/.test(title) ? '' : ' – kept as it is; edit it in Word'}`;
}

/** True for a key press that would type, delete or start a new paragraph. */
export function isEditingKey(e) {
  if (e.ctrlKey || e.metaKey || e.altKey || e.isComposing) return false;
  return e.key.length === 1 || e.key === 'Backspace' || e.key === 'Delete' || e.key === 'Enter';
}

/** The offset of DOM point (node, offset) in the text of `el`. */
function textOffset(el, node, offset) {
  let n = 0;
  const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
  for (let t = walker.nextNode(); t; t = walker.nextNode()) {
    if (t === node) return n + offset;
    n += t.data.length;
  }
  return n;
}

/**
 * A collapsed range just before or just after inline read-only element `el`, on the side of the
 * point (x, y) that was clicked, as Word places the caret beside an object. Null for block-level
 * content (a paragraph-sized block or a table cell), where no caret position is beside it.
 */
export function caretBeside(el, x, y) {
  if (el.tagName !== 'SPAN') return null;
  let after;
  const pos = document.caretPositionFromPoint?.(x, y);
  const hit = pos ? [pos.offsetNode, pos.offset] : (() => {
    const r = document.caretRangeFromPoint?.(x, y);
    return r ? [r.startContainer, r.startOffset] : [null, 0];
  })();
  const length = el.textContent.length;
  if (hit[0] && hit[0].nodeType === 3 && el.contains(hit[0]) && length) {
    after = textOffset(el, hit[0], hit[1]) * 2 >= length;
  } else {
    const rect = el.getBoundingClientRect();
    after = x > rect.left + rect.width / 2;
  }
  const range = document.createRange();
  if (after) range.setStartAfter(el);
  else range.setStartBefore(el);
  range.collapse(true);
  return range;
}
