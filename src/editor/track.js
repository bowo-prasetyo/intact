// Track Changes for the user's own edits, live in the editor, as Word shows them:
//  * an insertion is <span class="tc-ins" data-tc="ins" data-tc-author data-tc-date>;
//  * a deletion keeps its text, wrapped in a non-editable <span class="tc-del" data-tc="del" …>;
//  * a paragraph break inserted or deleted is data-tc-mark="ins" | "del" on the paragraph (with
//    data-tc-mark-author and -date); a deleted break doesn't merge the paragraphs.
// DocxDocument writes these as <w:ins>, <w:del> (with <w:delText>) and paragraph-mark revisions.
// Everything here is from this session, so deleting your own insertion removes it, as in Word.
// Others' revisions (labelled data-tracked when rendered) are protected by protect.js.

export const BLOCKS = 'p, h1, h2, h3, h4, h5, h6, li';

/** A revision date as Word writes it: UTC, without milliseconds. */
export function revisionDate() {
  return new Date().toISOString().replace(/\.\d{3}Z$/, 'Z');
}

function revisionTitle(kind, author) {
  return `${kind === 'ins' ? 'Inserted' : 'Deleted'} by ${author} – tracked change: Accept or Reject on the toolbar`;
}

function makeSpan(kind, { author, date = revisionDate() }) {
  const span = document.createElement('span');
  span.className = kind === 'ins' ? 'tc-ins' : 'tc-del';
  span.dataset.tc = kind;
  span.dataset.tcAuthor = author;
  span.dataset.tcDate = date;
  span.title = revisionTitle(kind, author);
  if (kind === 'del') span.contentEditable = 'false';
  return span;
}

const isEl = (n) => n?.nodeType === 1;
const insOf = (node) => (isEl(node) ? node : node?.parentElement)?.closest('[data-tc="ins"]') || null;
const hostOf = (node) => (isEl(node) ? node : node?.parentElement)?.closest('[contenteditable="true"]') || null;

function blockOf(node) {
  const el = isEl(node) ? node : node?.parentElement;
  const block = el?.closest(BLOCKS);
  const host = hostOf(node);
  return block && host?.contains(block) ? block : null;
}

/** The paragraphs of an editing host, in order (not those of tables or fields inside it). */
function hostBlocks(host) {
  return [...host.querySelectorAll(BLOCKS)].filter((b) => hostOf(b) === host);
}

/** True inside a form field: edits there are not tracked. */
export function inField(node) {
  return !!(isEl(node) ? node : node?.parentElement)?.closest('.sdt-content');
}

function setCaret(node, offset) {
  const range = document.createRange();
  range.setStart(node, offset);
  range.collapse(true);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
  return range;
}

function caretAfter(node) {
  const parent = node.parentNode;
  return setCaret(parent, [...parent.childNodes].indexOf(node) + 1);
}

function caretBefore(node) {
  const parent = node.parentNode;
  return setCaret(parent, [...parent.childNodes].indexOf(node));
}

/** True if DOM point (a, ao) comes strictly before (b, bo). */
function isBefore(a, ao, b, bo) {
  const r = document.createRange();
  r.setStart(a, ao);
  return r.comparePoint(b, bo) > 0;
}

// ------------------------------------------------------------------ insertions

/** Inserts `nodes` at the collapsed selection as one tracked insertion; the caret goes after them. */
export function insertTracked(nodes, opts) {
  const sel = window.getSelection();
  if (!sel.rangeCount) return null;
  const range = sel.getRangeAt(0);
  const span = makeSpan('ins', opts);
  span.append(...nodes);
  range.insertNode(span);
  const last = span.lastChild;
  if (last?.nodeType === 3) setCaret(last, last.length);
  else caretAfter(last || span);
  return span;
}

/**
 * Types `text` as a tracked insertion. Returns false when the browser can type it itself: the caret
 * is already in an insertion of this session, which the typed text then extends.
 */
export function typeTracked(text, opts) {
  const sel = window.getSelection();
  if (!sel.rangeCount || !sel.isCollapsed) return false;
  if (insOf(sel.anchorNode) && sel.anchorNode.nodeType === 3) return false;
  insertTracked([document.createTextNode(text)], opts);
  return true;
}

/** Wraps a line break the browser just inserted before the caret as a tracked insertion. */
export function wrapInsertedBreak(opts) {
  const sel = window.getSelection();
  if (!sel.rangeCount) return;
  const { startContainer: c, startOffset: o } = sel.getRangeAt(0);
  let br = c.nodeType === 1 ? c.childNodes[o - 1] : (o === 0 ? c.previousSibling : null);
  while (br && br.nodeType === 3 && !br.data) br = br.previousSibling;
  if (br?.tagName !== 'BR' || insOf(br)) return;
  const span = makeSpan('ins', opts);
  br.replaceWith(span);
  span.append(br);
  caretAfter(span);
}

/** Wraps text just typed before the caret (by an input method, which can't be intercepted). */
export function wrapComposedText(text, opts) {
  const sel = window.getSelection();
  if (!text || !sel.rangeCount) return;
  const { startContainer: node, startOffset: end } = sel.getRangeAt(0);
  if (node.nodeType !== 3 || insOf(node) || end < text.length || node.data.slice(end - text.length, end) !== text) return;
  const range = document.createRange();
  range.setStart(node, end - text.length);
  range.setEnd(node, end);
  const span = makeSpan('ins', opts);
  range.surroundContents(span);
  setCaret(span.firstChild, span.firstChild.length);
}

/** Marks the paragraph before the caret's as ending in an inserted paragraph break (after Enter). */
export function markSplit(opts) {
  const sel = window.getSelection();
  if (!sel.rangeCount) return;
  const block = blockOf(sel.anchorNode);
  const host = block && hostOf(block);
  if (!host) return;
  const blocks = hostBlocks(host);
  const prev = blocks[blocks.indexOf(block) - 1];
  if (!prev) return;
  prev.dataset.tcMark = 'ins';
  prev.dataset.tcMarkAuthor = opts.author;
  prev.dataset.tcMarkDate = opts.date || revisionDate();
}

/** The marker to show after a paragraph whose break is a tracked change. */
export function paragraphMarkTitle(block) {
  return revisionTitle(block.dataset.tcMark, block.dataset.tcMarkAuthor).replace(/^(\w+)/, '$1 paragraph break');
}

// ------------------------------------------------------------------ deletions

/** What a deletion can never take while tracking: hidden markers, comment marks, locked parts. */
function untouchable(el) {
  return el.classList.contains('raw-hidden') || el.classList.contains('comment-ref') || el.classList.contains('raw-block')
    || el.classList.contains('sdt-locked') || el.dataset.tc === 'del' || el.dataset.tcCaret != null
    || (el.tagName === 'TD' && el.contentEditable === 'false');
}

/** A paragraph's deletable units in order: text nodes, line breaks and non-editable items. */
function units(el, out = []) {
  for (const n of el.childNodes) {
    if (n.nodeType === 3) { if (n.data) out.push(n); continue; }
    if (!isEl(n) || untouchable(n)) continue;
    if (n.tagName === 'BR' || n.getAttribute('contenteditable') === 'false') out.push(n);
    else units(n, out);
  }
  return out;
}

const pointIndex = (n) => [...n.parentNode.childNodes].indexOf(n);

/**
 * The range a tracked Backspace (`backward`) or Delete at the caret removes: the character, line
 * break or item before or after it, skipping what is already deleted, hidden markers and comment
 * marks; or else the paragraph break between this paragraph and the previous (or next) one. A
 * break that is already deleted is stepped over, as Word does. Null at the start (or end) of the
 * text. A trailing placeholder <br> of an empty paragraph is not a unit.
 */
export function unitRange(backward) {
  const sel = window.getSelection();
  if (!sel.rangeCount) return null;
  let { startContainer: node, startOffset: offset } = sel.getRangeAt(0);
  const host = hostOf(node);
  if (!host) return null;
  let block = blockOf(node) || host;
  for (let guard = 0; guard < 10000; guard++) {
    const at = document.createRange();
    at.setStart(node, offset);
    // A <br> with nothing after it in the paragraph is the browser's placeholder, not a line break.
    const trailing = (u) => {
      for (let n = u; n && n !== block; n = n.parentNode) {
        let s = n.nextSibling;
        while (s && s.nodeType === 3 && !s.data) s = s.nextSibling;
        if (s) return false;
      }
      return true;
    };
    const list = units(block).filter((u) => !(u.tagName === 'BR' && trailing(u)));
    let unit = null;
    if (backward) {
      for (const u of list) {
        if (u === node && u.nodeType === 3) { if (offset > 0) unit = u; break; }
        const end = u.nodeType === 3 ? [u, u.length] : [u.parentNode, pointIndex(u) + 1];
        if (at.comparePoint(...end) <= 0) unit = u; else break;
      }
    } else {
      for (const u of list) {
        if (u === node && u.nodeType === 3) { if (offset < u.length) { unit = u; break; } continue; }
        const start = u.nodeType === 3 ? [u, 0] : [u.parentNode, pointIndex(u)];
        if (at.comparePoint(...start) >= 0) { unit = u; break; }
      }
    }
    const r = document.createRange();
    if (unit?.nodeType === 3) {
      const { data } = unit;
      const from = unit === node ? offset : (backward ? data.length : 0);
      const pair = backward ? /[\uDC00-\uDFFF]/.test(data[from - 1]) : /[\uD800-\uDBFF]/.test(data[from]);
      const size = pair ? 2 : 1;
      if (backward) { r.setStart(unit, from - size); r.setEnd(unit, from); } else { r.setStart(unit, from); r.setEnd(unit, from + size); }
      return r;
    }
    if (unit) {
      r.selectNode(unit);
      return r;
    }
    if (block === host) return null;
    const blocks = hostBlocks(host);
    const i = blocks.indexOf(block);
    const [first, second] = backward ? [blocks[i - 1], block] : [block, blocks[i + 1]];
    if (!first || !second) return null;
    if (first.dataset.tcMark === 'del') {
      // Already deleted: carry on in the neighbouring paragraph.
      block = backward ? first : second;
      [node, offset] = backward ? [first, first.childNodes.length] : [second, 0];
      continue;
    }
    r.setStart(first, first.childNodes.length);
    r.setEnd(second, 0);
    return r;
  }
  return null;
}

/**
 * Deletes `range` as tracked changes: text and items become deletions (or, if they were inserted
 * this session, go), and paragraph breaks inside the range are marked deleted (or, if inserted this
 * session, removed, merging the paragraphs). The caret goes to the start of the range (`backward`)
 * or its end. Hidden markers, comment marks, locked fields and table structure are left alone.
 */
export function deleteTracked(range, backward, opts) {
  const date = opts.date || revisionDate();
  // Markers at both ends, so the caret can be placed after the DOM has changed.
  const marker = () => { const m = document.createElement('span'); m.dataset.tcCaret = ''; return m; };
  const endMark = marker();
  const startMark = marker();
  const end = range.cloneRange(); end.collapse(false); end.insertNode(endMark);
  const start = range.cloneRange(); start.collapse(true); start.insertNode(startMark);
  const span = document.createRange();
  span.setStartAfter(startMark);
  span.setEndBefore(endMark);

  // Leaves: text portions and whole items inside the range.
  const leaves = [];
  const root = span.commonAncestorContainer.nodeType === 1 ? span.commonAncestorContainer : span.commonAncestorContainer.parentNode;
  const walk = (el) => {
    for (const n of [...el.childNodes]) {
      if (!span.intersectsNode(n)) continue;
      if (n.nodeType === 3) {
        // Text inside a non-editable item (or a deletion) goes with the item, or stays.
        const holder = n.parentElement.closest('[contenteditable]');
        if (!n.data || holder?.getAttribute('contenteditable') !== 'true') continue;
        const s = n === span.startContainer ? span.startOffset : 0;
        const e = n === span.endContainer ? span.endOffset : n.length;
        if (e > s) leaves.push({ text: n, s, e });
        continue;
      }
      if (!isEl(n) || untouchable(n)) continue;
      const ce = n.getAttribute('contenteditable');
      if (n.tagName === 'BR' || (ce === 'false' && n.tagName !== 'TABLE')) {
        // Only whole items: both of their edges inside the range.
        const i = pointIndex(n);
        if (span.comparePoint(n.parentNode, i) === 0 && span.comparePoint(n.parentNode, i + 1) === 0) leaves.push({ node: n });
        continue;
      }
      walk(n);
    }
  };
  walk(root);

  // Paragraph breaks inside the range: paragraphs that end before the range does (never the last
  // paragraph of the text or of a table cell, whose break can't be deleted).
  const marks = [];
  const scope = root.closest?.(BLOCKS) ? root.closest(BLOCKS).parentElement : root;
  for (const b of scope.querySelectorAll(BLOCKS)) {
    if (!span.intersectsNode(b)) continue;
    const host = hostOf(b);
    const blocks = host ? hostBlocks(host) : [];
    if (blocks[blocks.length - 1] === b) continue;
    if (isBefore(b, b.childNodes.length, endMark.parentNode, pointIndex(endMark))
      && !isBefore(b, b.childNodes.length, startMark.parentNode, pointIndex(startMark))) marks.push(b);
  }

  // Apply, last first, so earlier positions stay valid.
  const created = [];
  for (const leaf of leaves.reverse()) {
    if (leaf.text) {
      const { text, s, e } = leaf;
      if (insOf(text)) { text.deleteData(s, e - s); continue; }
      let part = text;
      if (s > 0) part = part.splitText(s);
      if (part.length > e - s) part.splitText(e - s);
      const del = makeSpan('del', { ...opts, date });
      part.replaceWith(del);
      del.append(part);
      created.push(del);
    } else if (insOf(leaf.node)) {
      leaf.node.remove();
    } else {
      const del = makeSpan('del', { ...opts, date });
      leaf.node.replaceWith(del);
      del.append(leaf.node);
      created.push(del);
    }
  }
  for (const b of marks.reverse()) {
    if (b.dataset.tcMark === 'ins') {
      // A break inserted this session: remove it, merging the paragraphs. The merged paragraph keeps
      // the next paragraph's properties, as the surviving break is that one's.
      const blocks = hostBlocks(hostOf(b));
      const next = blocks[blocks.indexOf(b) + 1];
      const kids = [...b.childNodes].filter((n) => !(n.tagName === 'BR' && n === b.lastChild && !b.textContent));
      next.prepend(...kids);
      b.remove();
    } else if (b.dataset.tcMark !== 'del') {
      b.dataset.tcMark = 'del';
      b.dataset.tcMarkAuthor = opts.author;
      b.dataset.tcMarkDate = date;
    }
  }
  for (const del of created) mergeDeletions(del);
  // Insertions emptied by the deletion go.
  for (const ins of (root.isConnected ? root : document).querySelectorAll('[data-tc="ins"]')) {
    if (!ins.textContent && !ins.querySelector('br, img, [data-raw]')) ins.remove();
  }

  const sel = window.getSelection();
  const at = backward ? startMark : endMark;
  const caret = document.createRange();
  if (at.isConnected) {
    caret.setStartBefore(at);
  } else {
    caret.setStart(endMark.isConnected ? endMark.parentNode : root, 0);
  }
  caret.collapse(true);
  startMark.remove();
  endMark.remove();
  sel.removeAllRanges();
  sel.addRange(caret);
  return caret;
}

/** Joins a new deletion with a neighbouring one, so repeated Backspace makes one deletion. */
function mergeDeletions(del) {
  if (!del.isConnected) return;
  const sibling = (dir) => {
    let s = dir < 0 ? del.previousSibling : del.nextSibling;
    // Empty text and the caret markers of deleteTracked don't separate deletions.
    while (s && ((s.nodeType === 3 && !s.data) || (isEl(s) && s.dataset.tcCaret != null))) s = dir < 0 ? s.previousSibling : s.nextSibling;
    return isEl(s) && s.dataset.tc === 'del' && s.dataset.tcAuthor === del.dataset.tcAuthor ? s : null;
  };
  const prev = sibling(-1);
  if (prev) { prev.append(...del.childNodes); del.remove(); del = prev; }
  const next = sibling(1);
  if (next) { del.append(...next.childNodes); next.remove(); }
  del.normalize(); // one text node, so one <w:delText>
}

/**
 * Replaces `range` with `text` as tracked changes: a deletion, then an insertion after it. Returns
 * the point just after the replacement.
 */
export function replaceTracked(range, text, opts) {
  const date = opts.date || revisionDate();
  const caret = deleteTracked(range, false, { ...opts, date });
  if (!text) return { node: caret.startContainer, offset: caret.startOffset };
  const node = document.createTextNode(text);
  insertTracked([node], { ...opts, date });
  return { node, offset: node.length };
}

/** The insertion that ends right before DOM point (node, offset) in its paragraph, if any. */
function insEndingAt(node, offset) {
  let prev = null;
  if (node.nodeType === 3) {
    if (offset > 0) return null;
    for (let n = node; n && !prev; n = n.parentNode) {
      if (isEl(n) && (n.matches(BLOCKS) || n.getAttribute('contenteditable'))) return null;
      prev = n.previousSibling;
      while (prev && prev.nodeType === 3 && !prev.data) prev = prev.previousSibling;
    }
  } else {
    prev = node.childNodes[offset - 1] || null;
  }
  while (prev && prev.nodeType === 3 && !prev.data) prev = prev.previousSibling;
  while (isEl(prev) && prev.dataset.tc !== 'ins' && prev.lastChild) prev = prev.lastChild;
  if (isEl(prev) && prev.dataset.tc === 'ins') return prev;
  return prev ? insOf(prev) : null;
}

/**
 * Types `text` untracked while the caret is in (or at the end of) an insertion of this session:
 * the insertion is split around it, as Word does with Track Changes off. Returns false when the
 * caret isn't in an insertion, so the browser can type it itself.
 */
export function typeUntracked(text) {
  const sel = window.getSelection();
  if (!sel.rangeCount || !sel.isCollapsed) return false;
  const ins = insOf(sel.anchorNode);
  if (!ins) {
    // Right after an insertion, the browser would type into it: put the text outside it.
    if (!insEndingAt(sel.anchorNode, sel.anchorOffset)) return false;
    const node = document.createTextNode(text);
    sel.getRangeAt(0).insertNode(node);
    setCaret(node, node.length);
    return true;
  }
  const rest = document.createRange();
  rest.setStart(sel.anchorNode, sel.anchorOffset);
  rest.setEnd(ins, ins.childNodes.length);
  const tail = ins.cloneNode(false);
  tail.append(rest.extractContents());
  const node = document.createTextNode(text);
  ins.after(node);
  if (tail.textContent || tail.querySelector('br, img, [data-raw]')) node.after(tail);
  if (!ins.textContent && !ins.querySelector('br, img, [data-raw]')) ins.remove();
  setCaret(node, node.length);
  return true;
}
