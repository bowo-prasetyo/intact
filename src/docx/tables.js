// Changing a table's structure as Word does: rows inserted above or below, columns inserted left
// or right, rows and columns deleted, cells merged and split. Rows and cells inside content
// controls (repeating sections, fields around cells) are found through them; a row added in a
// repeating section is a new item, and a control left empty goes. Everything here works on a table's XML (<w:tbl>), in grid
// columns (<w:tblGrid>): a cell spanning columns (gridSpan) is widened or narrowed rather than
// split, vertical merges (vMerge) carry on through rows inserted inside them, and widths follow
// the grid. DocxDocument keeps the rendered table in step.
import { NS, wChild, wChildren, wAttr, createW, removeW, elementChildren, ensureOrdered } from './xml.js';

const W = NS.w;
const TCPR_ORDER = [
  'cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd', 'noWrap', 'tcMar', 'textDirection',
  'tcFitText', 'vAlign', 'hideMark', 'headers', 'cellIns', 'cellDel', 'cellMerge', 'tcPrChange',
];
const REVISION_MARKS = ['ins', 'del', 'moveTo', 'moveFrom'];

const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const isW = (n, local) => n?.nodeType === 1 && n.namespaceURI === W && n.localName === local;

/** Rows (`local` 'tr') or cells ('tc') of `parent` in order, looking through content controls and custom XML around them. */
function parts(parent, local) {
  const out = [];
  for (const c of elementChildren(parent)) {
    if (c.namespaceURI !== W) continue;
    if (c.localName === local) out.push(c);
    else if (c.localName === 'sdt') out.push(...parts(wChild(c, 'sdtContent') || c.ownerDocument.createElement('x'), local));
    else if (c.localName === 'customXml') out.push(...parts(c, local));
  }
  return out;
}

export const rowsOf = (tbl) => parts(tbl, 'tr');
export const cellsOf = (tr) => parts(tr, 'tc');

/** The content controls (and custom XML) around a row or cell, innermost first, up to its table or row. */
function wrappersOf(el) {
  const out = [];
  for (let n = el.parentNode; n && !isW(n, 'tbl') && !isW(n, 'tr'); n = n.parentNode) {
    if (isW(n, 'sdt') || isW(n, 'customXml')) out.push(n);
  }
  return out;
}

/** The element standing for a row or cell among its table's or row's children: its outermost wrapper, or itself. */
function outermost(el) {
  const w = wrappersOf(el);
  return w.length ? w[w.length - 1] : el;
}

/** Takes a row or cell out, and any content control left empty by it. */
export function removeRow(el) {
  detach(el);
}

function detach(el) {
  let parent = el.parentNode;
  parent.removeChild(el);
  while (parent && (isW(parent, 'sdtContent') || isW(parent, 'customXml'))) {
    if (elementChildren(parent).some((c) => !(c.namespaceURI === W && ['bookmarkStart', 'bookmarkEnd', 'proofErr'].includes(c.localName)))) break;
    const holder = isW(parent, 'sdtContent') ? parent.parentNode : parent;
    const up = holder.parentNode;
    up.removeChild(holder);
    parent = up;
  }
}

/**
 * Why a table's rows or cells can't be changed here, or null: content controls around them that
 * are locked, or linked to the document's data (Word would overwrite them). Only what a change
 * touches counts, so a locked row (a form's total, say) doesn't stop changes elsewhere in the
 * table: `rows` ([first, last] row index; null for every row) are the rows whose own field is
 * checked, and `cells(cell, rowIndex)` says which cells' fields are (by default, every cell of
 * those rows). With no scope, the whole table.
 */
export function lockedParts(tbl, { rows = null, cells = null } = {}) {
  const inRows = (r) => rows == null || (r >= rows[0] && r <= rows[1]);
  const cellTest = cells || ((c, r) => inRows(r));
  const wrappers = new Set();
  rowsOf(tbl).forEach((tr, r) => {
    if (inRows(r)) for (const w of wrappersOf(tr)) wrappers.add(w);
    for (const c of layout(tr)) if (cellTest(c, r)) for (const w of wrappersOf(c.tc)) wrappers.add(w);
  });
  for (const w of wrappers) {
    if (!isW(w, 'sdt')) continue;
    const sdtPr = wChild(w, 'sdtPr');
    const lock = wAttr(wChild(sdtPr, 'lock'), 'val');
    const name = wAttr(wChild(sdtPr, 'alias'), 'val');
    const field = name ? `the field “${name}”` : 'a field';
    if (lock === 'sdtContentLocked' || lock === 'contentLocked') return `This would change ${field}, which is locked: change it in Word`;
    if (wChild(sdtPr, 'dataBinding')) return `This would change ${field}, which is linked to the document’s data: change it in Word`;
  }
  return null;
}

/** True if a cell sits inside a content control of its own (not just one around its row). */
export const isWrappedCell = (tc) => wrappersOf(tc).length > 0;

/** The repeating section item (Word's repeatable content control) holding just this row, if any. */
function repeatingItem(tr) {
  const sdt = wrappersOf(tr).find((w) => isW(w, 'sdt'));
  const sdtPr = sdt && wChild(sdt, 'sdtPr');
  if (!sdtPr || ![...sdtPr.children].some((c) => c.namespaceURI === W15 && c.localName === 'repeatingSectionItem')) return null;
  return rowsOf(wChild(sdt, 'sdtContent')).length === 1 ? sdt : null;
}

/** Content controls copied with a row get new ids, and show no placeholder (their copies are empty). */
function freshControls(el) {
  const sdts = isW(el, 'sdt') ? [el, ...el.getElementsByTagNameNS(W, 'sdt')] : [...el.getElementsByTagNameNS(W, 'sdt')];
  for (const sdt of sdts) {
    const sdtPr = wChild(sdt, 'sdtPr');
    const id = wChild(sdtPr, 'id');
    if (id) id.setAttributeNS(W, 'w:val', String(Math.floor(Math.random() * 0x7fffffff) - 0x3fffffff));
    removeW(sdtPr, 'showingPlcHdr');
  }
}

const int = (el, attr = 'val') => {
  const v = parseInt(wAttr(el, attr), 10);
  return Number.isFinite(v) ? v : 0;
};

export function span(tc) {
  return Math.max(1, int(wChild(wChild(tc, 'tcPr'), 'gridSpan')) || 1);
}

function rowValue(tr, local) {
  return int(wChild(wChild(tr, 'trPr'), local));
}

/** A row's cells with the grid columns they cover: [{ tc, start, span }]. */
export function layout(tr) {
  let g = rowValue(tr, 'gridBefore');
  return cellsOf(tr).map((tc) => {
    const s = span(tc);
    const cell = { tc, start: g, span: s };
    g += s;
    return cell;
  });
}

/** The grid columns a rendered cell covers: { row, start, span } (row and cell found by position). */
export function cellPosition(tbl, rowIndex, cellIndex) {
  const tr = rowsOf(tbl)[rowIndex];
  const cell = tr && layout(tr)[cellIndex];
  return cell ? { row: rowIndex, start: cell.start, span: cell.span } : null;
}

export function gridWidths(tbl) {
  return wChildren(wChild(tbl, 'tblGrid'), 'gridCol').map((c) => int(c, 'w'));
}

function vMerge(tc) {
  const v = wChild(wChild(tc, 'tcPr'), 'vMerge');
  if (!v) return null;
  return wAttr(v, 'val') === 'restart' ? 'restart' : 'continue';
}

function tcPr(tc) {
  return wChild(tc, 'tcPr') || tc.insertBefore(createW(tc.ownerDocument, 'tcPr'), tc.firstChild);
}

function setVMerge(tc, kind) {
  if (!kind) {
    removeW(wChild(tc, 'tcPr'), 'vMerge');
    return;
  }
  const v = ensureOrdered(tcPr(tc), 'vMerge', TCPR_ORDER);
  if (kind === 'restart') v.setAttributeNS(W, 'w:val', 'restart');
  else v.removeAttributeNS(W, 'val');
}

function setSpan(tc, n) {
  if (n > 1) ensureOrdered(tcPr(tc), 'gridSpan', TCPR_ORDER).setAttributeNS(W, 'w:val', String(n));
  else removeW(wChild(tc, 'tcPr'), 'gridSpan');
}

function trPr(tr) {
  const existing = wChild(tr, 'trPr');
  if (existing) return existing;
  const before = elementChildren(tr).find((c) => !(c.namespaceURI === W && c.localName === 'tblPrEx'));
  return tr.insertBefore(createW(tr.ownerDocument, 'trPr'), before || null);
}

/** Sets gridBefore / gridAfter (removed at 0), before the row's revision marks. */
function setRowValue(tr, local, n) {
  if (n <= 0) {
    const props = wChild(tr, 'trPr');
    removeW(props, local);
    if (props && !elementChildren(props).length) tr.removeChild(props);
    return;
  }
  const props = trPr(tr);
  let el = wChild(props, local);
  if (!el) {
    el = createW(tr.ownerDocument, local);
    props.insertBefore(el, wChildren(props).find((c) => ['ins', 'del', 'trPrChange'].includes(c.localName)) || null);
  }
  el.setAttributeNS(W, 'w:val', String(n));
}

/** Adds a row revision (<w:ins> or <w:del> made by the caller) to a row's properties. */
export function markRow(tr, rev) {
  const props = trPr(tr);
  props.insertBefore(rev, rev.localName === 'ins' ? wChild(props, 'del') || wChild(props, 'trPrChange') : wChild(props, 'trPrChange'));
}

/** An empty copy of a cell: its properties (not merges or revisions) and its first paragraph's formatting. */
function blankCell(tc) {
  const copy = tc.cloneNode(true);
  const props = wChild(copy, 'tcPr');
  removeW(props, 'gridSpan', 'hMerge', 'vMerge', 'cellIns', 'cellDel', 'cellMerge', 'tcPrChange');
  emptyCell(copy, tc);
  return copy;
}

/** Leaves only a cell's properties and an empty paragraph formatted like the first one of `like`. */
function emptyCell(tc, like = tc) {
  const firstPPr = wChild(wChild(like, 'p'), 'pPr');
  for (const c of elementChildren(tc)) if (!(c.namespaceURI === W && c.localName === 'tcPr')) tc.removeChild(c);
  const p = tc.appendChild(createW(tc.ownerDocument, 'p'));
  if (firstPPr) {
    const pPr = p.appendChild(firstPPr.cloneNode(true));
    removeW(pPr, 'pPrChange');
    const rPr = wChild(pPr, 'rPr');
    removeW(rPr, ...REVISION_MARKS, 'rPrChange');
    if (rPr && !elementChildren(rPr).length && !rPr.attributes.length) pPr.removeChild(rPr);
  }
}

/** An empty copy of a row: its cells (spans kept, merges and revisions not), no text, bookmarks or the like. */
function blankRow(tr) {
  const copy = tr.cloneNode(true);
  for (const c of elementChildren(copy)) {
    const cellHolder = c.namespaceURI === W && (c.localName === 'sdt' || c.localName === 'customXml') && cellsOf(c.localName === 'sdt' ? wChild(c, 'sdtContent') || c : c).length;
    if (!(c.namespaceURI === W && ['trPr', 'tblPrEx', 'tc'].includes(c.localName)) && !cellHolder) copy.removeChild(c);
  }
  freshControls(copy);
  const props = wChild(copy, 'trPr');
  if (props) {
    removeW(props, 'ins', 'del', 'trPrChange');
    if (!elementChildren(props).length) copy.removeChild(props);
  }
  cellsOf(copy).forEach((tc, i) => {
    removeW(wChild(tc, 'tcPr'), 'vMerge', 'hMerge', 'cellIns', 'cellDel', 'cellMerge', 'tcPrChange');
    emptyCell(tc, cellsOf(tr)[i]);
  });
  return copy;
}

const isHeader = (tr) => !!wChild(wChild(tr, 'trPr'), 'tblHeader');

/**
 * Inserts `count` rows above or below row `index`, copies of it (as Word does): same cells, empty.
 * A header row's repeat goes with them only inside the header rows; a vertical merge carries on
 * through new rows inserted inside it. `revision()` (if given) makes each a tracked insertion.
 * Returns the new rows.
 */
export function insertRows(tbl, index, count, below, revision = null) {
  const rows = rowsOf(tbl);
  const ref = rows[index];
  const next = below ? rows[index + 1] : ref;
  const prev = below ? ref : rows[index - 1];
  const made = [];
  for (let k = 0; k < count; k++) {
    const tr = blankRow(ref);
    if (isHeader(ref) && !(below ? next && isHeader(next) : true)) {
      const props = wChild(tr, 'trPr');
      removeW(props, 'tblHeader');
      if (!elementChildren(props).length) tr.removeChild(props);
    }
    // Inside a vertical merge (the cells above and below the new row are merged): merged too.
    for (const cell of layout(tr)) {
      const above = prev && layout(prev).find((c) => c.start === cell.start && c.span === cell.span);
      const belowCell = next && layout(next).find((c) => c.start === cell.start && c.span === cell.span);
      const inside = above && vMerge(above.tc) && belowCell && vMerge(belowCell.tc) === 'continue';
      if (inside) setVMerge(cell.tc, 'continue');
    }
    if (revision) markRow(tr, revision());
    made.push(tr);
  }
  // In a repeating section (Word's repeatable rows), each new row is a new item, as Word's + adds one.
  const item = repeatingItem(ref);
  const placed = item ? made.map((tr) => {
    const copy = item.cloneNode(true);
    freshControls(copy);
    wChild(copy, 'sdtContent').replaceChildren(tr);
    return copy;
  }) : made;
  const at = item || ref;
  if (below) at.after(...placed);
  else at.before(...placed);
  return made;
}

/**
 * Inserts `count` columns at grid boundary `at` (0 is before the first column), each as wide as
 * column `like`. A cell spanning the boundary is widened; elsewhere new cells copy the cell next to
 * them (on the `like` side).
 */
export function insertColumns(tbl, at, count, like) {
  const made = [];
  for (const tr of rowsOf(tbl)) {
    const cells = layout(tr);
    const before = rowValue(tr, 'gridBefore');
    const end = cells.length ? cells[cells.length - 1].start + cells[cells.length - 1].span : before;
    if (at < before || (!cells.length && at <= before)) {
      setRowValue(tr, 'gridBefore', before + count);
      continue;
    }
    if (at > end) {
      setRowValue(tr, 'gridAfter', rowValue(tr, 'gridAfter') + count);
      continue;
    }
    const across = cells.find((c) => c.start < at && at < c.start + c.span);
    if (across) {
      setSpan(across.tc, across.span + count);
      continue;
    }
    const after = cells.find((c) => c.start === at); // the cell the new ones go before
    const left = cells.find((c) => c.start + c.span === at);
    const model = (like < at ? left || after : after || left).tc;
    const fresh = Array.from({ length: count }, () => blankCell(model));
    made.push(...fresh);
    if (after) outermost(after.tc).before(...fresh);
    else outermost(left.tc).after(...fresh);
  }
  const grid = wChild(tbl, 'tblGrid') || tbl.insertBefore(createW(tbl.ownerDocument, 'tblGrid'), rowsOf(tbl)[0] || null);
  const cols = wChildren(grid, 'gridCol');
  const width = int(cols[like], 'w') || int(cols[cols.length - 1], 'w') || 1440;
  const fresh = Array.from({ length: count }, () => createW(tbl.ownerDocument, 'gridCol', { w: width }));
  if (cols[at]) cols[at].before(...fresh);
  else grid.append(...fresh);
  return made; // the new cells (not cells widened across the place)
}

/** Deletes grid columns `from` to `to` (inclusive): cells in them go, cells spanning them narrow. */
export function deleteColumns(tbl, from, to) {
  const overlap = (s, e) => Math.max(0, Math.min(e, to + 1) - Math.max(s, from));
  for (const tr of rowsOf(tbl)) {
    const before = rowValue(tr, 'gridBefore');
    const cells = layout(tr);
    const end = cells.length ? cells[cells.length - 1].start + cells[cells.length - 1].span : before;
    const after = rowValue(tr, 'gridAfter');
    for (const c of cells) {
      const gone = overlap(c.start, c.start + c.span);
      if (gone === c.span) detach(c.tc);
      else if (gone) setSpan(c.tc, c.span - gone);
    }
    if (before) setRowValue(tr, 'gridBefore', before - overlap(0, before));
    if (after) setRowValue(tr, 'gridAfter', after - overlap(end, end + after));
    if (!cellsOf(tr).length) detach(tr);
  }
  const cols = wChildren(wChild(tbl, 'tblGrid'), 'gridCol');
  cols.slice(from, to + 1).forEach((c) => c.parentNode.removeChild(c));
}

/** Deletes rows (by index), as Word does: a vertical merge starting in one starts below it instead. */
export function deleteRows(tbl, indices) {
  const rows = rowsOf(tbl);
  for (const i of indices) if (rows[i]) detach(rows[i]);
}

/**
 * Tidies merges after rows or columns changed: a cell continuing a vertical merge with nothing
 * merged above it starts one.
 */
export function fixMerges(tbl) {
  let prev = null;
  for (const tr of rowsOf(tbl)) {
    for (const cell of layout(tr)) {
      if (vMerge(cell.tc) !== 'continue') continue;
      const above = prev && layout(prev).find((c) => c.start === cell.start && c.span === cell.span);
      if (!above || !vMerge(above.tc)) setVMerge(cell.tc, 'restart');
    }
    prev = tr;
  }
}

/**
 * Widths after columns changed: if the table fitted between the margins (`textWidth`) before and
 * doesn't now, every column narrows in proportion, as Word fits it; cell widths given in twips
 * follow the grid, and so does a table width given in twips.
 */
export function fitWidths(tbl, { before, textWidth }) {
  const grid = wChildren(wChild(tbl, 'tblGrid'), 'gridCol');
  let widths = grid.map((c) => int(c, 'w'));
  const total = widths.reduce((a, b) => a + b, 0);
  if (textWidth && total > textWidth && before <= textWidth + 15) {
    widths = widths.map((w) => Math.max(1, Math.floor((w * textWidth) / total)));
    grid.forEach((c, i) => c.setAttributeNS(W, 'w:w', String(widths[i])));
  }
  for (const tr of rowsOf(tbl)) {
    for (const cell of layout(tr)) {
      const tcW = wChild(wChild(cell.tc, 'tcPr'), 'tcW');
      if (!tcW || wAttr(tcW, 'type') !== 'dxa') continue;
      const w = widths.slice(cell.start, cell.start + cell.span).reduce((a, b) => a + b, 0);
      if (w) tcW.setAttributeNS(W, 'w:w', String(w));
    }
  }
  const tblW = wChild(wChild(tbl, 'tblPr'), 'tblW');
  if (tblW && wAttr(tblW, 'type') === 'dxa') tblW.setAttributeNS(W, 'w:w', String(widths.reduce((a, b) => a + b, 0)));
}

export const totalWidth = (tbl) => gridWidths(tbl).reduce((a, b) => a + b, 0);

/**
 * A row deleted with tracking, as Word writes it: the row's revision (`rev('del')`) and its text
 * in <w:del> (inside an insertion, the deletion goes in it).
 */
export function deleteRowTracked(tr, rev) {
  markRow(tr, rev('del'));
  deleteRunsTracked(tr, rev);
}

/** Puts the text in `el` (a row or cell) in <w:del> made by `rev('del')`, as a tracked deletion. */
function deleteRunsTracked(el, rev) {
  const tr = el;
  for (const r of [...tr.getElementsByTagNameNS(W, 'r')]) {
    let inside = null;
    for (let n = r.parentNode; n && n !== tr; n = n.parentNode) {
      if (n.namespaceURI === W && ['del', 'moveFrom', 'r'].includes(n.localName)) { inside = n; break; }
    }
    if (inside) continue; // already deleted, or a run inside a run's content (a text box)
    const del = rev('del');
    r.before(del);
    del.appendChild(r);
    for (const [from, to] of [['t', 'delText'], ['instrText', 'delInstrText']]) {
      for (const t of [...r.getElementsByTagNameNS(W, from)]) {
        const d = createW(tr.ownerDocument, to);
        for (const a of [...t.attributes]) d.setAttributeNS(a.namespaceURI, a.name, a.value);
        d.textContent = t.textContent;
        t.replaceWith(d);
      }
    }
  }
}

/** Rejecting a deleted row (or cell) brings its text back: its deletions are undone. */
export function undeleteRow(tr) {
  for (const del of [...tr.getElementsByTagNameNS(W, 'del')]) {
    if (del.parentNode?.namespaceURI === W && ['trPr', 'tcPr'].includes(del.parentNode.localName)) continue;
    if (del.parentNode?.namespaceURI === W && del.parentNode.localName === 'rPr') continue; // a paragraph mark
    for (const [from, to] of [['delText', 't'], ['delInstrText', 'instrText']]) {
      for (const t of [...del.getElementsByTagNameNS(W, from)]) {
        const d = createW(tr.ownerDocument, to);
        for (const a of [...t.attributes]) d.setAttributeNS(a.namespaceURI, a.name, a.value);
        d.textContent = t.textContent;
        t.replaceWith(d);
      }
    }
    del.replaceWith(...del.childNodes);
  }
  for (const rPr of [...tr.getElementsByTagNameNS(W, 'rPr')]) {
    if (rPr.parentNode?.localName === 'pPr') removeW(rPr, 'del');
  }
}

/** A new table's XML: `rows` × `cols` empty cells across `textWidth` twips, in table style `styleId`. */
export function newTableXml(doc, rows, cols, textWidth, styleId, revision = null, weights = null) {
  const tbl = createW(doc, 'tbl');
  const tblPr = tbl.appendChild(createW(doc, 'tblPr'));
  if (styleId) tblPr.appendChild(createW(doc, 'tblStyle', { val: styleId }));
  tblPr.appendChild(createW(doc, 'tblW', { w: 0, type: 'auto' }));
  tblPr.appendChild(createW(doc, 'tblLook', {
    val: '04A0', firstRow: 1, lastRow: 0, firstColumn: 1, lastColumn: 0, noHBand: 0, noVBand: 1,
  }));
  // Equal columns, or in proportion to `weights` (relative widths, as a pasted table gives them).
  const total = weights?.length === cols ? weights.reduce((a, b) => a + b, 0) : 0;
  const widths = Array.from({ length: cols }, (_, c) => Math.max(1, Math.floor(total ? (textWidth * weights[c]) / total : textWidth / cols)));
  const grid = tbl.appendChild(createW(doc, 'tblGrid'));
  for (let c = 0; c < cols; c++) grid.appendChild(createW(doc, 'gridCol', { w: widths[c] }));
  for (let r = 0; r < rows; r++) {
    const tr = tbl.appendChild(createW(doc, 'tr'));
    if (revision) markRow(tr, revision());
    for (let c = 0; c < cols; c++) {
      const tc = tr.appendChild(createW(doc, 'tc'));
      tc.appendChild(createW(doc, 'tcPr')).appendChild(createW(doc, 'tcW', { w: widths[c], type: 'dxa' }));
      tc.appendChild(createW(doc, 'p'));
    }
  }
  return tbl;
}

// ---------------------------------------------------------------- merging and splitting cells

const cellsWithin = (tr, from, to) => layout(tr).filter((c) => c.start >= from && c.start + c.span - 1 <= to);

/**
 * The cells to merge for a selection of rows `r0`..`r1` and grid columns `c0`..`c1`, grown (as Word
 * does) until no cell or vertical merge sticks out of it: { r0, r1, c0, c1 }, or null if it isn't a
 * rectangle of whole cells (a row ends inside it).
 */
export function mergeArea(tbl, r0, r1, c0, c1) {
  const rows = rowsOf(tbl);
  let changed = true;
  while (changed) {
    changed = false;
    for (let r = r0; r <= r1; r++) {
      for (const c of layout(rows[r])) {
        if (c.start + c.span - 1 < c0 || c.start > c1) continue;
        if (c.start < c0) { c0 = c.start; changed = true; }
        if (c.start + c.span - 1 > c1) { c1 = c.start + c.span - 1; changed = true; }
        // A vertical merge running out of the area at the top or bottom pulls the area with it.
        if (r === r0 && vMerge(c.tc) === 'continue' && r0 > 0) { r0--; changed = true; }
        const below = r === r1 && rows[r + 1] && layout(rows[r + 1]).find((d) => d.start === c.start);
        if (below && vMerge(c.tc) && vMerge(below.tc) === 'continue') { r1++; changed = true; }
      }
    }
  }
  for (let r = r0; r <= r1; r++) {
    const covered = cellsWithin(rows[r], c0, c1).reduce((n, c) => n + c.span, 0);
    if (covered !== c1 - c0 + 1) return null;
  }
  return { r0, r1, c0, c1 };
}

/** The cells of `area` (see mergeArea), row by row: [[{ tc, start, span, index }]] (index: in its row). */
export function areaCells(tbl, { r0, r1, c0, c1 }) {
  const rows = rowsOf(tbl);
  const out = [];
  for (let r = r0; r <= r1; r++) {
    const all = layout(rows[r]);
    out.push(all.map((c, index) => ({ ...c, index })).filter((c) => c.start >= c0 && c.start + c.span - 1 <= c1));
  }
  return out;
}

/**
 * Merges the cells of `area` into one, as Word does: the first cell of each row spans the area's
 * columns (the others go), and with several rows the first row's cell starts a vertical merge the
 * rows below continue. (The cells' content is moved by the caller.)
 */
export function mergeCells(tbl, area, rev = null) {
  const width = area.c1 - area.c0 + 1;
  // Across a row, tracked: the cells stay, merged with w:hMerge (which Word reads), each with a
  // tracked property change keeping its old properties, so rejecting the merge undoes it.
  const across = rev && area.c1 > area.c0;
  areaCells(tbl, area).forEach((cells, i) => {
    const [first, ...rest] = cells;
    if (across) {
      cells.forEach((c, k) => {
        const old = oldProps(c.tc);
        setHMerge(c.tc, k === 0 ? 'restart' : 'continue');
        markPropChange(c.tc, rev('tcPrChange'), old);
      });
    } else {
      for (const c of rest) detach(c.tc);
      setSpan(first.tc, width);
    }
    const was = vMerge(first.tc);
    const now = area.r1 > area.r0 ? (i === 0 ? 'restart' : 'continue') : null;
    setVMerge(first.tc, now);
    // Merging down the rows, tracked: <w:cellMerge> records each cell's new and old merge.
    if (rev && now && now !== was) markMerge(first.tc, rev('cellMerge'), now, was);
    else if (i > 0) emptyCell(first.tc); // its content went into the merged cell
  });
}

const MERGE_VALUE = { restart: 'rest', continue: 'cont' };

/** Records a tracked change of a cell's vertical merge: `rev` is the <w:cellMerge>; states are 'restart', 'continue' or null. */
export function markMerge(tc, rev, now, was) {
  rev.setAttributeNS(W, 'w:vMerge', MERGE_VALUE[now] || 'rest');
  if (was) rev.setAttributeNS(W, 'w:vMergeOrig', MERGE_VALUE[was]);
  markCell(tc, rev);
}

/** A tracked merge or split of cells down a column, resolved: accepted, the record goes; rejected, the old merge comes back. */
export function resolveMerge(tc, accept) {
  const props = wChild(tc, 'tcPr');
  const rev = wChild(props, 'cellMerge');
  if (!rev) return;
  props.removeChild(rev);
  if (accept) return;
  const orig = wAttr(rev, 'vMergeOrig');
  setVMerge(tc, orig === 'rest' ? 'restart' : orig === 'cont' ? 'continue' : null);
}

/**
 * Makes sure the grid has a column boundary at `x` twips from the table's start (within `near`
 * twips of one there already counts): the column it falls in is split in two, and cells covering
 * that column in every row span one more. Returns the boundary's index.
 */
function boundaryAt(tbl, x, near = 20) {
  const grid = wChildren(wChild(tbl, 'tblGrid'), 'gridCol');
  const widths = grid.map((c) => int(c, 'w'));
  let edge = 0;
  for (let g = 0; g < widths.length; g++) {
    if (g > 0 && Math.abs(edge - x) <= near) return g;
    if (x > edge + near && x < edge + widths[g] - near) {
      grid[g].setAttributeNS(W, 'w:w', String(x - edge));
      grid[g].after(createW(tbl.ownerDocument, 'gridCol', { w: edge + widths[g] - x }));
      for (const tr of rowsOf(tbl)) {
        const before = rowValue(tr, 'gridBefore');
        const cell = layout(tr).find((c) => c.start <= g && g < c.start + c.span);
        if (g < before) setRowValue(tr, 'gridBefore', before + 1);
        else if (cell) setSpan(cell.tc, cell.span + 1);
        else setRowValue(tr, 'gridAfter', rowValue(tr, 'gridAfter') + 1);
      }
      return g + 1;
    }
    edge += widths[g];
  }
  return Math.abs(edge - x) <= near ? widths.length : null;
}

/** The rows of the vertical merge a cell is in (just its own row if none): [first, last]. */
function mergeRows(tbl, rowIndex, start) {
  const rows = rowsOf(tbl);
  const at = (r) => layout(rows[r]).find((c) => c.start === start);
  let first = rowIndex;
  while (first > 0 && vMerge(at(first)?.tc) === 'continue' && at(first - 1) && vMerge(at(first - 1).tc)) first--;
  let last = rowIndex;
  while (rows[last + 1] && at(last + 1) && vMerge(at(last + 1).tc) === 'continue' && vMerge(at(last).tc)) last++;
  return [first, last];
}

/**
 * Splits a cell (row `rowIndex`, cell `cellIndex`) into `cols` × `rows` cells, as Word's Split Cells
 * does: columns of equal width (new grid columns where needed: cells across them in other rows
 * span them), its content in the first; new rows below it, where the row's other cells are merged
 * vertically. A cell merged across several rows splits into 1 row (each keeping the merge) or into
 * that many rows (the merge undone). Returns { reason } if it can't be split so.
 */
export function splitCell(tbl, rowIndex, cellIndex, cols, rows, rev = null) {
  const cell = layout(rowsOf(tbl)[rowIndex])[cellIndex];
  if (!cell) return { reason: 'There is no cell to split here' };
  const [first, last] = mergeRows(tbl, rowIndex, cell.start);
  const merged = last - first + 1;
  if (merged > 1 && rows !== 1 && rows !== merged) return { reason: `This cell is merged across ${merged} rows: split it into 1 or ${merged} rows` };
  const at = (r) => layout(rowsOf(tbl)[r]).find((c) => c.start === cell.start);
  const keepMerge = merged > 1 && rows === 1;
  if (merged > 1 && rows === merged) {
    for (let r = first; r <= last; r++) {
      const tc = at(r).tc;
      const was = vMerge(tc);
      if (rev && r === first) continue; // tracked, the first cell keeps starting the merge, so rejecting restores it
      setVMerge(tc, null);
      if (rev && r > first) markMerge(tc, rev('cellMerge'), null, was); // tracked: the cell no longer continues the merge
    }
  }
  if (cols > 1) {
    const widths = gridWidths(tbl);
    const x0 = widths.slice(0, cell.start).reduce((a, b) => a + b, 0);
    const span = widths.slice(cell.start, cell.start + cell.span).reduce((a, b) => a + b, 0);
    if (span / cols < 60) return { reason: 'The cell is too narrow to split into that many columns' };
    const grid = wChild(tbl, 'tblGrid');
    const oldCols = wChildren(grid, 'gridCol').map((c) => c.cloneNode(true));
    const bounds = [];
    for (let k = 1; k < cols; k++) bounds.push(boundaryAt(tbl, x0 + Math.round((k * span) / cols)));
    // Tracked, new grid columns are a tracked grid change (the old grid kept), so rejecting gives the widths back.
    if (rev && wChildren(grid, 'gridCol').length !== oldCols.length && !wChild(grid, 'tblGridChange')) {
      const change = rev('tblGridChange');
      change.appendChild(createW(tbl.ownerDocument, 'tblGrid')).append(...oldCols);
      grid.appendChild(change);
    }
    const ends = [cell.start, ...bounds, null];
    for (let r = merged > 1 ? first : rowIndex; r <= (merged > 1 ? last : rowIndex); r++) {
      const c = at(r);
      ends[ends.length - 1] = c.start + c.span;
      const kind = keepMerge ? vMerge(c.tc) : null;
      setSpan(c.tc, ends[1] - ends[0]);
      let prev = outermost(c.tc);
      for (let k = 1; k < cols; k++) {
        const fresh = blankCell(c.tc);
        setSpan(fresh, ends[k + 1] - ends[k]);
        if (kind) setVMerge(fresh, kind);
        if (rev) markCell(fresh, rev('cellIns'));
        prev.after(fresh);
        prev = fresh;
      }
    }
  }
  if (rows > 1 && merged === 1) {
    const tr = rowsOf(tbl)[rowIndex];
    const from = cell.start;
    const to = cell.start + cell.span - 1;
    const inside = (c) => c.start >= from && c.start + c.span - 1 <= to;
    for (const c of layout(tr)) if (!inside(c) && !vMerge(c.tc)) setVMerge(c.tc, 'restart');
    const made = [];
    for (let k = 1; k < rows; k++) {
      const copy = blankRow(tr);
      const props = wChild(copy, 'trPr');
      removeW(props, 'tblHeader');
      if (props && !elementChildren(props).length) copy.removeChild(props);
      for (const c of layout(copy)) if (!inside(c)) setVMerge(c.tc, 'continue');
      if (rev) markRow(copy, rev('ins'));
      made.push(copy);
    }
    tr.after(...made);
  }
  return {};
}

// ---------------------------------------------------------------- tracked columns

/** Marks a cell inserted or deleted with tracking: `rev` is the <w:cellIns> or <w:cellDel>. */
export function markCell(tc, rev) {
  const props = tcPr(tc);
  removeW(props, ...(rev.localName === 'cellMerge' ? ['cellMerge'] : ['cellIns', 'cellDel']));
  const after = TCPR_ORDER.indexOf(rev.localName);
  props.insertBefore(rev, wChildren(props).find((c) => TCPR_ORDER.indexOf(c.localName) > after) || null);
}

/** A cell deleted with tracking, as Word writes it: <w:cellDel>, and its text in <w:del>. */
export function deleteCellTracked(tc, rev) {
  markCell(tc, rev('cellDel'));
  deleteRunsTracked(tc, rev);
}

/** A deleted cell rejected: its revision goes and its text comes back. */
export function undeleteCell(tc) {
  removeW(wChild(tc, 'tcPr'), 'cellDel');
  undeleteRow(tc);
}

/** An inserted cell accepted: its revision goes. */
export function acceptInsertedCell(tc) {
  removeW(wChild(tc, 'tcPr'), 'cellIns');
}

/**
 * Takes cells out of a table (an inserted column rejected, a deleted one accepted): grid columns
 * that no other cell starts in are deleted (cells spanning them narrow); a cell elsewhere goes and
 * its row ends early (gridAfter), as Word keeps such a row.
 */
export function removeCells(tbl, tcs) {
  const set = new Set(tcs);
  const columns = new Set();
  for (const tr of rowsOf(tbl)) for (const c of layout(tr)) if (set.has(c.tc)) for (let g = c.start; g < c.start + c.span; g++) columns.add(g);
  const removable = [...columns].filter((g) => rowsOf(tbl).every((tr) => {
    const c = layout(tr).find((d) => d.start <= g && g < d.start + d.span);
    return !c || set.has(c.tc) || c.span > 1;
  })).sort((a, b) => b - a);
  // Cells in removable columns go with them; any other goes on its own.
  for (const tr of rowsOf(tbl)) {
    for (const c of layout(tr)) {
      if (!set.has(c.tc)) continue;
      const inRemovable = [...Array(c.span).keys()].every((k) => removable.includes(c.start + k));
      if (!inRemovable) {
        detach(c.tc);
        setRowValue(tr, 'gridAfter', rowValue(tr, 'gridAfter') + c.span);
      }
    }
  }
  // Contiguous removable columns, from the right, so indices stay right.
  let i = 0;
  while (i < removable.length) {
    let j = i;
    while (j + 1 < removable.length && removable[j + 1] === removable[j] - 1) j++;
    deleteColumns(tbl, removable[j], removable[i]);
    i = j + 1;
  }
}

/** After cells split off with tracking are taken out again: the old grid widths back, from the tracked grid change. */
export function restoreGrid(tbl, key) {
  const grid = wChild(tbl, 'tblGrid');
  const change = wChild(grid, 'tblGridChange');
  if (!change || `${wAttr(change, 'author') || ''}|${wAttr(change, 'date') || ''}` !== key) return;
  grid.removeChild(change);
  const old = wChildren(wChild(change, 'tblGrid'), 'gridCol');
  const now = wChildren(grid, 'gridCol');
  if (old.length === now.length) now.forEach((c, i) => c.setAttributeNS(W, 'w:w', wAttr(old[i], 'w')));
}

/** The tracked grid change made with cells of change `key` (author|date) is accepted with them. */
export function dropGridChange(tbl, key) {
  const grid = wChild(tbl, 'tblGrid');
  const change = wChild(grid, 'tblGridChange');
  if (change && `${wAttr(change, 'author') || ''}|${wAttr(change, 'date') || ''}` === key) grid.removeChild(change);
}

// ---------------------------------------------------------------- merges across a row (w:hMerge)

/** A cell's horizontal merge, as older Word files (and tracked merges) mark it: 'restart', 'continue' or null. */
export function hMerge(tc) {
  const h = wChild(wChild(tc, 'tcPr'), 'hMerge');
  if (!h) return null;
  return wAttr(h, 'val') === 'restart' ? 'restart' : 'continue';
}

function setHMerge(tc, kind) {
  if (!kind) {
    removeW(wChild(tc, 'tcPr'), 'hMerge');
    return;
  }
  const h = ensureOrdered(tcPr(tc), 'hMerge', TCPR_ORDER);
  if (kind === 'restart') h.setAttributeNS(W, 'w:val', 'restart');
  else h.removeAttributeNS(W, 'val');
}

/** A copy of a cell's properties as they are, without revisions: the "old" side of a tracked change. */
function oldProps(tc) {
  const old = (wChild(tc, 'tcPr') || createW(tc.ownerDocument, 'tcPr')).cloneNode(true);
  removeW(old, 'tcPrChange', 'cellIns', 'cellDel', 'cellMerge');
  return old;
}

/** Records a tracked change of a cell's properties: `rev` (<w:tcPrChange>) holds the old ones. */
function markPropChange(tc, rev, old) {
  const props = tcPr(tc);
  removeW(props, 'tcPrChange');
  rev.appendChild(old);
  props.appendChild(rev);
}

/** True if a cell's properties have a tracked change that merges it across its row (w:hMerge). */
export function isTrackedHMerge(tc) {
  const props = wChild(tc, 'tcPr');
  return !!wChild(props, 'tcPrChange') && !!wChild(props, 'hMerge');
}

/** A tracked change of a cell's properties, resolved: accepted, its record goes; rejected, the old properties come back. */
export function resolveCellProps(tc, accept) {
  const props = wChild(tc, 'tcPr');
  const change = wChild(props, 'tcPrChange');
  if (!change) return;
  props.removeChild(change);
  if (accept) return;
  const kept = ['cellIns', 'cellDel', 'cellMerge'].map((local) => wChild(props, local)).filter(Boolean);
  const old = wChild(change, 'tcPr');
  while (props.firstChild) props.removeChild(props.firstChild);
  for (const n of [...(old?.childNodes || [])]) props.appendChild(n);
  for (const n of kept) props.appendChild(n);
}

/**
 * Cells merged across a row with w:hMerge (no change pending) become one cell spanning the columns
 * (w:gridSpan), as Word writes a merge: the first keeps its content, the others go.
 */
export function normalizeHMerge(tbl) {
  for (const tr of rowsOf(tbl)) {
    let restart = null;
    for (const c of layout(tr)) {
      const h = hMerge(c.tc);
      const pending = !!wChild(wChild(c.tc, 'tcPr'), 'tcPrChange');
      if (h === 'continue' && restart && !pending) {
        setSpan(restart.tc, span(restart.tc) + c.span);
        detach(c.tc);
      } else {
        restart = h === 'restart' && !pending ? c : null;
        if (restart) setHMerge(c.tc, null);
      }
    }
  }
}

// ---------------------------------------------------------------- cell shading and borders

const BORDER_SIDES = ['top', 'left', 'bottom', 'right'];
const TCBORDERS_ORDER = ['top', 'start', 'left', 'bottom', 'end', 'right', 'insideH', 'insideV', 'tl2br', 'tr2bl'];

/**
 * Changes the cells of `area` ({ r0, r1, c0, c1 }, in rows and grid columns) as Word's Shading and
 * Borders buttons do, in each cell's own properties (<w:tcPr>): `change(tc, edges)` changes one,
 * where `edges` says which sides of the area it is on ({ top, left, bottom, right }). Tracked
 * (`rev` makes a <w:tcPrChange>), the old properties are recorded once: changed again, the first
 * record stays, as Word keeps it.
 */
export function formatCells(tbl, area, change, rev = null) {
  areaCells(tbl, area).forEach((cells, i) => {
    const r = area.r0 + i;
    for (const c of cells) {
      const edges = { top: r === area.r0, bottom: r === area.r1, left: c.start === area.c0, right: c.start + c.span - 1 === area.c1 };
      const pending = !!wChild(wChild(c.tc, 'tcPr'), 'tcPrChange');
      const old = rev && !pending ? oldProps(c.tc) : null;
      change(c.tc, edges);
      if (old) markPropChange(c.tc, rev('tcPrChange'), old);
    }
  });
}

/** A cell's shading: a fill colour (RRGGBB), as Word writes it (<w:shd w:val="clear" w:color="auto" w:fill="…"/>), or none. */
export function setShading(tc, hex) {
  if (!hex) {
    removeW(wChild(tc, 'tcPr'), 'shd');
    return;
  }
  const shd = ensureOrdered(tcPr(tc), 'shd', TCPR_ORDER);
  for (const a of [...shd.attributes]) shd.removeAttributeNode(a);
  shd.setAttributeNS(W, 'w:val', 'clear');
  shd.setAttributeNS(W, 'w:color', 'auto');
  shd.setAttributeNS(W, 'w:fill', hex);
}

/**
 * Sets sides of a cell's borders (<w:tcBorders>): `sides` maps 'top', 'left', 'bottom', 'right' to
 * true (a single line, ½ pt, automatic colour: Word's default border) or false (no border, "nil",
 * which hides a table style's border too). Sides not named stay as they are.
 */
export function setBorders(tc, sides) {
  const props = tcPr(tc);
  const borders = ensureOrdered(props, 'tcBorders', TCPR_ORDER);
  for (const side of BORDER_SIDES) {
    if (!(side in sides)) continue;
    // Word's logical sides (start, end) are the same edges: one of each is kept.
    removeW(borders, side === 'left' ? 'start' : side === 'right' ? 'end' : side);
    const b = ensureOrdered(borders, side, TCBORDERS_ORDER);
    for (const a of [...b.attributes]) b.removeAttributeNode(a);
    if (sides[side]) {
      b.setAttributeNS(W, 'w:val', 'single');
      b.setAttributeNS(W, 'w:sz', '4');
      b.setAttributeNS(W, 'w:space', '0');
      b.setAttributeNS(W, 'w:color', 'auto');
    } else {
      b.setAttributeNS(W, 'w:val', 'nil');
    }
  }
}

/** A cell's own borders (<w:tcBorders>) by side: { top, left, bottom, right } -> its border element, if set. */
export function cellBorders(tc) {
  const borders = wChild(wChild(tc, 'tcPr'), 'tcBorders');
  if (!borders) return {};
  return {
    top: wChild(borders, 'top'),
    bottom: wChild(borders, 'bottom'),
    left: wChild(borders, 'left') || wChild(borders, 'start'),
    right: wChild(borders, 'right') || wChild(borders, 'end'),
  };
}
