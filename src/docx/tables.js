// Changing a table's structure as Word does: rows inserted above or below, columns inserted left
// or right, rows and columns deleted. Everything here works on a table's XML (<w:tbl>), in grid
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

export const rowsOf = (tbl) => wChildren(tbl, 'tr');
export const cellsOf = (tr) => wChildren(tr, 'tc');

/** True if content controls or custom XML wrap rows or cells: such tables are changed in Word. */
export function hasWrappedParts(tbl) {
  const wrapped = (el) => elementChildren(el).some((c) => c.namespaceURI === W && (c.localName === 'sdt' || c.localName === 'customXml'));
  return wrapped(tbl) || rowsOf(tbl).some(wrapped);
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
    if (!(c.namespaceURI === W && ['trPr', 'tblPrEx', 'tc'].includes(c.localName))) copy.removeChild(c);
  }
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
  if (below) ref.after(...made);
  else ref.before(...made);
  return made;
}

/**
 * Inserts `count` columns at grid boundary `at` (0 is before the first column), each as wide as
 * column `like`. A cell spanning the boundary is widened; elsewhere new cells copy the cell next to
 * them (on the `like` side).
 */
export function insertColumns(tbl, at, count, like) {
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
    if (after) after.tc.before(...fresh);
    else left.tc.after(...fresh);
  }
  const grid = wChild(tbl, 'tblGrid') || tbl.insertBefore(createW(tbl.ownerDocument, 'tblGrid'), rowsOf(tbl)[0] || null);
  const cols = wChildren(grid, 'gridCol');
  const width = int(cols[like], 'w') || int(cols[cols.length - 1], 'w') || 1440;
  const fresh = Array.from({ length: count }, () => createW(tbl.ownerDocument, 'gridCol', { w: width }));
  if (cols[at]) cols[at].before(...fresh);
  else grid.append(...fresh);
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
      if (gone === c.span) tr.removeChild(c.tc);
      else if (gone) setSpan(c.tc, c.span - gone);
    }
    if (before) setRowValue(tr, 'gridBefore', before - overlap(0, before));
    if (after) setRowValue(tr, 'gridAfter', after - overlap(end, end + after));
    if (!cellsOf(tr).length) tbl.removeChild(tr);
  }
  const cols = wChildren(wChild(tbl, 'tblGrid'), 'gridCol');
  cols.slice(from, to + 1).forEach((c) => c.parentNode.removeChild(c));
}

/** Deletes rows (by index), as Word does: a vertical merge starting in one starts below it instead. */
export function deleteRows(tbl, indices) {
  const rows = rowsOf(tbl);
  for (const i of indices) if (rows[i]) tbl.removeChild(rows[i]);
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

/** Rejecting a deleted row brings its text back: its deletions are undone. */
export function undeleteRow(tr) {
  for (const del of [...tr.getElementsByTagNameNS(W, 'del')]) {
    if (del.parentNode?.namespaceURI === W && del.parentNode.localName === 'trPr') continue;
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
export function newTableXml(doc, rows, cols, textWidth, styleId, revision = null) {
  const tbl = createW(doc, 'tbl');
  const tblPr = tbl.appendChild(createW(doc, 'tblPr'));
  if (styleId) tblPr.appendChild(createW(doc, 'tblStyle', { val: styleId }));
  tblPr.appendChild(createW(doc, 'tblW', { w: 0, type: 'auto' }));
  tblPr.appendChild(createW(doc, 'tblLook', {
    val: '04A0', firstRow: 1, lastRow: 0, firstColumn: 1, lastColumn: 0, noHBand: 0, noVBand: 1,
  }));
  const width = Math.floor(textWidth / cols);
  const grid = tbl.appendChild(createW(doc, 'tblGrid'));
  for (let c = 0; c < cols; c++) grid.appendChild(createW(doc, 'gridCol', { w: width }));
  for (let r = 0; r < rows; r++) {
    const tr = tbl.appendChild(createW(doc, 'tr'));
    if (revision) markRow(tr, revision());
    for (let c = 0; c < cols; c++) {
      const tc = tr.appendChild(createW(doc, 'tc'));
      tc.appendChild(createW(doc, 'tcPr')).appendChild(createW(doc, 'tcW', { w: width, type: 'dxa' }));
      tc.appendChild(createW(doc, 'p'));
    }
  }
  return tbl;
}
