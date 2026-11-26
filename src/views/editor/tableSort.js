// Word's Sort for a table (Layout › Sort), split out of EditorView.vue: its dialog's state, the
// columns it offers, and the sort itself. See src/views/editor/README.md for how the editor is being
// split: each part gets the editor's shared state as `ctx`, read when it is used (never copied at
// setup, as some of it, such as the undo stack, is set later), and returns what the template uses.
import { reactive, ref, nextTick } from 'vue';
import { t } from '../../i18n.js';
import { parseCellNumber, separators } from '../../docx/formulas.js';
import { readOnlyAt } from '../../editor/readonly.js';

/**
 * `ctx`: { doc, page, message, tracking, tableMenu, layoutButton, undoStack, tableSelection(),
 * anchorRect(), keepInView(), refocus(), explainReadOnly(), onInput(), markDirty() }, as the editor
 * has them. Returns { sortDialog, sortDialogEl, openSortDialog, closeSortDialog, refreshSortColumns, sortTable }.
 */
export function useTableSort(ctx) {
  // Word's Sort for a table: the rows below the heading row, by one column's text (as the document's
  // language orders it), number (as it writes them) or date; empty cells last. One undoable step.
  const sortDialog = reactive({ open: false, x: 0, y: 0, col: 0, type: 'text', desc: false, header: true, columns: [] });
  const sortDialogEl = ref(null);
  let sortTarget = null;

  /** The cell of grid column `col` in row `tr` (a merged cell for each column it spans), or null. */
  function cellInColumn(tr, col) {
    for (const td of tr.cells) {
      const pos = ctx.doc.cellPosition(td);
      if (pos && col >= pos.start && col < pos.start + pos.span) return td;
    }
    return null;
  }

  function refreshSortColumns() {
    const rows = sortTarget ? ctx.doc.tableRows(sortTarget.table) : [];
    const count = Math.max(0, ...rows.map((r) => [...r.cells].reduce((n, td) => n + (ctx.doc.cellPosition(td)?.span || 1), 0)));
    sortDialog.columns = Array.from({ length: count }, (_, c) => {
      const head = sortDialog.header && rows[0] ? cellInColumn(rows[0], c)?.textContent.trim() : '';
      return { value: c, label: head ? head.slice(0, 40) : t('Column {n}', { n: c + 1 }) };
    });
    if (sortDialog.col >= count) sortDialog.col = 0;
  }

  function openSortDialog() {
    ctx.tableMenu.open = false;
    const sel = ctx.tableSelection();
    if (!sel) {
      ctx.message.value = t('Click in a table first');
      return;
    }
    sortTarget = { table: sel.table };
    const rect = ctx.anchorRect(ctx.layoutButton.value);
    const rows = ctx.doc.tableRows(sel.table);
    // A heading row: a row Word repeats on each page, or a first row that doesn't hold numbers where the next does.
    const pos = ctx.doc.cellPosition(sel.cell);
    Object.assign(sortDialog, { open: true, x: Math.max(8, Math.min(rect.left, window.innerWidth - 300)), y: rect.bottom + 4, col: pos?.start ?? 0, type: 'text', desc: false, header: rows.length > 1 });
    refreshSortColumns();
    // Numbers in the column: sort them as numbers.
    const seps = separators(ctx.doc.styles.language() || navigator.language);
    const values = rows.slice(sortDialog.header ? 1 : 0).map((r) => cellInColumn(r, sortDialog.col)?.textContent.trim()).filter(Boolean);
    if (values.length && values.every((v) => parseCellNumber(v, seps) != null)) sortDialog.type = 'number';
    nextTick(() => {
      ctx.keepInView(sortDialog, sortDialogEl.value);
      sortDialogEl.value?.querySelector('select')?.focus();
    });
  }

  function closeSortDialog() {
    sortDialog.open = false;
    ctx.refocus();
  }

  /** A date written in a cell ("17/08/1945", "17 Agustus 1945", "2026-11-20"), as a time, or null. */
  function cellDate(text) {
    const s = text.trim();
    let m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s);
    if (m) return Date.UTC(+m[1], +m[2] - 1, +m[3]);
    m = /^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})/.exec(s);
    if (m) {
      const lang = ctx.doc.styles.language() || navigator.language || 'en';
      const us = /^en-US$/i.test(lang);
      const [d, mo] = us ? [+m[2], +m[1]] : [+m[1], +m[2]];
      const y = +m[3] < 100 ? 2000 + +m[3] : +m[3];
      return Date.UTC(y, mo - 1, d);
    }
    // Month names in the document's language or English.
    const lang = ctx.doc.styles.language() || navigator.language || 'en';
    for (const l of [lang, 'en']) {
      for (let mo = 0; mo < 12; mo++) {
        for (const width of ['long', 'short']) {
          const name = new Intl.DateTimeFormat(l, { month: width }).format(new Date(2001, mo, 1)).replace('.', '').toLowerCase();
          const re = new RegExp(`(\\d{1,2})\\s+${name}\\.?\\s+(\\d{4})|${name}\\.?\\s+(\\d{1,2}),?\\s+(\\d{4})`, 'i');
          m = re.exec(s);
          if (m) return m[1] ? Date.UTC(+m[2], mo, +m[1]) : Date.UTC(+m[4], mo, +m[3]);
        }
      }
    }
    const t2 = Date.parse(s);
    return Number.isNaN(t2) ? null : t2;
  }

  function sortTable() {
    const d = { ...sortDialog };
    sortDialog.open = false;
    ctx.refocus();
    const table = sortTarget?.table;
    if (!table?.isConnected) return;
    const why = ctx.doc.tableLocked(table);
    if (why) {
      ctx.message.value = why;
      return;
    }
    const locked = readOnlyAt(table, ctx.page.value);
    if (locked) return ctx.explainReadOnly(locked, true);
    if (ctx.tracking.value) {
      ctx.message.value = t('Sorting isn’t tracked: turn Track Changes off to sort the table, or move the rows by hand');
      return;
    }
    const rows = ctx.doc.tableRows(table);
    const body = rows.filter((r, i) => !(d.header && i === 0));
    if (body.some((r) => r.dataset.revRow || [...r.cells].some((td) => td.classList.contains('vmerge') || td.dataset.revCell || td.dataset.revMerge))) {
      ctx.message.value = t('Rows with cells merged down, or with tracked changes to their rows or cells, can’t be sorted: split those cells, or accept or reject the changes first');
      return;
    }
    if (body.length < 2) return;
    const lang = ctx.doc.styles.language() || navigator.language || 'en';
    const seps = separators(lang);
    const collator = new Intl.Collator(lang, { numeric: true, sensitivity: 'base' });
    const keyed = body.map((tr, i) => {
      const text = cellInColumn(tr, d.col)?.textContent.trim() || '';
      const key = d.type === 'number' ? parseCellNumber(text, seps) : d.type === 'date' ? cellDate(text) : text || null;
      return { tr, i, key };
    });
    const order = [...keyed].sort((a, b) => {
      if (a.key == null || b.key == null) return a.key == null && b.key == null ? a.i - b.i : a.key == null ? 1 : -1; // empty last
      const c = d.type === 'text' ? collator.compare(a.key, b.key) : a.key - b.key;
      return (d.desc ? -c : c) || a.i - b.i;
    });
    if (order.every((o, i) => o.i === i)) {
      ctx.message.value = t('The rows are already in that order');
      return;
    }
    ctx.undoStack.flush();
    // Moved as they are, in their new order after the heading row (rows are kept whole, with their cells' properties).
    const after = body[body.length - 1].nextElementSibling;
    const parent = body[0].parentElement;
    for (const { tr } of order) parent.insertBefore(tr, after);
    ctx.doc.renumber(ctx.page.value);
    ctx.onInput();
    ctx.undoStack.commit();
    ctx.markDirty();
    ctx.message.value = t('Sorted {n} rows by {column}, {order}', { n: body.length, column: d.columns[d.col]?.label || '', order: t(d.desc ? 'descending' : 'ascending') });
  }

  return { sortDialog, sortDialogEl, openSortDialog, closeSortDialog, refreshSortColumns, sortTable };
}
