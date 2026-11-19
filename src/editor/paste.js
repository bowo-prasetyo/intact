// Paste with basic formatting. Clipboard HTML (from web pages, Google Docs or Word) is reduced to
// paragraphs, list items, line breaks and runs that are bold, italic, underlined, struck through,
// superscript or subscript, and links to web and e-mail addresses (written as Word's hyperlinks).
// Tables (from Excel, Google Sheets, LibreOffice Calc, Word or a web page) stay tables: their rows,
// columns, merged cells and relative column widths, each cell's text with the formatting above.
// Everything else (fonts, colours, styles, pictures, other links, data-* attributes) is dropped,
// because it can't be mapped faithfully to the document's own formatting.
import { isLinkable } from '../docx/links.js';

const BLOCKS = new Set(['P', 'DIV', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'LI', 'BLOCKQUOTE', 'PRE', 'TR', 'TABLE', 'UL', 'OL', 'SECTION', 'ARTICLE', 'HEADER', 'FOOTER', 'DL', 'DT', 'DD', 'FIGURE']);
const SKIP = new Set(['SCRIPT', 'STYLE', 'TEMPLATE', 'HEAD', 'TITLE', 'META', 'LINK', 'NOSCRIPT', 'IFRAME', 'OBJECT', 'EMBED', 'SVG', 'CANVAS', 'VIDEO', 'AUDIO', 'BUTTON', 'SELECT', 'TEXTAREA', 'INPUT']);

// As many columns as Word allows in a table, and a limit on rows; a bigger table is pasted as text.
export const TABLE_MAX_COLS = 63;
export const TABLE_MAX_ROWS = 1000;

/**
 * Parses clipboard data into { blocks, skippedImages }. A block is
 * { list: null | { kind: 'bullet' | 'number', level }, runs: [{ text, b, i, u, s, va, href } | { br: true }] }
 * or a table: { table: { rows, cols, widths, cells: [{ r, c, rowspan, colspan, blocks }] } }, where
 * `widths` are the columns' relative widths (or null) and each cell's `blocks` are paragraph blocks.
 */
export function parseClipboard(html, text) {
  if (!html) return { blocks: plainBlocks(text || ''), skippedImages: 0 };
  const doc = new DOMParser().parseFromString(html, 'text/html'); // inert: no scripts, no loading
  const state = { blocks: [], current: null, skippedImages: 0, inCell: false, classes: excelClasses(doc) };
  walk(doc.body, { b: false, i: false, u: false, s: false, pre: false }, state);
  // Whitespace between elements (Word's clipboard HTML has a line break between paragraphs) isn't
  // a paragraph of its own.
  const blocks = trimBlocks(state.blocks.filter((b) => !(b.loose && isEmpty(b))).map((b) => (b.table ? b : tidy(b))));
  if (!blocks.length && text) return { blocks: plainBlocks(text), skippedImages: state.skippedImages };
  return { blocks, skippedImages: state.skippedImages };
}

/** Drops empty blocks at the start and end (fragment wrappers produce them). */
function trimBlocks(blocks) {
  while (blocks.length && isEmpty(blocks[0])) blocks.shift();
  while (blocks.length && isEmpty(blocks[blocks.length - 1])) blocks.pop();
  return blocks;
}

/**
 * Tables as text, for where a table can't go: a row becomes a paragraph, its cells separated by
 * tabs (a cell's paragraphs joined by spaces), as pasting did before tables were kept.
 */
export function tablesAsText(blocks) {
  return blocks.flatMap((block) => {
    if (!block.table) return [block];
    const { rows, cells } = block.table;
    const out = [];
    for (let r = 0; r < rows; r++) {
      const runs = [];
      for (const cell of cells.filter((c) => c.r === r).sort((a, b) => a.c - b.c)) {
        if (runs.length) runs.push({ text: '\t', b: false, i: false, u: false, s: false });
        cell.blocks.forEach((b, i) => {
          if (i) runs.push({ text: ' ', b: false, i: false, u: false, s: false });
          runs.push(...b.runs.filter((x) => !x.br));
        });
      }
      out.push({ list: null, runs });
    }
    return out;
  });
}

/**
 * Excel's clipboard HTML formats cells by class (<td class=xl65> with ".xl65 {font-weight:700}" in
 * a <style>): those classes' declarations, by class name, so a bold heading row stays bold.
 */
function excelClasses(doc) {
  const out = new Map();
  for (const style of doc.querySelectorAll('style')) {
    for (const m of style.textContent.matchAll(/(?:^|[\s,}])(?:td)?\.(xl\d+)\s*\{([^}]*)\}/g)) out.set(m[1], m[2]);
  }
  return out;
}

function plainBlocks(text) {
  return text.replace(/\r\n?/g, '\n').split('\n').map((line) => ({ list: null, runs: line ? [{ text: line, b: false, i: false, u: false, s: false }] : [] }));
}

const isEmpty = (block) => !block.table && !block.list && !block.runs.some((r) => r.br || r.text.trim());

/** Bold, italic, underline, strikethrough and raised or lowered text from CSS declarations. */
function applyStyle(f, st) {
  const w = st.fontWeight;
  if (w) f.b = w === 'bold' || w === 'bolder' || Number(w) >= 600;
  if (st.fontStyle) f.i = st.fontStyle === 'italic' || st.fontStyle === 'oblique';
  const deco = `${st.textDecorationLine || ''} ${st.textDecoration || ''}`;
  if (/underline/.test(deco)) f.u = true;
  if (/line-through/.test(deco)) f.s = true;
  if (/\bnone\b/.test(deco) && !/underline|line-through/.test(deco)) { f.u = false; f.s = false; }
  if (st.verticalAlign === 'super') f.va = 'superscript';
  if (st.verticalAlign === 'sub') f.va = 'subscript';
  if (st.verticalAlign === 'baseline') f.va = null;
}

/** Styles from an element's tag and inline style; inline style wins (Google Docs' <b style="font-weight:normal">). */
function formatOf(el, inherited, classes = null) {
  const f = { ...inherited };
  const tag = el.tagName;
  if (tag === 'B' || tag === 'STRONG') f.b = true;
  if (tag === 'I' || tag === 'EM' || tag === 'CITE' || tag === 'DFN') f.i = true;
  if (tag === 'U' || tag === 'INS') f.u = true;
  if (tag === 'S' || tag === 'STRIKE' || tag === 'DEL') f.s = true;
  if (tag === 'SUP') f.va = 'superscript';
  if (tag === 'SUB') f.va = 'subscript';
  if (/^H[1-6]$/.test(tag) || tag === 'TH') f.b = true;
  if (tag === 'PRE') f.pre = true;
  for (const name of classes?.size ? el.classList : []) {
    const declared = classes.get(name);
    if (!declared) continue;
    const probe = document.createElement('span');
    probe.style.cssText = declared;
    applyStyle(f, probe.style);
  }
  if (el.style) applyStyle(f, el.style);
  return f;
}

const cssLength = (value) => {
  const m = /^\s*([\d.]+)\s*(px|pt|cm|mm|in)?\s*$/i.exec(value || '');
  if (!m) return null;
  const n = Number(m[1]) * ({ pt: 4 / 3, cm: 96 / 2.54, mm: 96 / 25.4, in: 96 }[(m[2] || 'px').toLowerCase()] || 1);
  return n > 0 ? n : null;
};
const widthOf = (el) => cssLength(el.style?.width) || cssLength(el.getAttribute('width'));

/**
 * A table's rows and cells, laid out on its grid as the browser lays them out (colspan, rowspan),
 * with each cell's content as paragraph blocks and the columns' relative widths where the HTML gives
 * them (<col width>, as Excel and Google Sheets write them, or the cells' widths).
 */
function parseTable(table, format, state) {
  const trs = [...table.rows].filter((tr) => !/display:\s*none/i.test(tr.getAttribute('style') || ''));
  const taken = [];
  const cells = [];
  let cols = 0;
  trs.forEach((tr, r) => {
    taken[r] ||= [];
    let c = 0;
    for (const td of tr.cells) {
      while (taken[r][c]) c++;
      const colspan = Math.max(1, Math.min(TABLE_MAX_COLS + 1, td.colSpan || 1));
      const rowspan = td.rowSpan === 0 ? trs.length - r : Math.max(1, Math.min(trs.length - r, td.rowSpan || 1));
      for (let y = r; y < r + rowspan; y++) for (let x = c; x < c + colspan; x++) (taken[y] ||= [])[x] = true;
      const sub = { blocks: [], current: null, skippedImages: 0, inCell: true, classes: state.classes };
      walk(td, formatOf(td, format, state.classes), sub);
      state.skippedImages += sub.skippedImages;
      const blocks = trimBlocks(sub.blocks.filter((b) => !(b.loose && isEmpty(b))).map(tidy));
      cells.push({ r, c, rowspan, colspan, width: colspan === 1 ? widthOf(td) : null, blocks });
      c += colspan;
    }
    cols = Math.max(cols, taken[r].length);
  });
  const rows = trs.length;
  if (!rows || !cols) return null;
  // Column widths: from <col> (with its span), else from the first cell in each column that has one.
  const widths = [];
  for (const col of table.querySelectorAll(':scope > col, :scope > colgroup > col')) {
    const w = widthOf(col);
    for (let i = 0; i < Math.max(1, col.span || 1); i++) widths.push(w);
  }
  for (let x = 0; x < cols; x++) widths[x] ||= cells.find((cell) => cell.c === x && cell.width)?.width || null;
  widths.length = cols;
  for (const cell of cells) delete cell.width;
  return { rows, cols, cells, widths: widths.every(Boolean) ? widths : null };
}

/** Word's clipboard HTML marks list paragraphs with "mso-list: l0 level2 lfo1". */
function msoList(el) {
  const m = /mso-list:\s*l\d+\s+level(\d+)/i.exec(el.getAttribute('style') || '');
  if (!m) return null;
  const marker = el.querySelector('[style*="mso-list:Ignore"], [style*="mso-list: Ignore"]')?.textContent || '';
  return { kind: /^\s*(\(?[0-9a-z]{1,4}[.)])/i.test(marker) ? 'number' : 'bullet', level: Number(m[1]) - 1 };
}

/** A list paragraph copied from Intact itself: <p class="list" data-marker="2." data-ilvl="1">. */
function intactList(el) {
  if (!el.classList.contains('list') || el.dataset.marker == null) return null;
  const level = Math.max(0, Math.min(8, Number(el.dataset.ilvl) || 0));
  return { kind: /^[•◦▪o\uf0b7\uf0a7-]$/.test(el.dataset.marker.trim()) ? 'bullet' : 'number', level };
}

function htmlList(li) {
  let level = -1;
  let kind = 'bullet';
  for (let p = li.parentElement; p; p = p.parentElement) {
    if (p.tagName === 'UL' || p.tagName === 'OL') {
      if (level === -1) kind = p.tagName === 'OL' ? 'number' : 'bullet';
      level++;
    }
  }
  return { kind, level: Math.max(0, Math.min(8, level)) };
}

/** A new paragraph block; `loose`: started by text outside any block element. */
function startBlock(state, list = null, loose = false) {
  state.current = { list, runs: [], loose };
  state.blocks.push(state.current);
}

function walk(node, format, state) {
  for (const child of node.childNodes) {
    if (child.nodeType === 3) {
      if (!state.current) startBlock(state, null, true);
      let text = child.data;
      if (!format.pre) text = text.replace(/[\t\n\r ]+/g, ' ');
      text = text.replace(/ /g, ' ');
      if (text) state.current.runs.push({ text, b: format.b, i: format.i, u: format.u, s: format.s, va: format.va || null, href: format.href || null });
      continue;
    }
    if (child.nodeType !== 1) continue;
    const el = child;
    const tag = el.tagName;
    if (SKIP.has(tag)) continue;
    if (tag === 'IMG' || tag === 'PICTURE') { state.skippedImages++; continue; }
    const style = (el.getAttribute('style') || '').replace(/\s/g, '').toLowerCase();
    if (/display:none|mso-list:ignore|mso-hide:all/.test(style)) continue; // hidden text, Word's list markers
    if (tag === 'BR') {
      if (!state.current) startBlock(state);
      state.current.runs.push({ br: true });
      continue;
    }
    const f = formatOf(el, format, state.classes);
    if (tag === 'A') {
      const href = el.getAttribute('href');
      f.href = isLinkable(href) ? href.trim() : null;
      if (f.href) f.u = false; // the link's own style underlines it, as in Word
    }
    if (tag === 'TABLE' && !state.inCell) {
      // A table of its own (not one inside a cell, which is pasted as text in that cell).
      const table = parseTable(el, f, state);
      state.current = null;
      if (!table) continue;
      if (table.rows === 1 && table.cols === 1) {
        state.blocks.push(...table.cells[0].blocks); // one cell (as Excel copies one): its text
      } else if (table.cols > TABLE_MAX_COLS || table.rows > TABLE_MAX_ROWS) {
        state.blocks.push(...tablesAsText([{ table }]));
      } else {
        state.blocks.push({ table });
      }
      continue;
    }
    if (tag === 'TD' || tag === 'TH') {
      // Cells of a row become one paragraph, separated by tabs.
      if (!state.current) startBlock(state);
      else if (state.current.runs.length) state.current.runs.push({ text: '\t', ...f });
      walk(el, f, state);
      continue;
    }
    if (BLOCKS.has(tag)) {
      const list = tag === 'LI' ? htmlList(el) : msoList(el) || intactList(el);
      if (tag === 'UL' || tag === 'OL' || tag === 'TABLE') {
        walk(el, f, state);
        state.current = null;
        continue;
      }
      startBlock(state, list);
      walk(el, f, state);
      state.current = null;
      continue;
    }
    if (msoList(el)) {
      startBlock(state, msoList(el));
      walk(el, f, state);
      state.current = null;
      continue;
    }
    walk(el, f, state);
  }
}

/** Trims spaces at the edges of a block and merges neighbouring runs with the same formatting. */
function tidy(block) {
  const runs = [];
  for (const r of block.runs) {
    const last = runs[runs.length - 1];
    if (!r.br && last && !last.br && last.b === r.b && last.i === r.i && last.u === r.u && last.s === r.s && (last.va || null) === (r.va || null) && (last.href || null) === (r.href || null)) last.text += r.text;
    else runs.push({ ...r });
  }
  const first = runs.find((r) => !r.br);
  if (first && runs[0] === first) first.text = first.text.replace(/^ +/, '');
  const lastRun = runs[runs.length - 1];
  if (lastRun && !lastRun.br) lastRun.text = lastRun.text.replace(/ +$/, '');
  return { list: block.list, runs: runs.filter((r) => r.br || r.text) };
}

/**
 * DOM nodes for a block's runs: text wrapped in <strong>/<em>/<u>/<s> as needed, and neighbouring
 * runs of one link in a new link (<a data-link="new">).
 */
export function runNodes(runs) {
  const out = [];
  for (const r of runs) {
    if (r.br) {
      out.push(document.createElement('br'));
      continue;
    }
    let node = document.createTextNode(r.text);
    for (const [on, tag] of [[r.va === 'subscript', 'sub'], [r.va === 'superscript', 'sup'], [r.s, 's'], [r.u, 'u'], [r.i, 'em'], [r.b, 'strong']]) {
      if (!on) continue;
      const wrap = document.createElement(tag);
      wrap.append(node);
      node = wrap;
    }
    const last = out[out.length - 1];
    if (r.href && last?.tagName === 'A' && last.getAttribute('href') === r.href) last.append(node);
    else if (r.href) out.push(newLink(r.href, [node]));
    else out.push(node);
  }
  return out;
}

/** A link of this session to `href` around `nodes`: written as a Word hyperlink on save. */
export function newLink(href, nodes) {
  const a = document.createElement('a');
  a.href = href;
  a.setAttribute('href', href);
  a.dataset.link = 'new';
  a.title = `${href} (Ctrl+click to open)`;
  a.append(...nodes);
  return a;
}

/** Inserts nodes at the collapsed selection and puts the caret after them. */
export function insertAtCaret(nodes) {
  const sel = window.getSelection();
  if (!sel.rangeCount) return;
  const range = sel.getRangeAt(0);
  let last = null;
  for (const n of nodes) {
    range.insertNode(n);
    range.setStartAfter(n);
    range.collapse(true);
    last = n;
  }
  if (last) {
    sel.removeAllRanges();
    sel.addRange(range);
  }
}
