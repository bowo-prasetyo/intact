// Paste with basic formatting. Clipboard HTML (from web pages, Google Docs or Word) is reduced to
// paragraphs, list items, line breaks and runs that are bold, italic, underlined, struck through,
// superscript or subscript.
// Everything else (fonts, colours, styles, links, pictures, data-* attributes) is dropped, because it
// can't be mapped faithfully to the document's own formatting.

const BLOCKS = new Set(['P', 'DIV', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'LI', 'BLOCKQUOTE', 'PRE', 'TR', 'TABLE', 'UL', 'OL', 'SECTION', 'ARTICLE', 'HEADER', 'FOOTER', 'DL', 'DT', 'DD', 'FIGURE']);
const SKIP = new Set(['SCRIPT', 'STYLE', 'TEMPLATE', 'HEAD', 'TITLE', 'META', 'LINK', 'NOSCRIPT', 'IFRAME', 'OBJECT', 'EMBED', 'SVG', 'CANVAS', 'VIDEO', 'AUDIO', 'BUTTON', 'SELECT', 'TEXTAREA', 'INPUT']);

/**
 * Parses clipboard data into { blocks, skippedImages }. A block is
 * { list: null | { kind: 'bullet' | 'number', level }, runs: [{ text, b, i, u, s } | { br: true }] }.
 */
export function parseClipboard(html, text) {
  if (!html) return { blocks: plainBlocks(text || ''), skippedImages: 0 };
  const doc = new DOMParser().parseFromString(html, 'text/html'); // inert: no scripts, no loading
  const state = { blocks: [], current: null, skippedImages: 0 };
  walk(doc.body, { b: false, i: false, u: false, s: false, pre: false }, state);
  const blocks = state.blocks.map(tidy);
  // Drop empty blocks at the start and end (fragment wrappers produce them).
  while (blocks.length && isEmpty(blocks[0])) blocks.shift();
  while (blocks.length && isEmpty(blocks[blocks.length - 1])) blocks.pop();
  if (!blocks.length && text) return { blocks: plainBlocks(text), skippedImages: state.skippedImages };
  return { blocks, skippedImages: state.skippedImages };
}

function plainBlocks(text) {
  return text.replace(/\r\n?/g, '\n').split('\n').map((line) => ({ list: null, runs: line ? [{ text: line, b: false, i: false, u: false, s: false }] : [] }));
}

const isEmpty = (block) => !block.list && !block.runs.some((r) => r.br || r.text.trim());

/** Styles from an element's tag and inline style; inline style wins (Google Docs' <b style="font-weight:normal">). */
function formatOf(el, inherited) {
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
  const st = el.style;
  if (st) {
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
  return f;
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

function startBlock(state, list = null) {
  state.current = { list, runs: [] };
  state.blocks.push(state.current);
}

function walk(node, format, state) {
  for (const child of node.childNodes) {
    if (child.nodeType === 3) {
      if (!state.current) startBlock(state);
      let text = child.data;
      if (!format.pre) text = text.replace(/[\t\n\r ]+/g, ' ');
      text = text.replace(/ /g, ' ');
      if (text) state.current.runs.push({ text, b: format.b, i: format.i, u: format.u, s: format.s, va: format.va || null });
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
    const f = formatOf(el, format);
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
    if (!r.br && last && !last.br && last.b === r.b && last.i === r.i && last.u === r.u && last.s === r.s && (last.va || null) === (r.va || null)) last.text += r.text;
    else runs.push({ ...r });
  }
  const first = runs.find((r) => !r.br);
  if (first && runs[0] === first) first.text = first.text.replace(/^ +/, '');
  const lastRun = runs[runs.length - 1];
  if (lastRun && !lastRun.br) lastRun.text = lastRun.text.replace(/ +$/, '');
  return { list: block.list, runs: runs.filter((r) => r.br || r.text) };
}

/** DOM nodes for a block's runs: text wrapped in <strong>/<em>/<u>/<s> as needed. */
export function runNodes(runs) {
  return runs.map((r) => {
    if (r.br) return document.createElement('br');
    let node = document.createTextNode(r.text);
    for (const [on, tag] of [[r.va === 'subscript', 'sub'], [r.va === 'superscript', 'sup'], [r.s, 's'], [r.u, 'u'], [r.i, 'em'], [r.b, 'strong']]) {
      if (!on) continue;
      const wrap = document.createElement(tag);
      wrap.append(node);
      node = wrap;
    }
    return node;
  });
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
