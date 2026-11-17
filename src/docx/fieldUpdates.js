// Updating fields here, as Word's F9 does, for what doesn't need pages: caption numbers (SEQ), note
// numbers in cross-references (NOTEREF), cross-references (REF), and tables of contents (TOC),
// rebuilt from the headings. Page numbers are Word's: Intact doesn't lay the text out in pages, so a
// table of contents keeps the number each entry showed (a new entry takes the one above it) and the
// file asks Word to update them when it opens it (see EditorView updateAllFields).
//
// Works on the page: a field's result is the editable span.field the renderer wraps it in, with its
// instruction in data-instr; its markers (begin, separate, end) are hidden raw runs around it.
import { NS, wChild, wAttr, createW, elementChildren, plainText } from './xml.js';
import { formatNumber } from './numbering.js';

const W = NS.w;
const BLOCKS = 'p, h1, h2, h3, h4, h5, h6';
const OUTSIDE_TEXT = '.docx-notes, .docx-hf-band, .textbox';

// The switches that take an argument, by field.
const ARG_SWITCHES = {
  TOC: new Set(['\\o', '\\t', '\\b', '\\f', '\\l', '\\p', '\\s', '\\d', '\\a', '\\c', '\\n']),
  SEQ: new Set(['\\r', '\\s', '\\*']),
  REF: new Set(['\\*', '\\d']),
  NOTEREF: new Set(['\\*']),
};

/** A field instruction (' SEQ Figure \\* ARABIC ') as { code: 'SEQ', args: ['Figure'], sw: { '\\*': 'ARABIC' } }. */
export function parseInstr(instr) {
  const tokens = [...(instr || '').matchAll(/"([^"]*)"|(\\\S+)|(\S+)/g)].map((m) => (m[2] ? { sw: m[2].toLowerCase() } : { text: m[1] ?? m[3] }));
  const code = (tokens.shift()?.text || '').toUpperCase();
  const takes = ARG_SWITCHES[code] || new Set();
  const args = [];
  const sw = {};
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i];
    if (!t.sw) {
      args.push(t.text);
      continue;
    }
    const next = tokens[i + 1];
    if (takes.has(t.sw) && next && !next.sw) {
      sw[t.sw] = next.text;
      i++;
    } else sw[t.sw] = true;
  }
  return { code, args, sw };
}

const SEQ_FORMATS = { ARABIC: 'decimal', ALPHABETIC: 'upperLetter', alphabetic: 'lowerLetter', ROMAN: 'upperRoman', roman: 'lowerRoman' };

/** The text nodes of a field's result: not those of fields nested in it, nor of content kept as it is. */
function resultNodes(field) {
  const out = [];
  const walker = document.createTreeWalker(field, NodeFilter.SHOW_TEXT);
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    if (n.parentElement.closest('.field') === field && !n.parentElement.closest('[data-raw]')) out.push(n);
  }
  return out;
}

/** Sets a field's result to `text`, in the formatting of its first piece of text. True if it changed. */
function setResult(field, text) {
  const nodes = resultNodes(field);
  if (nodes.map((n) => n.data).join('') === text) return false;
  if (nodes.length) {
    nodes[0].data = text;
    for (const n of nodes.slice(1)) n.remove();
  } else {
    const end = field.lastElementChild?.matches('[data-raw]') ? field.lastElementChild : null;
    field.insertBefore(document.createTextNode(text), end);
  }
  return true;
}

/** The bookmarks in `root`: name -> { start, end } (their hidden markers; end may be missing). */
function bookmarks(doc, root) {
  const starts = new Map();
  const ends = new Map();
  for (const el of root.querySelectorAll('.raw-hidden[data-raw]')) {
    const node = doc.raw[Number(el.dataset.raw)];
    if (node?.localName === 'bookmarkStart') starts.set(wAttr(node, 'name'), { start: el, id: wAttr(node, 'id') });
    else if (node?.localName === 'bookmarkEnd') ends.set(wAttr(node, 'id'), el);
  }
  const out = new Map();
  for (const [name, { start, id }] of starts) out.set(name, { start, end: ends.get(id) || null });
  return out;
}

/** A range over a bookmark's content (to its end, or to the end of its paragraph). */
function bookmarkRange({ start, end }) {
  const r = document.createRange();
  r.setStartAfter(start);
  if (end) r.setEndBefore(end);
  else {
    const block = start.closest(BLOCKS);
    if (block) r.setEnd(block, block.childNodes.length);
    else r.setEndAfter(start);
  }
  return r;
}

const textOf = (range) => range.toString().replace(/\s*\n+\s*/g, ' ').trim();

/** The paragraph a bookmark is in (or, for one between paragraphs, the next). */
function blockOf({ start }) {
  return start.closest(BLOCKS) || (start.nextElementSibling?.matches(BLOCKS) ? start.nextElementSibling : null);
}

/**
 * Updates the fields in the text of `root` (not notes, headers or footers): caption numbers, note
 * numbers, cross-references and tables of contents. Returns what changed: { seq, noteref, ref, toc,
 * entries, missing (cross-references to bookmarks not found), skipped (tables of contents left as
 * they are, with the reason) }.
 */
export function updateFields(doc, root) {
  const tally = { seq: 0, noteref: 0, ref: 0, toc: 0, entries: 0, missing: 0, skipped: null };
  const all = [...root.querySelectorAll('span.field[data-instr]')].filter((f) => !f.closest(OUTSIDE_TEXT));
  const parsed = new Map(all.map((f) => [f, parseInstr(f.dataset.instr)]));
  const inText = all.filter((f) => !f.closest('.sdt-locked'));

  // Caption numbers, in the order of the text; \s resets them at headings of that level.
  const counters = new Map();
  const lastOf = new Map(); // identifier -> the heading count at the last number, per \s level
  const headingsSeen = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
  const order = [...root.querySelectorAll(`${BLOCKS}, span.field[data-instr]`)].filter((el) => !el.closest(OUTSIDE_TEXT));
  for (const el of order) {
    if (!el.matches('span.field')) {
      const level = doc.outlineLevel(el);
      if (level != null && level < 9) for (let l = level; l < 9; l++) headingsSeen[l]++;
      continue;
    }
    if (!inText.includes(el)) continue;
    const { code, args, sw } = parsed.get(el);
    if (code !== 'SEQ' || !args[0]) continue;
    const id = args[0].toLowerCase();
    let n = counters.get(id) || 0;
    const sLevel = Number(sw['\\s']);
    if (sLevel >= 1 && sLevel <= 9) {
      const seen = headingsSeen[sLevel - 1];
      if (lastOf.get(id) !== undefined && lastOf.get(id) !== seen) n = 0;
      lastOf.set(id, seen);
    }
    if (sw['\\r'] !== undefined && sw['\\r'] !== true) n = Number(sw['\\r']) || 0;
    else if (!sw['\\c']) n += 1;
    counters.set(id, n);
    if (sw['\\h']) continue;
    const fmt = SEQ_FORMATS[typeof sw['\\*'] === 'string' ? sw['\\*'] : 'ARABIC'] || 'decimal';
    if (setResult(el, formatNumber(n, fmt))) tally.seq++;
  }

  const marks = bookmarks(doc, root);
  // Note numbers, then cross-references (whose text may hold the numbers just updated).
  for (const el of inText) {
    const { code, args, sw } = parsed.get(el);
    if (code !== 'NOTEREF' && code !== 'REF') continue;
    const mark = marks.get(args[0]);
    if (!mark) {
      tally.missing++;
      continue;
    }
    const range = bookmarkRange(mark);
    let text;
    if (code === 'NOTEREF') {
      const ref = [...root.querySelectorAll('.note-ref[data-note-ref]')].find((r) => !r.closest('.docx-notes') && range.intersectsNode(r));
      if (!ref) continue;
      text = ref.dataset.n;
    } else {
      const numbered = sw['\\r'] || sw['\\n'] || sw['\\w'];
      text = numbered ? (blockOf(mark)?.dataset.marker || '').replace(/\.$/, '') : textOf(range);
      if (sw['\\p']) {
        const where = mark.start.compareDocumentPosition(el) & Node.DOCUMENT_POSITION_FOLLOWING ? 'above' : 'below';
        text = numbered ? `${text} ${where}` : where;
      }
    }
    if (setResult(el, text)) tally[code === 'REF' ? 'ref' : 'noteref']++;
  }

  // Tables of contents, last: their entries are the headings as they now are.
  for (const el of all) {
    if (parsed.get(el).code !== 'TOC' || !el.isConnected) continue;
    const res = updateToc(doc, root, el, parsed.get(el).sw);
    if (res.skipped) tally.skipped = res.skipped;
    else {
      tally.toc++;
      tally.entries += res.entries;
    }
  }
  return tally;
}

// ------------------------------------------------------------------ tables of contents

/** 'begin' | 'separate' | 'end' for a hidden field marker in the page, else null. */
function markerType(doc, el) {
  const node = doc.raw[Number(el.dataset.raw)];
  if (node?.localName !== 'r') return null;
  const fld = wChild(node, 'fldChar');
  return fld ? wAttr(fld, 'fldCharType') : null;
}

/** The headings a table of contents lists, by its switches: [{ el, level (1-9), text }]. */
function tocHeadings(doc, root, sw, exclude) {
  const [from, to] = (typeof sw['\\o'] === 'string' ? sw['\\o'] : '').split('-').map(Number);
  const styleLevels = new Map();
  if (typeof sw['\\t'] === 'string') {
    const parts = sw['\\t'].split(/[,;]/).map((x) => x.trim());
    for (let i = 0; i + 1 < parts.length; i += 2) styleLevels.set(parts[i].toLowerCase(), Number(parts[i + 1]));
  }
  const out = [];
  for (const el of root.querySelectorAll(BLOCKS)) {
    if (el.closest(OUTSIDE_TEXT) || exclude(el)) continue;
    let level = null;
    const name = doc.styles.byId.get(el.dataset.style || doc.styles.defaultParagraph)?.name.toLowerCase();
    if (name && styleLevels.has(name)) level = styleLevels.get(name);
    else {
      const outline = doc.outlineLevel(el, !!sw['\\u']);
      if (outline != null && from && outline + 1 >= from && outline + 1 <= (to || from)) level = outline + 1;
    }
    if (!level) continue;
    const r = document.createRange();
    r.selectNodeContents(el);
    const text = textOf(r);
    if (text) out.push({ el, level, text });
  }
  return out;
}

/** The name of the heading's _Toc bookmark, added (as Word adds one) if it has none. */
function tocBookmark(doc, el, taken) {
  for (const m of el.querySelectorAll('.raw-hidden[data-raw]')) {
    const node = doc.raw[Number(m.dataset.raw)];
    if (node?.localName === 'bookmarkStart' && /^_Toc/.test(wAttr(node, 'name') || '')) return wAttr(node, 'name');
  }
  let name;
  do name = `_Toc${Math.floor(100000000 + Math.random() * 900000000)}`; while (taken.has(name));
  taken.add(name);
  const [start, end] = doc.newBookmark(name);
  el.insertBefore(start, el.firstChild);
  if (el.lastChild?.tagName === 'BR') el.insertBefore(end, el.lastChild);
  else el.append(end);
  return name;
}

const isW = (n, local) => n.nodeType === 1 && n.namespaceURI === W && n.localName === local;

/** The marker runs (and instruction) of fields, in order, in XML paragraphs: [{ p, run, type, instr }]. */
function xmlMarkers(paras) {
  const out = [];
  for (const p of paras) {
    for (const r of p.getElementsByTagNameNS(W, 'r')) {
      const fld = wChild(r, 'fldChar');
      if (fld) out.push({ p, run: r, type: wAttr(fld, 'fldCharType'), instr: '' });
      const it = wChild(r, 'instrText');
      if (it) {
        const open = [...out].reverse().find((m) => m.type === 'begin' && !m.closed);
        if (open) open.instr += it.textContent;
      }
      if (fld && wAttr(fld, 'fldCharType') === 'separate') {
        const open = [...out].reverse().find((m) => m.type === 'begin' && !m.closed);
        if (open) open.closed = true;
      }
    }
  }
  return out;
}

/** The TOC field's begin, separate and end markers in `paras` (the first TOC field), or null. */
function tocMarkers(paras) {
  const marks = xmlMarkers(paras);
  const b = marks.findIndex((m) => m.type === 'begin' && /^\s*TOC\b/i.test(m.instr));
  if (b < 0) return null;
  let depth = 0;
  let sep = null;
  for (let i = b; i < marks.length; i++) {
    if (marks[i].type === 'begin') depth++;
    else if (marks[i].type === 'separate' && depth === 1 && !sep) sep = marks[i];
    else if (marks[i].type === 'end' && --depth === 0) return { begin: marks[b], sep, end: marks[i] };
  }
  return null;
}

/** The text of a run's w:t in document order, and a paragraph's text before and after its last tab. */
function entryParts(p) {
  let before = '';
  let after = '';
  let seenTab = false;
  const walk = (el) => {
    for (const c of elementChildren(el)) {
      if (isW(c, 'instrText') || isW(c, 'pPr') || isW(c, 'rPr')) continue;
      if (isW(c, 'tab')) {
        seenTab = true;
        before += after;
        after = '';
      } else if (isW(c, 't')) {
        if (seenTab) after += c.textContent;
        else before += c.textContent;
      } else walk(c);
    }
  };
  walk(p);
  const anchor = [...p.getElementsByTagNameNS(W, 'hyperlink')].map((h) => wAttr(h, 'anchor')).find(Boolean)
    || [...p.getElementsByTagNameNS(W, 'instrText')].map((t) => /PAGEREF\s+(\S+)/i.exec(t.textContent)?.[1]).find(Boolean) || null;
  return seenTab ? { text: before.trim(), page: after.trim(), anchor } : { text: before.trim(), page: '', anchor };
}

/**
 * Rebuilds the table of contents in XML paragraphs `paras` (those from the one with its field's
 * begin to the one with its end) from `entries` ([{ level, text, bookmark }]). Each entry is made as
 * Word makes one, in the paragraph and run formatting the table has for its level (Word's TOC 1–9
 * styles where it has none), linked to its heading with \h, with its page number as a PAGEREF field.
 * Returns the new paragraphs.
 */
function rebuildToc(doc, paras, sw, entries) {
  const xml = doc.xml;
  const m = tocMarkers(paras);
  const first = paras[0];
  const last = paras[paras.length - 1];
  // As Word writes one: the field's instruction in its first paragraph, its end in its last.
  if (!m?.sep || m.begin.p !== first || m.sep.p !== first || m.end.p !== last) return null;
  const top = (p, run) => elementChildren(p).find((c) => c === run || c.contains(run));
  // What comes before the entries (to the field's separate marker) and after them (from its end).
  const firstKids = elementChildren(first).filter((c) => !isW(c, 'pPr'));
  const prefix = firstKids.slice(0, firstKids.indexOf(top(first, m.sep.run)) + 1);
  const lastKids = elementChildren(last).filter((c) => !isW(c, 'pPr'));
  const endAt = lastKids.indexOf(top(last, m.end.run));
  const suffix = lastKids.slice(endAt);
  const lastOwn = last !== first && !lastKids.slice(0, endAt).some((c) => plainText(c).trim());

  // The table as it was: each level's paragraph and run formatting, and each entry's page number.
  const templates = new Map();
  const pages = new Map();
  for (const p of paras) {
    const style = wAttr(wChild(wChild(p, 'pPr'), 'pStyle'), 'val');
    const level = Number(/^toc (\d)$/.exec(doc.styles.byId.get(style)?.name.toLowerCase() || '')?.[1] || /^TOC(\d)$/i.exec(style || '')?.[1]);
    const { text, page, anchor } = entryParts(p);
    if (!text) continue;
    if (level && !templates.has(level)) {
      const runs = [...p.getElementsByTagNameNS(W, 'r')];
      const textRun = runs.find((r) => wChild(r, 't'));
      const tabRun = runs.find((r) => wChild(r, 'tab'));
      const pPr = wChild(p, 'pPr')?.cloneNode(true) || null;
      if (pPr) for (const c of elementChildren(pPr)) if (isW(c, 'rPr') || isW(c, 'pPrChange') || isW(c, 'sectPr')) pPr.removeChild(c);
      templates.set(level, { pPr, textRPr: wChild(textRun, 'rPr'), tabRPr: wChild(tabRun, 'rPr') });
    }
    if (page) {
      if (anchor) pages.set(anchor, page);
      pages.set(text, page);
    }
  }

  const contentWidth = Math.round((doc.page.width - doc.page.left - doc.page.right) * 15);
  const run = (rPr, child, text) => {
    const r = createW(xml, 'r');
    if (rPr) r.appendChild(rPr.cloneNode(true));
    const c = r.appendChild(createW(xml, child));
    if (child === 't') {
      c.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
      c.textContent = text;
    } else if (child === 'instrText') {
      c.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
      c.textContent = text;
    } else if (child === 'fldChar') c.setAttributeNS(W, 'w:fldCharType', text);
    return r;
  };
  const defaultRPr = (hidden) => {
    const rPr = createW(xml, 'rPr');
    rPr.appendChild(createW(xml, 'noProof'));
    if (hidden) rPr.appendChild(createW(xml, 'webHidden'));
    return rPr;
  };
  const noPages = sw['\\n'] === true;
  let previous = '1';
  const out = entries.map((e) => {
    const t = templates.get(e.level);
    const p = createW(xml, 'p');
    if (t?.pPr) p.appendChild(t.pPr.cloneNode(true));
    else {
      const pPr = p.appendChild(createW(xml, 'pPr'));
      pPr.appendChild(createW(xml, 'pStyle')).setAttributeNS(W, 'w:val', doc.styles.ensureBuiltin(`toc${e.level}`));
      const tab = pPr.appendChild(createW(xml, 'tabs')).appendChild(createW(xml, 'tab'));
      tab.setAttributeNS(W, 'w:val', 'right');
      tab.setAttributeNS(W, 'w:leader', 'dot');
      tab.setAttributeNS(W, 'w:pos', String(contentWidth));
    }
    let holder = p;
    if (sw['\\h']) {
      holder = p.appendChild(createW(xml, 'hyperlink'));
      holder.setAttributeNS(W, 'w:anchor', e.bookmark);
      holder.setAttributeNS(W, 'w:history', '1');
    }
    // The table's own run formatting for the level (none, if it had none); Word's where it has no such entry.
    holder.appendChild(run(t ? t.textRPr : defaultRPr(false), 't', e.text));
    if (!noPages) {
      const page = pages.get(e.bookmark) || pages.get(e.text) || previous;
      previous = page;
      const rPr = t ? t.tabRPr : defaultRPr(true);
      holder.appendChild(run(rPr, 'tab'));
      holder.appendChild(run(rPr, 'fldChar', 'begin'));
      holder.appendChild(run(rPr, 'instrText', ` PAGEREF ${e.bookmark} \\h `));
      holder.appendChild(run(rPr, 'fldChar', 'separate'));
      holder.appendChild(run(rPr, 't', page));
      holder.appendChild(run(rPr, 'fldChar', 'end'));
    }
    return p;
  });
  if (!out.length) {
    const p = createW(xml, 'p');
    const pPr = wChild(first, 'pPr');
    if (pPr) p.appendChild(pPr.cloneNode(true));
    p.appendChild(run(null, 't', 'No table of contents entries found.'));
    out.push(p);
  }
  // The field's begin (and what comes before it) in the first paragraph; its end in the last, or in
  // a paragraph of its own as it was.
  const firstOut = out[0];
  const at = wChild(firstOut, 'pPr')?.nextSibling || firstOut.firstChild;
  for (const c of prefix) firstOut.insertBefore(c.cloneNode(true), at);
  if (lastOwn) {
    const p = last.cloneNode(false);
    if (wChild(last, 'pPr')) p.appendChild(wChild(last, 'pPr').cloneNode(true));
    for (const c of suffix) p.appendChild(c.cloneNode(true));
    out.push(p);
  } else {
    for (const c of suffix) out[out.length - 1].appendChild(c.cloneNode(true));
  }
  return out;
}

/**
 * Rebuilds the table of contents whose result starts at `field`, in the page: in Word's content
 * control (read-only here, rebuilt in its XML) or in the text (its paragraphs replaced). Returns
 * { entries } or { skipped: why }.
 */
function updateToc(doc, root, field, sw) {
  const locked = field.closest('.sdt-locked');
  const taken = new Set(bookmarks(doc, root).keys());
  if (locked) {
    const sdt = doc.raw[Number(locked.dataset.raw)].cloneNode(true);
    const content = wChild(sdt, 'sdtContent');
    const paras = elementChildren(content).filter((c) => isW(c, 'p'));
    const m = tocMarkers(paras);
    if (!m) return { skipped: 'a table of contents Intact can’t read' };
    const region = paras.slice(paras.indexOf(m.begin.p), paras.indexOf(m.end.p) + 1);
    const entries = tocHeadings(doc, root, sw, (el) => locked.contains(el)).map((h) => ({ ...h, bookmark: tocBookmark(doc, h.el, taken) }));
    const fresh = rebuildToc(doc, region, sw, entries);
    if (!fresh) return { skipped: 'a table of contents Intact can’t read' };
    for (const p of fresh) content.insertBefore(p, region[0]);
    for (const p of region) content.removeChild(p);
    locked.replaceWith(doc.renderControl(sdt, true));
    return { entries: entries.length };
  }
  // In the text: from the paragraph with its result's start to the one with its end marker.
  const markers = [...root.querySelectorAll('.raw-hidden[data-raw]')].filter((el) => markerType(doc, el));
  let begin = field.previousSibling;
  while (begin && !(begin.nodeType === 1 && begin.matches('[data-raw]') && markerType(doc, begin) === 'begin')) begin = begin.previousSibling;
  const from = markers.indexOf(begin);
  if (from < 0) return { skipped: 'a table of contents Intact can’t read' };
  let depth = 0;
  let end = null;
  for (let i = from; i < markers.length && !end; i++) {
    const type = markerType(doc, markers[i]);
    if (type === 'begin') depth++;
    else if (type === 'end' && --depth === 0) end = markers[i];
  }
  const firstBlock = field.closest(BLOCKS);
  const lastBlock = end?.closest(BLOCKS);
  if (!firstBlock || !lastBlock || firstBlock.parentElement !== lastBlock.parentElement) return { skipped: 'a table of contents Intact can’t read' };
  const blocks = [firstBlock];
  for (let b = firstBlock; b !== lastBlock;) {
    b = b.nextElementSibling;
    if (!b) return { skipped: 'a table of contents Intact can’t read' };
    blocks.push(b);
  }
  if (blocks.some((b) => b.querySelector('[data-tc], .inserted, .deleted') || !b.matches(BLOCKS))) {
    return { skipped: 'a table of contents with tracked changes in it (accept or reject them first)' };
  }
  const paras = blocks.map((b) => doc.paragraphXml(b));
  if (!tocMarkers(paras)) return { skipped: 'a table of contents Intact can’t read' };
  const entries = tocHeadings(doc, root, sw, (el) => blocks.includes(el)).map((h) => ({ ...h, bookmark: tocBookmark(doc, h.el, taken) }));
  const rebuilt = rebuildToc(doc, paras, sw, entries);
  if (!rebuilt) return { skipped: 'a table of contents Intact can’t read' };
  const fresh = rebuilt.map((p) => doc.renderParagraph(p));
  firstBlock.before(...fresh);
  for (const b of blocks) b.remove();
  return { entries: entries.length };
}
