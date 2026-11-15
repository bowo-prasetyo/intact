// "Review my changes": what differs between the document as opened and as it is about to be saved,
// paragraph by paragraph, with the words that changed. Paragraphs are compared by meaning, as the
// round-trip tests compare them (tests/helpers.js canonicalize): revision-session ids, paragraph
// ids, spelling marks and how text is split into runs don't count, so an untouched paragraph is
// never reported.
import { NS, wChild, wAttr } from './xml.js';

const W = NS.w;
const IGNORED = new Set(['proofErr', 'lastRenderedPageBreak']);
const ON = new Set(['1', 'true', 'on']);
const MAX_CELLS = 4_000_000; // beyond this many comparisons, a stretch counts as rewritten

const isW = (n, local) => n.nodeType === 1 && n.namespaceURI === W && (!local || n.localName === local);

function attrs(el) {
  return [...el.attributes]
    .filter((a) => !/^w:rsid|^w14:paraId$|^w14:textId$|^xml:space$/.test(a.name) && !a.name.startsWith('xmlns'))
    .filter((a) => !(a.name === 'w:val' && ON.has(a.value)))
    .map((a) => `${a.name}="${a.value}"`)
    .sort()
    .join(' ');
}

/** A canonical string for an element: equal for elements Word treats as the same. */
function canon(el) {
  const out = [];
  let text = null; // the run text being merged: { key, value }
  const flush = () => { if (text) { out.push(`T[${text.key}]${JSON.stringify(text.value)}`); text = null; } };
  for (const c of el.childNodes) {
    if (c.nodeType !== 1 || (isW(c) && IGNORED.has(c.localName))) continue;
    if (isW(c, 'r')) {
      const rPr = wChild(c, 'rPr');
      const key = rPr ? canon(rPr) : '';
      for (const k of c.childNodes) {
        if (k.nodeType !== 1 || isW(k, 'rPr') || (isW(k) && IGNORED.has(k.localName))) continue;
        const piece = isW(k, 't') ? k.textContent : isW(k, 'tab') ? '\t'
          : (isW(k, 'br') && !wAttr(k, 'type')) || isW(k, 'cr') ? '\n' : null;
        if (piece === null) { flush(); out.push(`R[${key}]${canon(k)}`); continue; }
        if (!piece) continue;
        if (text && text.key === key) text.value += piece;
        else { flush(); text = { key, value: piece }; }
      }
      continue;
    }
    flush();
    out.push(`<${c.nodeName} ${attrs(c)}>${canon(c)}</>`);
  }
  flush();
  return out.join('');
}

/** A paragraph's text as it reads with tracked changes accepted (deleted text left out). */
export function paragraphText(p) {
  let s = '';
  const walk = (el) => {
    for (const c of el.childNodes) {
      if (c.nodeType !== 1) continue;
      if (isW(c, 'p') || isW(c, 'del') || isW(c, 'moveFrom') || isW(c, 'pPr') || isW(c, 'rPr')) continue;
      if (isW(c, 't')) s += c.textContent;
      else if (isW(c, 'tab')) s += '\t';
      else if ((isW(c, 'br') && !wAttr(c, 'type')) || isW(c, 'cr')) s += '\n';
      else if (isW(c, 'noBreakHyphen')) s += '‑';
      else walk(c);
    }
  };
  walk(p);
  return s;
}

/** Every paragraph of a body in order (in tables and fields too), with what is needed to compare. */
function paragraphs(body) {
  return [...body.getElementsByTagNameNS(W, 'p')].map((p) => ({
    p,
    key: canon(p),
    text: paragraphText(p),
    inTable: !!closestW(p, 'tc'),
  }));
}

function closestW(el, local) {
  for (let n = el.parentNode; n && n.nodeType === 1; n = n.parentNode) if (isW(n, local)) return n;
  return null;
}

/**
 * The longest common subsequence of a and b (arrays compared with ===) as edit operations:
 * [['same', i, j] | ['del', i] | ['ins', j]], or null if the arrays are too long to compare.
 */
function diff(a, b) {
  const n = a.length;
  const m = b.length;
  if (n * m > MAX_CELLS) return null;
  const table = new Uint32Array((n + 1) * (m + 1));
  const at = (i, j) => i * (m + 1) + j;
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      table[at(i, j)] = a[i] === b[j] ? table[at(i + 1, j + 1)] + 1 : Math.max(table[at(i + 1, j)], table[at(i, j + 1)]);
    }
  }
  const ops = [];
  let i = 0;
  let j = 0;
  while (i < n && j < m) {
    if (a[i] === b[j]) ops.push(['same', i++, j++]);
    else if (table[at(i + 1, j)] >= table[at(i, j + 1)]) ops.push(['del', i++]);
    else ops.push(['ins', j++]);
  }
  while (i < n) ops.push(['del', i++]);
  while (j < m) ops.push(['ins', j++]);
  return ops;
}

/** Operations for a and b with a common start and end trimmed first (most edits are local). */
function diffTrimmed(a, b) {
  let s = 0;
  while (s < a.length && s < b.length && a[s] === b[s]) s++;
  let e = 0;
  while (e < a.length - s && e < b.length - s && a[a.length - 1 - e] === b[b.length - 1 - e]) e++;
  const midA = a.slice(s, a.length - e);
  const midB = b.slice(s, b.length - e);
  const mid = diff(midA, midB)
    || [...midA.map((_, i) => ['del', i]), ...midB.map((_, j) => ['ins', j])];
  const ops = [];
  for (let k = 0; k < s; k++) ops.push(['same', k, k]);
  for (const op of mid) ops.push(op[0] === 'same' ? ['same', op[1] + s, op[2] + s] : op[0] === 'del' ? ['del', op[1] + s] : ['ins', op[1] + s]);
  for (let k = 0; k < e; k++) ops.push(['same', a.length - e + k, b.length - e + k]);
  return ops;
}

/** The words that changed between two texts: [{ type: 'same' | 'del' | 'ins', text }]. */
export function wordDiff(before, after) {
  const a = before.match(/\s+|[^\s]+/g) || [];
  const b = after.match(/\s+|[^\s]+/g) || [];
  const out = [];
  const push = (type, text) => {
    const last = out[out.length - 1];
    if (last?.type === type) last.text += text;
    else out.push({ type, text });
  };
  for (const op of diffTrimmed(a, b)) {
    if (op[0] === 'same') push('same', a[op[1]]);
    else if (op[0] === 'del') push('del', a[op[1]]);
    else push('ins', b[op[1]]);
  }
  // Show deletions before insertions within a changed stretch, as reviewers read them.
  for (let k = 0; k + 1 < out.length; k++) {
    if (out[k].type === 'ins' && out[k + 1].type === 'del') [out[k], out[k + 1]] = [out[k + 1], out[k]];
  }
  return out;
}

/** How alike two texts are, from 0 to 1: the share of their characters in common words. */
function similarity(a, b) {
  if (a === b) return 1;
  if (!a.trim() || !b.trim()) return 0;
  let same = 0;
  for (const w of wordDiff(a, b)) if (w.type === 'same') same += w.text.replace(/\s+/g, '').length;
  const size = a.replace(/\s+/g, '').length + b.replace(/\s+/g, '').length;
  return size ? (2 * same) / size : 0;
}

/** What changed in a paragraph whose text didn't: its style, alignment, list, or text formatting. */
function formattingChanges(before, after, styleName) {
  const pPr = (p) => wChild(p, 'pPr');
  const val = (p, local) => wAttr(wChild(pPr(p), local), 'val');
  const out = [];
  const mark = (p, local) => !!wChild(wChild(pPr(p), 'rPr'), local);
  if (mark(after, 'del') && !mark(before, 'del')) out.push('Paragraph break deleted: it joins the next paragraph when accepted');
  if (mark(after, 'ins') && !mark(before, 'ins')) out.push('Paragraph break inserted');
  if (out.length) return out;
  const [s1, s2] = [val(before, 'pStyle'), val(after, 'pStyle')];
  if (s1 !== s2) out.push(`Style: ${styleName(s1)} → ${styleName(s2)}`);
  const align = (v) => ({ both: 'justified', center: 'centred', right: 'right', end: 'right', left: 'left', start: 'left' }[v] || 'left');
  if (align(val(before, 'jc')) !== align(val(after, 'jc'))) out.push(`Alignment: ${align(val(before, 'jc'))} → ${align(val(after, 'jc'))}`);
  const list = (p) => { const n = wChild(pPr(p), 'numPr'); return n ? canon(n) : ''; };
  if (list(before) !== list(after)) out.push('List or list level changed');
  if (!out.length) out.push('Text formatting changed (bold, italic, underline or strikethrough)');
  return out;
}

/**
 * The paragraphs that differ between two bodies, in order:
 * [{ kind: 'changed' | 'added' | 'removed', number, inTable, words, formatting, tracked }].
 * `number` counts paragraphs in the saved document (in the original one for removed paragraphs);
 * `tracked(p)` says whether a saved paragraph holds this session's tracked changes.
 */
export function reviewParagraphs(beforeBody, afterBody, { tracked = () => false, styleName = (id) => id || 'Normal' } = {}) {
  const A = paragraphs(beforeBody);
  const B = paragraphs(afterBody);
  const ops = diffTrimmed(A.map((x) => x.key), B.map((x) => x.key));
  const entries = [];
  let gone = [];
  let added = [];
  const removed = (i) => ({ kind: 'removed', number: i + 1, inTable: A[i].inTable, words: [{ type: 'del', text: A[i].text }], formatting: null, tracked: false });
  const addition = (j) => ({ kind: 'added', number: j + 1, inTable: B[j].inTable, words: [{ type: 'ins', text: B[j].text }], formatting: null, tracked: tracked(B[j].p) });
  // Within a stretch of edits, a removed and an added paragraph are one changed paragraph only if
  // their text is alike (in order); the others were removed or added.
  const flush = () => {
    let g = 0;
    for (const j of added) {
      const b = B[j];
      let match = -1;
      for (let k = g; k < gone.length; k++) {
        if (similarity(A[gone[k]].text, b.text) >= 0.5) { match = k; break; }
      }
      if (match < 0) { entries.push(addition(j)); continue; }
      for (; g < match; g++) entries.push(removed(gone[g]));
      const a = A[gone[g++]];
      const same = a.text === b.text;
      entries.push({
        kind: 'changed',
        number: j + 1,
        inTable: b.inTable,
        words: same ? [{ type: 'same', text: b.text }] : wordDiff(a.text, b.text),
        formatting: same ? formattingChanges(a.p, b.p, styleName) : null,
        tracked: tracked(b.p),
      });
    }
    for (; g < gone.length; g++) entries.push(removed(gone[g]));
    gone = [];
    added = [];
  };
  for (const op of ops) {
    if (op[0] === 'same') flush();
    else if (op[0] === 'del') gone.push(op[1]);
    else added.push(op[1]);
  }
  flush();
  return entries;
}
