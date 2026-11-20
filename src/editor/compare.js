// Compare documents, as Word's Review > Compare does for the text: another version of the document
// (a .docx chosen on this device, never uploaded) against the one open, paragraph by paragraph and,
// in a paragraph that changed, word by word. The other version is read as Word's text of its body:
// the paragraphs in order (table cells' too), with tracked deletions left out and insertions in, as
// the text reads once its changes are accepted. Formatting isn't compared.

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/** Collapsed whitespace, for telling whether two paragraphs say the same. */
const norm = (s) => s.replace(/[\s ]+/g, ' ').trim();

/**
 * The paragraphs of a document.xml's body, as text, in order: w:t and tabs and breaks, without what a
 * tracked change deletes (w:del, w:moveFrom) or a text box holds (its own story). Empty ones are left out.
 */
export function paragraphsOfXml(xml) {
  const dom = new DOMParser().parseFromString(xml, 'application/xml');
  if (dom.getElementsByTagName('parsererror').length) throw new Error('This file’s text couldn’t be read');
  const body = dom.getElementsByTagNameNS(W, 'body')[0];
  if (!body) return [];
  const out = [];
  for (const p of body.getElementsByTagNameNS(W, 'p')) {
    if (hasAncestor(p, body, 'txbxContent')) continue;
    let text = '';
    const walk = (node) => {
      for (const c of node.children) {
        if (c.namespaceURI !== W) {
          if (c.localName !== 'Fallback') walk(c); // a drawing's copy for older readers: once is enough
          continue;
        }
        const n = c.localName;
        if (n === 'del' || n === 'moveFrom' || n === 'pPr' || n === 'rPr' || n === 'instrText' || n === 'delText' || n === 'txbxContent' || n === 'p') continue;
        if (n === 't') text += c.textContent;
        else if (n === 'tab' || n === 'ptab') text += '\t';
        else if (n === 'br' || n === 'cr') text += ' ';
        else if (n === 'noBreakHyphen') text += '‑';
        else walk(c);
      }
    };
    walk(p);
    if (norm(text)) out.push(text);
  }
  return out;
}

function hasAncestor(el, stop, local) {
  for (let n = el.parentNode; n && n !== stop; n = n.parentNode) if (n.localName === local && n.namespaceURI === W) return true;
  return false;
}

/** The longest common subsequence of `a` and `b` (by `eq`), as pairs of indexes, in order. */
function lcs(a, b, eq = (x, y) => x === y) {
  const n = a.length;
  const m = b.length;
  // Common ends first: most versions differ in a few places.
  let start = 0;
  while (start < n && start < m && eq(a[start], b[start])) start++;
  let end = 0;
  while (end < n - start && end < m - start && eq(a[n - 1 - end], b[m - 1 - end])) end++;
  const pairs = [];
  for (let i = 0; i < start; i++) pairs.push([i, i]);
  const a2 = a.slice(start, n - end);
  const b2 = b.slice(start, m - end);
  const rows = a2.length;
  const cols = b2.length;
  if (rows && cols) {
    if (rows * cols > 40_000_000) throw new Error('These versions are too long to compare here');
    const table = new Uint32Array((rows + 1) * (cols + 1));
    for (let i = rows - 1; i >= 0; i--) {
      for (let j = cols - 1; j >= 0; j--) {
        table[i * (cols + 1) + j] = eq(a2[i], b2[j]) ? table[(i + 1) * (cols + 1) + j + 1] + 1
          : Math.max(table[(i + 1) * (cols + 1) + j], table[i * (cols + 1) + j + 1]);
      }
    }
    for (let i = 0, j = 0; i < rows && j < cols;) {
      if (eq(a2[i], b2[j])) {
        pairs.push([start + i, start + j]);
        i++;
        j++;
      } else if (table[(i + 1) * (cols + 1) + j] >= table[i * (cols + 1) + j + 1]) i++;
      else j++;
    }
  }
  for (let k = end; k > 0; k--) pairs.push([n - k, m - k]);
  return pairs;
}

/** Words and the spaces between them, so a diff of them reads back as the text. */
const words = (s) => s.match(/\s+|[^\s]+/g) || [];

/**
 * Word by word, what changed from `before` to `after`: [{ kind: 'same' | 'del' | 'ins', text }],
 * neighbouring pieces of a kind joined.
 */
export function diffWords(before, after) {
  const a = words(before);
  const b = words(after);
  const pairs = lcs(a, b);
  const out = [];
  const push = (kind, text) => {
    if (!text) return;
    const last = out[out.length - 1];
    if (last && last.kind === kind) last.text += text;
    else out.push({ kind, text });
  };
  let i = 0;
  let j = 0;
  for (const [pi, pj] of [...pairs, [a.length, b.length]]) {
    push('del', a.slice(i, pi).join(''));
    push('ins', b.slice(j, pj).join(''));
    if (pi < a.length) push('same', a[pi]);
    i = pi + 1;
    j = pj + 1;
  }
  // A space left between a deletion and an insertion reads as part of both: keep it plain.
  return out;
}

/**
 * The changes from version `before` (paragraph texts) to `after`: [{ kind: 'changed' | 'added' |
 * 'removed', at (the index in `after` of the paragraph, or of the one it would come before), parts
 * (diffWords' pieces) }], in the order of the text, and { added, removed, changed } counts.
 */
export function compareParagraphs(before, after) {
  const pairs = lcs(before, after, (x, y) => norm(x) === norm(y));
  const changes = [];
  let i = 0;
  let j = 0;
  for (const [pi, pj] of [...pairs, [before.length, after.length]]) {
    const gone = before.slice(i, pi);
    const come = after.slice(j, pj);
    // Paragraphs changed in place: as many as both sides have, word by word; the rest added or removed.
    const both = Math.min(gone.length, come.length);
    for (let k = 0; k < both; k++) changes.push({ kind: 'changed', at: j + k, parts: diffWords(gone[k], come[k]) });
    for (let k = both; k < come.length; k++) changes.push({ kind: 'added', at: j + k, parts: [{ kind: 'ins', text: come[k] }] });
    for (let k = both; k < gone.length; k++) changes.push({ kind: 'removed', at: pj, parts: [{ kind: 'del', text: gone[k] }] });
    i = pi + 1;
    j = pj + 1;
  }
  const count = (kind) => changes.filter((c) => c.kind === kind).length;
  return { changes, added: count('added'), removed: count('removed'), changed: count('changed') };
}
