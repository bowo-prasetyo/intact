// Tracked changes (revisions) in plain language. Intact shows revisions read-only and writes them
// back verbatim; these labels say what would be lost if an edit removed one, so the editor can
// refuse that edit and tell the user why.
import { NS, wChild, wAttr } from './xml.js';

const W = NS.w;
const INLINE = { ins: 'insertion', del: 'deletion', moveTo: 'move', moveFrom: 'move' };
const CONTENT = ['ins', 'del', 'moveTo', 'moveFrom', 'cellIns', 'cellDel'];

const byAuthor = (authors) => (authors.length ? ` by ${authors.join(', ')}` : '');

function authorsOf(nodes) {
  return [...new Set(nodes.map((n) => wAttr(n, 'author')).filter(Boolean))];
}

/**
 * What removing `node` (an element kept verbatim) would throw away, if it is or holds a tracked
 * change: "a tracked insertion by Ann", "tracked changes by Ann, Bob"; otherwise null.
 */
export function trackedChangeLabel(node) {
  if (!node || node.nodeType !== 1) return null;
  if (node.namespaceURI === W && INLINE[node.localName]) {
    return `a tracked ${INLINE[node.localName]}${byAuthor(authorsOf([node]))}`;
  }
  const found = CONTENT.flatMap((local) => [...node.getElementsByTagNameNS(W, local)]);
  if (!found.length) return null;
  const what = node.namespaceURI === W && node.localName === 'tbl' ? 'a table with tracked changes' : 'tracked changes';
  return `${what}${byAuthor(authorsOf(found))}`;
}

/** If a paragraph's mark (its "¶") is itself a tracked change, what it is; otherwise null. */
export function paragraphMarkLabel(pPr) {
  const rPr = wChild(pPr, 'rPr');
  const marks = ['ins', 'del', 'moveTo', 'moveFrom'].map((l) => wChild(rPr, l)).filter(Boolean);
  if (!marks.length) return null;
  return `a tracked paragraph break${byAuthor(authorsOf(marks))}`;
}

/** If a paragraph's mark is a tracked change: { kind: 'ins' | 'del', local, author, date }; else null. */
export function paragraphMarkRevision(pPr) {
  const rPr = wChild(pPr, 'rPr');
  for (const local of ['ins', 'del', 'moveTo', 'moveFrom']) {
    const el = wChild(rPr, local);
    if (el) return { kind: local === 'ins' || local === 'moveTo' ? 'ins' : 'del', local, author: wAttr(el, 'author'), date: wAttr(el, 'date') };
  }
  return null;
}

/** For an inline revision element (w:ins, w:del, w:moveTo, w:moveFrom): 'ins' or 'del'; else null. */
export function inlineRevisionKind(node) {
  if (node?.namespaceURI !== W) return null;
  return { ins: 'ins', moveTo: 'ins', del: 'del', moveFrom: 'del' }[node.localName] || null;
}

/**
 * A tracked formatting change in run or paragraph properties (`props`: <w:rPr> or <w:pPr>):
 * { author, date } of its <w:rPrChange> / <w:pPrChange>; for a paragraph, also its mark's
 * <w:rPr><w:rPrChange>. Null if there is none.
 */
export function formattingChange(props) {
  if (!props) return null;
  const change = wChild(props, 'rPrChange') || wChild(props, 'pPrChange')
    || (props.localName === 'pPr' ? wChild(wChild(props, 'rPr'), 'rPrChange') : null);
  return change ? { author: wAttr(change, 'author'), date: wAttr(change, 'date') } : null;
}

/** If a table row (<w:tr>) was inserted or deleted with tracking: { kind, author, date }; else null. */
export function rowRevision(tr) {
  const trPr = wChild(tr, 'trPr');
  for (const local of ['ins', 'del']) {
    const el = wChild(trPr, local);
    if (el) return { kind: local, author: wAttr(el, 'author'), date: wAttr(el, 'date') };
  }
  return null;
}

/** If a table cell (<w:tc>) was inserted or deleted with tracking: { kind, author, date }; else null. */
export function cellRevision(tc) {
  const tcPr = wChild(tc, 'tcPr');
  for (const [local, kind] of [['cellIns', 'ins'], ['cellDel', 'del']]) {
    const el = wChild(tcPr, local);
    if (el) return { kind, author: wAttr(el, 'author'), date: wAttr(el, 'date'), id: wAttr(el, 'id') };
  }
  return null;
}

/**
 * A tracked change to a table's, row's or cell's properties: the <w:tblPrChange> (or
 * <w:tblGridChange>), <w:trPrChange> or <w:tcPrChange> of `el` (a w:tbl, w:tr or w:tc), as
 * { author, date }; null if there is none.
 */
export function propertyChange(el) {
  if (!el) return null;
  const local = el.localName;
  const changes = local === 'tbl'
    ? [wChild(wChild(el, 'tblPr'), 'tblPrChange'), wChild(wChild(el, 'tblGrid'), 'tblGridChange')]
    : [wChild(wChild(el, local === 'tr' ? 'trPr' : 'tcPr'), local === 'tr' ? 'trPrChange' : 'tcPrChange')];
  const found = changes.find(Boolean);
  return found ? { author: wAttr(found, 'author'), date: wAttr(found, 'date') } : null;
}
