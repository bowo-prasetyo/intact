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
