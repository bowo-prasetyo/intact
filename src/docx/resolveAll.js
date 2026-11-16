// Accepting or rejecting every tracked change inside a piece of a document at once, in its XML, as
// Word's Accept All / Reject All would: for content Intact keeps as it is (locked content controls,
// locked cells, fields), where the changes aren't shown one by one. Insertions and deletions (and
// moves), paragraph breaks, formatting changes, and table rows, cells and merges are all resolved.
import { NS, wChild, wAttr, createW, elementChildren } from './xml.js';
import { resolveMerge } from './tables.js';

const W = NS.w;
const isW = (n, local) => n?.nodeType === 1 && n.namespaceURI === W && (!local || n.localName === local);
const PROPS = new Set(['rPr', 'pPr', 'trPr', 'tcPr', 'tblPr', 'numPr', 'sectPr', 'tblGrid']);

/** Turns deleted text back into text (w:delText → w:t, w:delInstrText → w:instrText). */
function undelete(el) {
  for (const [from, to] of [['delText', 't'], ['delInstrText', 'instrText']]) {
    for (const t of [...el.getElementsByTagNameNS(W, from)]) {
      const d = createW(el.ownerDocument, to);
      for (const a of [...t.attributes]) d.setAttributeNS(a.namespaceURI, a.name, a.value);
      d.textContent = t.textContent;
      t.replaceWith(d);
    }
  }
}

const unwrap = (el) => el.replaceWith(...el.childNodes);

/** A paragraph whose break goes: its content joins the next paragraph (which keeps its properties). */
function joinWithNext(p) {
  let next = p.nextElementSibling;
  while (next && !isW(next, 'p') && (isW(next, 'bookmarkStart') || isW(next, 'bookmarkEnd'))) next = next.nextElementSibling;
  if (!isW(next, 'p')) return false;
  const kids = elementChildren(p).filter((c) => !isW(c, 'pPr'));
  const nextPPr = wChild(next, 'pPr');
  if (nextPPr) nextPPr.after(...kids);
  else next.prepend(...kids);
  p.remove();
  return true;
}

/** Replaces a property element's content with its old content from its tracked change, keeping `keep`. */
function restore(props, change, keep = []) {
  const kept = elementChildren(props).filter((c) => keep.includes(c.localName) && c !== change);
  const old = elementChildren(change)[0];
  while (props.firstChild) props.removeChild(props.firstChild);
  for (const n of [...(old?.childNodes || [])]) props.appendChild(n);
  for (const n of kept) props.appendChild(n);
}

/** How many tracked changes `root` holds (as Accept All would count them). */
export function countRevisions(root) {
  const locals = ['ins', 'del', 'moveTo', 'moveFrom', 'rPrChange', 'pPrChange', 'tblPrChange', 'trPrChange', 'tcPrChange', 'tblGridChange', 'sectPrChange', 'cellIns', 'cellDel', 'cellMerge', 'numberingChange'];
  const self = isW(root) && locals.includes(root.localName) ? 1 : 0;
  return self + locals.reduce((n, local) => n + root.getElementsByTagNameNS(W, local).length, 0);
}

/** The authors of the tracked changes in `root`. */
export function revisionAuthors(root) {
  const out = new Set();
  const all = [root, ...root.getElementsByTagNameNS(W, '*')];
  for (const el of all) {
    if (isW(el) && /^(ins|del|moveTo|moveFrom|\w+PrChange|tblGridChange|cell(Ins|Del|Merge)|numberingChange)$/.test(el.localName) && wAttr(el, 'author')) out.add(wAttr(el, 'author'));
  }
  return [...out];
}

/**
 * Accepts (or rejects) every tracked change in `root` (a w:sdt, w:tc, w:r, w:p, w:fldSimple…), in
 * place. `root` itself may be a revision (an inserted run, say); then what it becomes replaces it in
 * its parent, and the result is returned (null if it went).
 */
export function resolveRevisionsIn(root, accept) {
  const holder = root.ownerDocument.createElementNS(W, 'w:holder');
  root.replaceWith?.(holder); // keep its place while it may be replaced
  holder.appendChild(root);
  const all = (local) => [...holder.getElementsByTagNameNS(W, local)];
  const inProps = (el) => PROPS.has(el.parentNode?.localName);
  for (const local of ['moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd']) for (const el of all(local)) el.remove();
  // Formatting changes: the record goes, or the old formatting comes back.
  for (const el of all('rPrChange')) accept ? el.remove() : restore(el.parentNode, el, ['ins', 'del', 'moveTo', 'moveFrom']);
  for (const el of all('pPrChange')) accept ? el.remove() : restore(el.parentNode, el, ['rPr', 'sectPr']);
  for (const local of ['tblPrChange', 'tblGridChange', 'sectPrChange']) for (const el of all(local)) accept ? el.remove() : restore(el.parentNode, el);
  for (const el of all('trPrChange')) accept ? el.remove() : restore(el.parentNode, el, ['ins', 'del']);
  for (const el of all('tcPrChange')) accept ? el.remove() : restore(el.parentNode, el, ['cellIns', 'cellDel', 'cellMerge']);
  for (const el of all('numberingChange')) el.remove();
  // Content: inserted text stays (or goes), deleted text goes (or comes back).
  for (const local of ['ins', 'moveTo']) for (const el of all(local).filter((e) => !inProps(e))) accept ? unwrap(el) : el.remove();
  for (const local of ['del', 'moveFrom']) {
    for (const el of all(local).filter((e) => !inProps(e))) {
      if (accept) el.remove();
      else {
        undelete(el);
        unwrap(el);
      }
    }
  }
  // Table rows, cells and merges.
  for (const el of all('ins').filter((e) => e.parentNode?.localName === 'trPr')) accept ? el.remove() : el.parentNode.parentNode.remove();
  for (const el of all('del').filter((e) => e.parentNode?.localName === 'trPr')) {
    if (accept) el.parentNode.parentNode.remove();
    else {
      undelete(el.parentNode.parentNode);
      el.remove();
    }
  }
  for (const el of all('cellIns')) accept ? el.remove() : el.parentNode.parentNode.remove();
  for (const el of all('cellDel')) {
    if (accept) el.parentNode.parentNode.remove();
    else {
      undelete(el.parentNode.parentNode);
      el.remove();
    }
  }
  for (const el of all('cellMerge')) resolveMerge(el.parentNode.parentNode, accept);
  for (const el of all('ins').filter((e) => e.parentNode?.localName === 'numPr')) accept ? el.remove() : el.parentNode.remove();
  // Paragraph breaks: an inserted break rejected, or a deleted one accepted, joins its paragraph with the next.
  for (const el of all('ins').filter((e) => e.parentNode?.localName === 'rPr')) {
    const p = el.parentNode.parentNode?.parentNode;
    el.remove();
    if (!accept && isW(p, 'p')) joinWithNext(p);
  }
  for (const el of all('del').filter((e) => e.parentNode?.localName === 'rPr')) {
    const p = el.parentNode.parentNode?.parentNode;
    el.remove();
    if (accept && isW(p, 'p')) joinWithNext(p);
  }
  for (const local of ['moveTo', 'moveFrom']) for (const el of all(local)) el.remove(); // marks of moved paragraphs
  // Empty property elements left behind are tidied.
  for (const local of ['rPr', 'trPr']) {
    for (const el of all(local)) if (!el.attributes.length && !el.childNodes.length && el.parentNode?.localName !== 'rPrChange') el.remove();
  }
  const result = holder.firstElementChild === holder.lastElementChild ? holder.firstElementChild : null;
  if (holder.parentNode) holder.replaceWith(...holder.childNodes);
  return result;
}
