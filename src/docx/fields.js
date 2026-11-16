// Fields that show a result Word works out: tables of contents, cross-references, caption numbers,
// and citations and bibliographies from Word or a reference manager (Zotero, EndNote, Mendeley).
// Intact keeps them as they are and never updates them, so after editing the text they can be out
// of date: the review says so (see DocxDocument.review and ReviewDialog).
import { NS, wChild, wAttr } from './xml.js';

const W = NS.w;

/** What a field instruction (" TOC \\o "1-3" ", " ADDIN ZOTERO_ITEM {…} ") is, or null. */
export function fieldKind(instr) {
  const text = (instr || '').trim();
  const code = text.split(/\s+/)[0]?.toUpperCase() || '';
  if (code === 'TOC' || code === 'TOA' || code === 'INDEX') return 'toc';
  if (code === 'REF' || code === 'PAGEREF' || code === 'NOTEREF') return 'ref';
  if (code === 'SEQ') return 'seq';
  if (code === 'CITATION') return 'citation';
  if (code === 'BIBLIOGRAPHY') return 'bibliography';
  if (code === 'ADDIN') {
    if (/ZOTERO_BIBL|EN\.REFLIST|MENDELEY BIBLIOGRAPHY|CSL_BIBLIOGRAPHY/i.test(text)) return 'bibliography';
    if (/ZOTERO_ITEM|EN\.CITE|CSL_CITATION|MENDELEY CITATION/i.test(text)) return 'citation';
  }
  return null;
}

/**
 * How many fields of each kind are under `root` (a body or a part): { toc, ref, seq, citation,
 * bibliography }. Simple fields (w:fldSimple), complex ones (begin, instruction, separate, end; also
 * nested) and the content controls Word wraps a table of contents, citation or bibliography in.
 */
export function countFields(root) {
  const count = { toc: 0, ref: 0, seq: 0, citation: 0, bibliography: 0 };
  const add = (kind) => { if (kind) count[kind]++; };
  const stack = []; // instructions of the complex fields open at this point
  const walk = (el) => {
    for (let n = el.firstElementChild; n; n = n.nextElementSibling) {
      if (n.namespaceURI === W) {
        const local = n.localName;
        if (local === 'fldSimple') add(fieldKind(wAttr(n, 'instr')));
        else if (local === 'fldChar') {
          const type = wAttr(n, 'fldCharType');
          if (type === 'begin') stack.push({ instr: '', counted: false });
          else if (type === 'separate' || type === 'end') {
            const f = type === 'end' ? stack.pop() : stack[stack.length - 1];
            if (f && !f.counted) {
              f.counted = true;
              add(fieldKind(f.instr));
            }
          }
          continue;
        } else if (local === 'instrText' && stack.length) {
          stack[stack.length - 1].instr += n.textContent;
          continue;
        } else if (local === 'sdt') {
          const pr = wChild(n, 'sdtPr');
          const gallery = wAttr(wChild(wChild(pr, 'docPartObj'), 'docPartGallery'), 'val') || '';
          if (/table of contents/i.test(gallery)) count.toc++;
          else if (/bibliograph/i.test(gallery) || wChild(pr, 'bibliography')) count.bibliography++;
          else if (wChild(pr, 'citation')) count.citation++;
          // A table of contents in a content control is also a TOC field: count it once.
          if (/table of contents|bibliograph/i.test(gallery) || wChild(pr, 'bibliography') || wChild(pr, 'citation')) continue;
        }
      }
      walk(n);
    }
  };
  if (root) walk(root);
  return count;
}

/**
 * What the review says about fields, comparing the text before and after: the kinds of field Word
 * has to update ("a table of contents", "cross-references", "caption numbers"), and whether
 * citations were added or removed while a bibliography lists them.
 */
export function fieldReview(before, after) {
  const a = countFields(before);
  const b = countFields(after);
  const updatable = [];
  if (b.toc) updatable.push(b.toc === 1 ? 'a table of contents' : 'tables of contents');
  if (b.ref) updatable.push('cross-references');
  if (b.seq) updatable.push('figure and table numbers');
  return {
    updatable,
    citationsChanged: b.bibliography > 0 && a.citation !== b.citation,
  };
}
