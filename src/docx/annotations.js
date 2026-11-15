// Read-only parts shown next to the body: comments (with replies and resolved state), footnote and
// endnote text, and the header and footer of the first page. Plus plain-language names for the
// content the editor keeps but can't edit.
import { NS, parseXml, parseRels, wChild, wChildren, wAttr, onOff, plainText } from './xml.js';

const W = NS.w;
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';

/** Loads a part related to the main document by relationship type suffix. */
export async function loadPart(doc, rel) {
  if (!rel || rel.external || !rel.path) return null;
  const text = await doc.text(rel.path);
  if (!text) return null;
  const slash = rel.path.lastIndexOf('/');
  const dir = slash >= 0 ? rel.path.slice(0, slash) : '';
  const relsText = await doc.text(`${dir ? dir + '/' : ''}_rels/${rel.path.slice(slash + 1)}.rels`);
  return { path: rel.path, xml: parseXml(text), rels: parseRels(relsText, dir) };
}

const paragraphsOf = (el) => wChildren(el, 'p').map((p) => plainText(p).trim()).filter(Boolean);

/**
 * Comments as [{ id, author, initials, date, paragraphs, parentId, done }]. Replies and the
 * resolved state come from commentsExtended.xml (Word 2013+), keyed by each comment's last paraId.
 */
export function parseComments(commentsPart, extendedPart) {
  if (!commentsPart) return [];
  const ext = new Map();
  for (const ce of extendedPart?.xml.getElementsByTagNameNS(W15, 'commentEx') || []) {
    ext.set(ce.getAttributeNS(W15, 'paraId'), {
      parent: ce.getAttributeNS(W15, 'paraIdParent') || null,
      done: ['1', 'true'].includes(ce.getAttributeNS(W15, 'done')),
    });
  }
  const byParaId = new Map();
  const comments = wChildren(commentsPart.xml.documentElement, 'comment').map((c) => {
    const paras = wChildren(c, 'p');
    const paraId = paras.length ? paras[paras.length - 1].getAttributeNS(W14, 'paraId') : null;
    const comment = {
      id: wAttr(c, 'id'),
      author: wAttr(c, 'author') || 'Unknown author',
      initials: wAttr(c, 'initials') || '',
      date: wAttr(c, 'date') || '',
      paragraphs: paragraphsOf(c),
      paraId,
      parentId: null,
      done: false,
    };
    if (paraId) byParaId.set(paraId, comment);
    return comment;
  });
  for (const c of comments) {
    const e = c.paraId && ext.get(c.paraId);
    if (!e) continue;
    c.done = e.done;
    c.parentId = e.parent ? byParaId.get(e.parent)?.id ?? null : null;
  }
  return comments;
}

/** Footnote or endnote text by id, skipping the separator notes. */
export function parseNotes(part, local) {
  const map = new Map();
  for (const n of wChildren(part?.xml.documentElement, local)) {
    const type = wAttr(n, 'type');
    if (type && type !== 'normal') continue;
    map.set(wAttr(n, 'id'), paragraphsOf(n).join(' '));
  }
  return map;
}

/**
 * Every header (local 'headerReference') or footer in use, as [{ rid, label, type }], in the order
 * Word uses them: per section, first page / other (or odd) pages / even pages. A section without its
 * own reference of a type inherits the previous section's, as in Word. "First page" counts only
 * where the section has a different first page; "even pages" only with odd/even headers on.
 * The first entry is what page 1 shows.
 */
export function headerFooterVariants(body, local, evenAndOdd) {
  const sections = [...body.getElementsByTagNameNS(W, 'sectPr')];
  const inherited = {};
  const out = [];
  sections.forEach((sect, i) => {
    for (const ref of wChildren(sect, local)) inherited[wAttr(ref, 'type') || 'default'] = ref.getAttributeNS(NS.r, 'id');
    const titlePg = onOff(wChild(sect, 'titlePg')) === true;
    const used = [];
    if (titlePg && inherited.first) used.push(['first', 'First page']);
    if (inherited.default) used.push(['default', evenAndOdd ? 'Odd pages' : titlePg ? 'Other pages' : 'All pages']);
    if (evenAndOdd && inherited.even) used.push(['even', 'Even pages']);
    for (const [type, label] of used) {
      const rid = inherited[type];
      if (out.some((v) => v.rid === rid)) continue;
      out.push({ rid, type, label, section: i + 1 });
    }
  });
  const multi = out.some((v) => v.section > 1);
  for (const v of out) if (multi) v.label = `Section ${v.section}: ${v.label.toLowerCase()}`;
  return out;
}

// ------------------------------------------------------------------ plain-language names

const FIELD_NAMES = {
  PAGE: 'Page number', NUMPAGES: 'Number of pages', SECTIONPAGES: 'Pages in section',
  DATE: 'Date', TIME: 'Time', CREATEDATE: 'Creation date', SAVEDATE: 'Date last saved',
  PRINTDATE: 'Date last printed', TOC: 'Table of contents', REF: 'Cross-reference',
  PAGEREF: 'Page reference', NOTEREF: 'Note reference', HYPERLINK: 'Link',
  MERGEFIELD: 'Mail merge field', FILENAME: 'File name', AUTHOR: 'Author', TITLE: 'Title',
  SEQ: 'Numbering', DOCPROPERTY: 'Document property', FORMTEXT: 'Form field',
  FORMCHECKBOX: 'Form check box', FORMDROPDOWN: 'Form drop-down', INDEX: 'Index',
};

/** "Page number", "Cross-reference", … for a field instruction such as " PAGE \\* MERGEFORMAT ". */
export function fieldName(instr) {
  const code = (instr || '').trim().split(/\s+/)[0]?.toUpperCase() || '';
  return FIELD_NAMES[code] || (code ? `Field (${code})` : 'Field');
}

/** What a piece of preserved content is, in words a Word user knows. */
export function describe(node) {
  const has = (ns, local) => node.getElementsByTagNameNS(ns, local).length > 0;
  const uri = [...node.getElementsByTagNameNS(NS.a, 'graphicData')].map((g) => g.getAttribute('uri') || '').join(' ');
  if (/chart/.test(uri)) return 'Chart';
  if (/diagram/.test(uri)) return 'SmartArt graphic';
  if (/wordprocessingShape|wordprocessingGroup/.test(uri)) return 'Shape or text box';
  if (has(NS.a, 'blip') || has(NS.v, 'imagedata')) return 'Picture';
  if (has(W, 'object')) return 'Embedded object';
  if (has(W, 'pict') || has(NS.v, 'shape')) return 'Shape';
  const MATH = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
  if (has(MATH, 'oMath') || (node.namespaceURI === MATH && /^oMath(Para)?$/.test(node.localName))) return 'Equation';
  if (node.localName === 'fldSimple') return fieldName(wAttr(node, 'instr'));
  if (has(W, 'fldChar') || has(W, 'instrText')) return 'Field';
  if (has(W, 'ptab')) return 'Tab';
  if (node.localName === 'altChunk') return 'Embedded content';
  if (node.localName === 'customXml') return 'Tagged content';
  return 'Content Intact can’t edit';
}
