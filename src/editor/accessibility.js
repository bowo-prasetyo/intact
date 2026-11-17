// An accessibility check of the document, as Word's Accessibility Checker makes one (Review › Check
// Accessibility), for documents that must meet accessibility rules when published (Section 508,
// EN 301 549, WCAG). It reads the page as rendered: each issue lists the elements it was found in, so
// the editor can show them one by one.
import { isEmptyBlock } from './structure.js';

const DECORATIVE = 'http://schemas.microsoft.com/office/drawing/2017/decorative';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const BLOCKS = 'p, h1, h2, h3, h4, h5, h6';
const VAGUE_LINK = /^(click here|here|click|link|this link|this|more|read more|learn more|details|go|page)$/i;

/** Elements of the text itself: not in notes, headers, footers or the page's other bands. */
const inBody = (el) => !el.closest('.docx-notes, .docx-hf-band');
const onOff = (el) => !!el && !/^(0|false|off)$/.test(el.getAttributeNS(W, 'val') || '');
const wChild = (el, local) => el && [...el.children].find((c) => c.namespaceURI === W && c.localName === local);

/**
 * The accessibility issues in `page` (the editor's page) of `doc`: [{ id, severity, title, text,
 * targets }], severity 'error', 'warning' or 'tip' as in Word, `targets` the elements to show.
 */
export function checkAccessibility(page, doc) {
  const issues = [];
  const add = (id, severity, title, text, targets = []) => issues.push({ id, severity, title, text, targets });
  const plural = (n, one, many = `${one}s`) => `${n} ${n === 1 ? one : many}`;

  // Pictures need alternative text, unless marked decorative (Word 2019 and later).
  const pictures = [...page.querySelectorAll('.picture[data-raw]')];
  const decorative = (span) => {
    const run = doc.raw[Number(span.dataset.raw)];
    return [...(run?.getElementsByTagNameNS(DECORATIVE, 'decorative') || [])].some((d) => /^(1|true)$/.test(d.getAttribute('val') || ''));
  };
  const noAlt = pictures.filter((span) => !doc.pictureLayout(span)?.descr?.trim() && !decorative(span));
  if (noAlt.length) {
    add('alt', 'error', 'Missing alternative text', `${plural(noAlt.length, 'picture')} without alt text: people who can’t see ${noAlt.length === 1 ? 'it' : 'them'} won’t know what ${noAlt.length === 1 ? 'it shows' : 'they show'}. Show ${noAlt.length === 1 ? 'it' : 'each'} to describe it in Layout Options, or mark decorative pictures as decorative in Word.`, noAlt);
  }

  // Tables need a header row, so screen readers can name each cell's column.
  const tables = [...page.querySelectorAll('table.docx-table[data-raw]')].filter((t) => inBody(t) && !t.parentElement.closest('table') && t.rows.length > 1);
  const noHeader = tables.filter((t) => {
    const tbl = doc.raw[Number(t.dataset.raw)];
    const first = tbl && [...tbl.children].find((c) => c.namespaceURI === W && c.localName === 'tr');
    return !onOff(wChild(wChild(first, 'trPr'), 'tblHeader'));
  });
  if (noHeader.length) {
    add('header', 'error', 'No header row', `${plural(noHeader.length, 'table')} without a header row: screen readers can’t say which column a cell is in. In Word, select the first row and tick Table Properties › Row › Repeat as header row at the top of each page.`, noHeader);
  }

  // Headings in order, without skipping a level (Heading 1, then 2, not 3).
  const headings = [...page.querySelectorAll('h1, h2, h3, h4, h5, h6')].filter(inBody)
    .map((el) => ({ el, level: Number((doc.blockKind(el.dataset.style) || '').match(/^h(\d)$/)?.[1]) }))
    .filter((h) => h.level);
  const skipped = [];
  let previous = 1;
  for (const h of headings) {
    if (h.level > previous + 1) skipped.push(h.el);
    previous = h.level;
  }
  if (skipped.length) {
    add('levels', 'warning', 'Skipped heading level', `${plural(skipped.length, 'heading')} ${skipped.length === 1 ? 'skips' : 'skip'} a level (such as Heading 1 followed by Heading 3), which makes the document’s outline hard to follow. Change ${skipped.length === 1 ? 'its' : 'their'} style to the next level down.`, skipped);
  }

  // The document's language, for screen readers' pronunciation (and spelling).
  if (!doc.styles.language()) {
    add('language', 'warning', 'No document language', 'The document doesn’t say what language it is in, so screen readers may pronounce it wrongly. Set it in Word (Review › Language › Set Proofing Language, Set As Default).');
  }

  // Text in colours too faint to read: WCAG's contrast minimum (4.5:1, 3:1 for large text).
  const faint = [];
  for (const span of page.querySelectorAll('span[data-rpr]')) {
    if (!span.textContent.trim() || span.closest('[data-tracked], [data-hidden], .showing-placeholder, .sdt-locked, a')) continue;
    const rPr = doc.raw[Number(span.dataset.rpr)];
    const color = wChild(rPr, 'color');
    if (!color || (color.getAttributeNS(W, 'val') === 'auto' && !color.getAttributeNS(W, 'themeColor'))) continue;
    const style = getComputedStyle(span);
    const fg = rgb(style.color);
    if (!fg) continue;
    const ratio = contrast(fg, background(span));
    const size = parseFloat(style.fontSize);
    const large = size >= 24 || (size >= 18.66 && Number(style.fontWeight) >= 700);
    if (ratio < (large ? 3 : 4.5)) faint.push(span);
  }
  if (faint.length) {
    add('contrast', 'warning', 'Hard-to-read text contrast', `${plural(faint.length, 'passage')} of text in a colour too light (or dark) for its background, such as “${sample(faint[0])}” – choose a darker colour, or Automatic.`, faint);
  }

  // Links that say where they go: not "click here" or a bare web address.
  const vague = [...page.querySelectorAll('a[href]')].filter((a) => {
    const t = a.textContent.trim();
    return VAGUE_LINK.test(t) || /^(https?:\/\/|www\.)\S{30,}$/i.test(t);
  });
  if (vague.length) {
    add('links', 'tip', 'Unclear link text', `${plural(vague.length, 'link')} whose text doesn’t say where ${vague.length === 1 ? 'it goes' : 'they go'}, such as “${sample(vague[0])}” – people using screen readers often list links on their own: write what each one opens.`, vague);
  }

  // Blank paragraphs used for spacing: screen readers announce each one as "blank".
  const blanks = [];
  let run = [];
  const flush = () => { if (run.length >= 2) blanks.push(run[0]); run = []; };
  for (const el of page.children) {
    if (el.matches(BLOCKS) && isEmptyBlock(el)) run.push(el);
    else flush();
  }
  flush();
  if (blanks.length) {
    add('blank', 'tip', 'Repeated blank paragraphs', `${plural(blanks.length, 'place')} where empty paragraphs make space: screen readers announce each one as “blank”. Use the paragraph’s spacing (Paragraph › Spacing before and after) instead.`, blanks);
  }

  // Pictures not in line with text may be read out of order.
  const floating = pictures.filter((span) => (doc.pictureLayout(span)?.wrap || 'inline') !== 'inline');
  if (floating.length) {
    add('inline', 'tip', 'Pictures not in line with text', `${plural(floating.length, 'picture')} with text wrapping around ${floating.length === 1 ? 'it' : 'them'}: screen readers may read ${floating.length === 1 ? 'it' : 'them'} out of order. Where the layout allows, choose In line with text in Layout Options.`, floating);
  }

  // A title, which screen readers announce and PDFs made from the document show.
  if (!doc.properties?.core.some((p) => p.label === 'Title') && doc.newTitle == null) {
    add('title', 'tip', 'No title', `The document has no title in its properties, which screen readers announce and PDFs made from it show instead of the file name.${doc.coreOriginal == null ? ' Add one in Word (File › Info › Title).' : ''}`);
    issues.at(-1).form = doc.coreOriginal != null; // a title can be written into its core properties
  }
  return issues;
}

function sample(el) {
  const t = el.textContent.trim();
  return t.length > 40 ? `${t.slice(0, 40)}…` : t;
}

/** [r, g, b, a] from a computed CSS colour, or null. */
function rgb(css) {
  const m = /rgba?\(([\d.]+),\s*([\d.]+),\s*([\d.]+)(?:,\s*([\d.]+))?\)/.exec(css || '');
  return m ? [Number(m[1]), Number(m[2]), Number(m[3]), m[4] == null ? 1 : Number(m[4])] : null;
}

/** The colour behind `el`: the nearest background that isn't transparent, the page's white if none. */
function background(el) {
  for (let n = el; n && n.nodeType === 1; n = n.parentElement) {
    const c = rgb(getComputedStyle(n).backgroundColor);
    if (c && c[3] > 0.5) return c;
    if (n.classList.contains('docx-sheet')) break;
  }
  return [255, 255, 255, 1];
}

/** WCAG 2 contrast ratio of two colours. */
function contrast(a, b) {
  const lum = ([r, g, bl]) => {
    const f = (v) => {
      const c = v / 255;
      return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
    };
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(bl);
  };
  const [hi, lo] = [lum(a), lum(b)].sort((x, y) => y - x);
  return (hi + 0.05) / (lo + 0.05);
}
