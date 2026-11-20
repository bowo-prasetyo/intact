// Headers and footers on every printed page, as Word prints them. The page here is one long sheet
// with its header once at the top and its footer once at the bottom; printing (and so saving as PDF)
// cuts it into pages, and each page should have them, in its margins, with its own page number.
//
// Browsers give two ways, each used where it can be faithful:
//  - "fixed": a header or footer that is the same on every page is printed as a fixed box, which
//    browsers repeat on each page. It has to be inside the page's printable area, so that area runs
//    to the paper's top and bottom edges and the margins are given back as padding repeated on each
//    page (box-decoration-break: clone). Pictures, tables and formatting print as they show.
//  - "box": a page number (PAGE, NUMPAGES), a different first page or different even pages need
//    something drawn per page: CSS page-margin boxes, with counter(page) and counter(pages) and
//    rules for :first and :left (even) pages. They hold text only, so this is for headers and
//    footers that are text (as page numbers mostly are), in a browser that has them.
// Otherwise (several sections with their own headers and footers, or a header with pictures and a
// page number where page-margin boxes aren't to be had), the header and footer print once, at the
// start and the end, as they show: a wrong page number on every page would be worse.

const PAGE_FIELD = /^(PAGE|NUMPAGES|SECTIONPAGES)\b/i;
const NOT_TEXT = 'table, img, svg, canvas, .picture, .textbox, .raw-block, .docx-columns, math';
const BLOCKS = 'p, h1, h2, h3, h4, h5, h6';

/** Whether this browser has CSS page-margin boxes (@top-center and the like). */
export function hasMarginBoxes() {
  return typeof window !== 'undefined' && typeof window.CSSMarginRule !== 'undefined';
}

const quote = (s) => `"${s.replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/\n/g, '\\A ')}"`;

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
let raw = []; // the document's kept XML (doc.raw), for simple fields (w:fldSimple), kept as they are

/** The page-number field `el` is, or null: 'page' or 'pages'. */
function pageField(el) {
  let instr = null;
  if (el.classList?.contains('field')) instr = el.dataset.instr || '';
  else if (el.classList?.contains('raw-inline') && el.dataset.raw != null) {
    const node = raw[Number(el.dataset.raw)];
    if (node?.localName === 'fldSimple') instr = node.getAttributeNS(W, 'instr') || '';
  }
  const m = instr != null && PAGE_FIELD.exec(instr.trim());
  return m ? (m[1].toUpperCase() === 'PAGE' ? 'page' : 'pages') : null;
}

/**
 * A header or footer `host` as page-margin box contents: { left, center, right }, each null or
 * { content (a CSS content value), style (its font and colour) }. A paragraph's alignment puts its
 * text in a box; text after tabs goes to the next box along, as Word's Header and Footer styles set
 * their tab stops at the centre and the right margin. Paragraphs are lines of the boxes.
 */
function boxContents(host) {
  const slots = ['left', 'center', 'right'];
  const lines = { left: [], center: [], right: [] };
  const styles = {};
  for (const block of host.querySelectorAll(BLOCKS)) {
    const align = getComputedStyle(block).textAlign;
    const first = align === 'center' ? 1 : align === 'right' || align === 'end' ? 2 : 0;
    const parts = [[]];
    const walk = (node) => {
      for (const child of node.childNodes) {
        if (child.nodeType === 3) {
          child.data.split('\t').forEach((piece, i) => {
            if (i) parts.push([]);
            if (piece) {
              parts[parts.length - 1].push(quote(piece));
              const at = Math.min(first + parts.length - 1, 2);
              styles[slots[at]] ||= child.parentElement;
            }
          });
        } else if (child.nodeType === 1) {
          if (child.matches('.raw-hidden, .page-break')) continue;
          // Word's absolute tab (w:ptab: to the centre or the right margin), kept as it is.
          const node = child.classList.contains('raw-inline') && child.dataset.raw != null ? raw[Number(child.dataset.raw)] : null;
          if (node && (node.localName === 'ptab' || node.getElementsByTagNameNS?.(W, 'ptab').length)) {
            parts.push([]);
            continue;
          }
          const field = pageField(child);
          if (field) {
            parts[parts.length - 1].push(`counter(${field})`);
            styles[slots[Math.min(first + parts.length - 1, 2)]] ||= child;
          } else walk(child);
        }
      }
    };
    walk(block);
    // Tabs past the right margin's stop stay in the right box.
    while (parts.length > 3 - first && parts.length > 1) {
      const last = parts.pop();
      parts[parts.length - 1].push('" "', ...last);
    }
    slots.forEach((slot, i) => {
      const part = parts[i - first];
      lines[slot].push(part && part.length ? part.join(' ') : '""');
    });
  }
  const out = {};
  for (const slot of slots) {
    if (!lines[slot].some((l) => l !== '""')) { out[slot] = null; continue; }
    const cs = getComputedStyle(styles[slot]);
    out[slot] = {
      content: lines[slot].join(' "\\A " '),
      style: `font-family: ${cs.fontFamily}; font-size: ${cs.fontSize}; font-weight: ${cs.fontWeight}; font-style: ${cs.fontStyle}; color: ${cs.color}; line-height: ${cs.lineHeight};`,
    };
  }
  return out;
}

/**
 * How the header (or footer) prints: { mode: 'none' | 'once' | 'fixed' | 'box', … }. `variants` are
 * doc.headers (doc.footers): { type, section } in the order of their hosts' data-variant.
 */
function plan(page, which, variants, zoom, boxes) {
  if (!variants.length) return { mode: 'none' };
  if (variants.some((v) => v.section > 1)) return { mode: 'once' };
  const host = (type) => {
    const i = variants.findIndex((v) => v.type === type);
    return i < 0 ? null : page.querySelector(`:scope > .docx-hf-band.${which} .docx-hf-part[data-variant="${i}"]`);
  };
  const hosts = { default: host('default'), first: host('first'), even: host('even') };
  const present = Object.values(hosts).filter(Boolean);
  if (!present.length) return { mode: 'once' };
  // Measured where they show (each one shown in turn by the view's tabs isn't needed: the hidden ones
  // are measured by showing them for a moment).
  const height = (el) => {
    const shown = el.getClientRects().length;
    if (!shown) el.style.setProperty('display', 'block', 'important');
    const h = el.getBoundingClientRect().height / zoom;
    if (!shown) el.style.removeProperty('display');
    return h;
  };
  const numbered = present.some((el) => [...el.querySelectorAll('.field, .raw-inline')].some(pageField));
  if (!numbered && hosts.default && !hosts.first && !hosts.even) {
    return { mode: 'fixed', variant: variants.findIndex((v) => v.type === 'default'), height: height(hosts.default) };
  }
  const text = present.every((el) => !el.querySelector(NOT_TEXT) && el.querySelectorAll(BLOCKS).length <= 4);
  if (!boxes || !text) return { mode: 'once' };
  return {
    mode: 'box',
    height: Math.max(...present.map(height)),
    pages: Object.fromEntries(Object.entries(hosts).map(([type, el]) => [type, el ? boxContents(el) : null])),
  };
}

/**
 * The stylesheet that prints `page` (the editor's page) with its headers and footers on every page:
 * the @page rule (paper size and margins) and what goes with it. `doc` gives the page geometry
 * (doc.page) and the header and footer variants; `zoom` is how much bigger the page is shown.
 */
export function printCss(page, doc, zoom = 1, boxes = hasMarginBoxes()) {
  raw = doc.raw || [];
  const p = doc.page;
  const header = plan(page, 'header', inUse(doc, doc.headers), zoom, boxes);
  const footer = plan(page, 'footer', inUse(doc, doc.footers), zoom, boxes);
  // The page's top: the paper's edge with a fixed header (the margin is padding then), the top of a
  // header in page-margin boxes, or the top margin. The text starts at the top margin, or below the
  // header where it is taller, as in Word.
  const edge = (plan_, margin, distance) => {
    if (plan_.mode === 'fixed') return { margin: 0, padding: Math.max(margin, distance + plan_.height) };
    if (plan_.mode === 'box') {
      const m = distance + plan_.height;
      return { margin: m, padding: Math.max(0, margin - m) };
    }
    return { margin, padding: 0 };
  };
  const top = edge(header, p.top, p.header);
  const bottom = edge(footer, p.bottom, p.footer);
  const css = [];
  const marginBoxes = { default: [], first: [], even: [] };
  for (const [which, plan_] of [['header', header], ['footer', footer]]) {
    const band = `.docx-sheet > .docx-page > .docx-hf-band.${which}`;
    if (plan_.mode === 'fixed') {
      css.push(`${band} { position: fixed !important; left: 0 !important; right: 0 !important; ${which === 'header' ? `top: ${p.header}px !important; bottom: auto !important;` : `bottom: ${p.footer}px !important; top: auto !important;`} margin: 0 !important; padding: 0 !important; min-height: 0 !important; display: block !important; }`);
      css.push(`${band} .docx-hf-part { display: none !important; }`);
      css.push(`${band} .docx-hf-part[data-variant="${plan_.variant}"] { display: block !important; }`);
    } else if (plan_.mode === 'box') {
      css.push(`${band} { display: none !important; }`);
      const row = which === 'header' ? 'top' : 'bottom';
      for (const [type, slots] of Object.entries(plan_.pages)) {
        // A page with no header of its own (none for the first page, say) shows none.
        if (!slots && type !== 'default') continue;
        for (const slot of ['left', 'center', 'right']) {
          const box = slots?.[slot];
          const at = `@${row}-${slot}`;
          if (!box) {
            if (type !== 'default' || plan_.pages.first || plan_.pages.even) marginBoxes[type].push(`${at} { content: none; }`);
            continue;
          }
          marginBoxes[type].push(`${at} { content: ${box.content}; white-space: pre; text-align: ${slot}; vertical-align: ${row === 'top' ? 'bottom' : 'top'}; ${box.style} }`);
        }
      }
    }
  }
  const pageRule = `size: ${p.width}px ${p.minHeight}px; margin: ${top.margin}px ${p.right}px ${bottom.margin}px ${p.left}px;`;
  const rules = [`@page { ${pageRule} ${marginBoxes.default.join(' ')} }`];
  if (marginBoxes.first.length) rules.push(`@page :first { ${marginBoxes.first.join(' ')} }`);
  if (marginBoxes.even.length) rules.push(`@page :left { ${marginBoxes.even.join(' ')} }`);
  const media = [];
  if (top.padding || bottom.padding) {
    media.push(`.docx-sheet > .docx-page { padding-top: ${top.padding}px !important; padding-bottom: ${bottom.padding}px !important; box-decoration-break: clone; -webkit-box-decoration-break: clone; }`);
  }
  media.push(...css);
  return `${rules.join('\n')}\n@media print {\n${media.join('\n')}\n}`;
}

/**
 * How the header and footer will print: { header, footer }, each 'none' (there is none), 'every' (on
 * every page) or 'once' (at the start or the end).
 */
export function printModes(page, doc, zoom = 1, boxes = hasMarginBoxes()) {
  raw = doc.raw || [];
  const mode = (which, variants) => {
    const m = plan(page, which, variants, zoom, boxes).mode;
    return m === 'none' ? 'none' : m === 'once' ? 'once' : 'every';
  };
  return { header: mode('header', inUse(doc, doc.headers)), footer: mode('footer', inUse(doc, doc.footers)) };
}

/**
 * The variants as printing uses them: a first page's or even pages' one the document doesn't use
 * (no different first page, or odd and even pages alike) is kept in its place, but as no type.
 */
function inUse(doc, variants) {
  return variants.map((v) => (doc.variantShown && !doc.variantShown(v) ? { ...v, type: 'unused' } : v));
}

/** The plain @page rule (paper size and margins), before anything is measured. */
export function basicPrintCss(p) {
  return `@page { size: ${p.width}px ${p.minHeight}px; margin: ${p.top}px ${p.right}px ${p.bottom}px ${p.left}px; }`;
}
