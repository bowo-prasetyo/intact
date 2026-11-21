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
// A document of several sections prints each on its own paper, with its own header, footer and
// page numbers (see "a document of several sections" below). Otherwise (a header with pictures and
// a page number, or page-margin boxes not to be had), the header and footer print once, at the
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
let numbering = { numFmt: 'decimal', numStart: null }; // the (final) section's page numbering

// Word's page number formats (w:pgNumType's w:fmt, a field's \* switch) as CSS counter styles.
const COUNTER_STYLES = {
  decimal: 'decimal', numberInDash: 'decimal', lowerRoman: 'lower-roman', upperRoman: 'upper-roman', lowerLetter: 'lower-alpha', upperLetter: 'upper-alpha',
  roman: 'lower-roman', Roman: 'upper-roman', ROMAN: 'upper-roman', alphabetic: 'lower-alpha', ALPHABETIC: 'upper-alpha', Arabic: 'decimal', arabic: 'decimal',
};

/**
 * A page-number field as a CSS content value: counter(page) in the section's number format (i, ii,
 * iii; - 1 -), or as the field's own format switch (PAGE \* roman) says; counter(pages) for NUMPAGES.
 */
function counterContent(field, instr) {
  if (field !== 'page') return 'counter(pages)';
  const own = /\\\*\s*(\w+)/.exec(instr || '')?.[1];
  const fmt = own && COUNTER_STYLES[own] ? own : numbering.numFmt || 'decimal';
  const style = COUNTER_STYLES[fmt] || 'decimal';
  const counter = style === 'decimal' ? 'counter(page)' : `counter(page, ${style})`;
  return fmt === 'numberInDash' ? `"- " ${counter} " -"` : counter;
}

/** The instruction of field `el` (a complex field's, a simple one's kept as it is, or one without a result), or null. */
function fieldInstr(el) {
  if (el.classList?.contains('field') || el.classList?.contains('page-field-bare')) return el.dataset.instr || '';
  if (el.classList?.contains('raw-inline') && el.dataset.raw != null) {
    const node = raw[Number(el.dataset.raw)];
    if (node?.localName === 'fldSimple') return node.getAttributeNS(W, 'instr') || '';
  }
  return null;
}

/** The page-number field `el` is, or null: 'page', 'pages' or 'sectionpages'. */
function pageField(el) {
  const instr = fieldInstr(el);
  const m = instr != null && PAGE_FIELD.exec(instr.trim());
  return m ? { PAGE: 'page', NUMPAGES: 'pages', SECTIONPAGES: 'sectionpages' }[m[1].toUpperCase()] : null;
}

/**
 * A header or footer `host` as page-margin box contents: { left, center, right }, each null or
 * { lines, style (its font and colour) }, where a line is a list of tokens: { text } or { field
 * ('page', 'pages' or 'sectionpages'), instr }. A paragraph's alignment puts its text in a box;
 * text after tabs goes to the next box along, as Word's Header and Footer styles set their tab
 * stops at the centre and the right margin. Paragraphs are lines of the boxes.
 */
function slotTokens(host) {
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
              parts[parts.length - 1].push({ text: piece });
              const at = Math.min(first + parts.length - 1, 2);
              styles[slots[at]] ||= child.parentElement;
            }
          });
        } else if (child.nodeType === 1) {
          const field = pageField(child);
          if (field) {
            parts[parts.length - 1].push({ field, instr: fieldInstr(child) });
            styles[slots[Math.min(first + parts.length - 1, 2)]] ||= child;
            continue;
          }
          if (child.matches('.raw-hidden, .page-break')) continue;
          // Word's absolute tab (w:ptab: to the centre or the right margin), kept as it is.
          const node = child.classList.contains('raw-inline') && child.dataset.raw != null ? raw[Number(child.dataset.raw)] : null;
          if (node && (node.localName === 'ptab' || node.getElementsByTagNameNS?.(W, 'ptab').length)) {
            parts.push([]);
            continue;
          }
          walk(child);
        }
      }
    };
    walk(block);
    // Tabs past the right margin's stop stay in the right box.
    while (parts.length > 3 - first && parts.length > 1) {
      const last = parts.pop();
      parts[parts.length - 1].push({ text: ' ' }, ...last);
    }
    slots.forEach((slot, i) => lines[slot].push(parts[i - first] || []));
  }
  const out = {};
  for (const slot of slots) {
    if (!lines[slot].some((l) => l.length)) { out[slot] = null; continue; }
    const cs = getComputedStyle(styles[slot]);
    out[slot] = {
      lines: lines[slot],
      style: `font-family: ${cs.fontFamily}; font-size: ${cs.fontSize}; font-weight: ${cs.fontWeight}; font-style: ${cs.fontStyle}; color: ${cs.color}; line-height: ${cs.lineHeight};`,
    };
  }
  return out;
}

/** A header or footer `host` as page-margin box contents, as CSS content values (see slotTokens). */
function boxContents(host) {
  const out = slotTokens(host);
  for (const [slot, box] of Object.entries(out)) {
    if (!box) continue;
    const line = (tokens) => (tokens.length ? tokens.map((t) => (t.field ? counterContent(t.field === 'page' ? 'page' : 'pages', t.instr) : quote(t.text))).join(' ') : '""');
    out[slot] = { content: box.lines.map(line).join(' "\\A " '), style: box.style };
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
  if (severalSections(page, doc)) return sectionsCss(page, doc, zoom, boxes).css;
  numbering = { numFmt: 'decimal', numStart: null, ...(doc.pageSetupNow?.() || {}) };
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
  // Numbering that starts at another number (w:start): the first page is set to it, the next ones
  // count on from there.
  const reset = numbering.numStart != null && numbering.numStart !== 1 && (header.mode === 'box' || footer.mode === 'box')
    ? ` counter-set: page ${numbering.numStart};` : '';
  const rules = [`@page { ${pageRule} ${marginBoxes.default.join(' ')} }`];
  if (reset) rules.push(`@page :first {${reset} }`);
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
  if (severalSections(page, doc)) {
    const { header, footer, reasons } = sectionsCss(page, doc, zoom, boxes);
    return { header, footer, why: ONCE_REASONS[reasons.header || reasons.footer] || '' };
  }
  const plans = { header: plan(page, 'header', inUse(doc, doc.headers), zoom, boxes), footer: plan(page, 'footer', inUse(doc, doc.footers), zoom, boxes) };
  const mode = (m) => (m === 'none' ? 'none' : m === 'once' ? 'once' : 'every');
  const once = Object.values(plans).some((p) => p.mode === 'once');
  // Once: headers of sections after the first that this path doesn't print section by section (a
  // document whose later sections all begin on the same page), else pictures or the browser.
  const why = !once ? '' : [...doc.headers, ...doc.footers].some((v) => v.section > 1) ? ONCE_REASONS.continuous : boxes ? ONCE_REASONS.pictures : ONCE_REASONS.browser;
  return { header: mode(plans.header.mode), footer: mode(plans.footer.mode), why };
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

// ---------------------------------------------------------------- a document of several sections
//
// Each section prints on its own paper (size, orientation, margins) with its own header and footer
// and its own page numbers, as Word prints it. Each run of pages a section begins (a new page; a
// section that begins on the same page, continuous, joins the one before) is a CSS named page,
// `page: intact-s<n>`, with an @page rule of its own: browsers lay the text of each page out at its
// paper's width and start a new page where the name changes. Its header and footer go in that
// rule's page-margin boxes, as for one section.
//
// Page numbers: counter(page) counts every page of the printout, so each run of numbering (from the
// start, or from a section that starts its numbering at a number) has a counter of its own, that
// the pages of its sections step on. Counters set within an @page rule last for that page only,
// so a run's start is added on the first page instead (as a step), where every counter is.
//
// A section's different first page: CSS can name the document's first page (:first), not a later
// section's. Its first page is told by its number instead, which is known where the section starts
// its numbering at a number: its header and footer are written as counters in counter styles that
// show the first page's text for that number (range S S), and the other pages' for every other.
// Where a later section with a different first page numbers on from the one before, which page is
// its first isn't known: then the headers and footers print once, as they show (see reasons).

const NAME = 'intact-s';
const RUN = 'intact-run';

/** Why a header or footer prints once rather than on every page, in words for the PDF dialog. */
export const ONCE_REASONS = {
  browser: 'printing it on every page needs a browser that prints page-margin text, such as Chrome or Edge.',
  pictures: 'it has pictures or a table beside a page number (or differs from section to section), which only Word puts on every page.',
  continuing: 'a section after the first has a different first page and numbers its pages on from the section before it, so which page is its first is Word’s to work out. Start its page numbers at a number (Page number › Format page numbers) to print it on every page.',
  sectionPages: 'it shows the number of pages in its section, which Word counts.',
  continuous: 'a section that begins on the same page as the one before has a header or footer of its own, which only Word puts on the right pages.',
  firstPages: 'a section after the first has a different first page and shows the number of pages, which together only Word puts on the right pages.',
};

/** A page number as text in Word's format `fmt` (for the dashes of - 1 -) and its CSS counter style. */
function styleOf(fmt, instr) {
  const own = /\\\*\s*(\w+)/.exec(instr || '')?.[1];
  const f = own && COUNTER_STYLES[own] ? own : fmt || 'decimal';
  return { style: COUNTER_STYLES[f] || 'decimal', dashes: f === 'numberInDash' };
}

/**
 * The sections of `doc` as they print: the runs of pages, each { name, sections, setup (the first
 * section's), bands: { header, footer } ({ default, first, even } variant indexes), titlePg, run
 * (its numbering run), start (that run's first number, where this group starts it, else null) },
 * and which run of the page's children each is: { groups, ranges: [[from, to, group]] } (1-based).
 */
function groupsOf(page, doc) {
  const sections = doc.printSections(page);
  const groups = [];
  const groupOf = [];
  let run = -1;
  sections.forEach((sec, i) => {
    if (i && sec.start === 'continuous') {
      groups[groups.length - 1].sections.push(sec);
    } else {
      const restart = !i || sec.setup.numStart != null;
      if (restart) run++;
      groups.push({
        name: `${NAME}${groups.length}`, sections: [sec], setup: sec.setup, bands: { header: sec.header, footer: sec.footer },
        headerDist: sec.headerDist, footerDist: sec.footerDist, titlePg: sec.titlePg,
        run, start: restart ? sec.setup.numStart ?? 1 : null,
      });
    }
    groupOf.push(groups.length - 1);
  });
  // The page's children, each in its section's group: a child holding a section's end is that
  // section's last; the header before them is the first group's, the notes and footer the last's.
  const kids = [...page.children];
  const last = groups.length - 1;
  let sec = 0;
  const of = kids.map((kid) => {
    if (kid.matches('.docx-hf-band.header')) return 0;
    if (kid.matches('.docx-hf-band.footer, .docx-notes')) return last;
    const g = groupOf[Math.min(sec, groupOf.length - 1)];
    const ends = [...(kid.matches('[data-section-end]') ? [kid] : []), ...kid.querySelectorAll('[data-section-end]')]
      .filter((el) => !el.closest('.docx-hf-band, .docx-notes, .textbox'));
    sec += ends.length;
    return g;
  });
  const ranges = [];
  of.forEach((g, i) => {
    const r = ranges[ranges.length - 1];
    if (r && r[2] === g) r[1] = i + 1;
    else ranges.push([i + 1, i + 1, g]);
  });
  return { groups, ranges };
}

/**
 * How the header (or footer) prints in a document of several sections: { mode: 'none' | 'once' |
 * 'box', reason (when once, see ONCE_REASONS), height (the tallest, per group), slots (per group:
 * { default, first, even }, each slotTokens' or null for none) }.
 */
function sectionsPlan(page, which, doc, groups, zoom, boxes, evenAndOdd) {
  const list = which === 'header' ? doc.headers : doc.footers;
  if (!list.length) return { mode: 'none' };
  const host = (i) => (i == null ? null : page.querySelector(`:scope > .docx-hf-band.${which} .docx-hf-part[data-variant="${i}"]`));
  const height = (el) => {
    const shown = el.getClientRects().length;
    if (!shown) el.style.setProperty('display', 'block', 'important');
    const h = el.getBoundingClientRect().height / zoom;
    if (!shown) el.style.removeProperty('display');
    return h;
  };
  const heights = [];
  const slots = [];
  for (const g of groups) {
    const ids = g.bands[which];
    // Word's: the first page's where the section has a different first page (none if it has
    // none of its own or before it), the even pages' where odd and even pages differ.
    const hosts = { default: host(ids.default), first: g.titlePg ? host(ids.first) : undefined, even: evenAndOdd ? host(ids.even) : undefined };
    const present = Object.values(hosts).filter(Boolean);
    if (present.some((el) => el.querySelector(NOT_TEXT) || el.querySelectorAll(BLOCKS).length > 4)) return { mode: 'once', reason: 'pictures' };
    const tokens = Object.fromEntries(Object.entries(hosts).map(([type, el]) => [type, el === undefined ? undefined : el ? slotTokens(el) : null]));
    const fields = JSON.stringify(tokens);
    if (fields.includes('"field":"sectionpages"')) return { mode: 'once', reason: 'sectionPages' };
    heights.push(present.length ? Math.max(...present.map(height)) : 0);
    slots.push(tokens);
  }
  if (!boxes) return { mode: 'once', reason: 'browser' };
  // A later group's different first page, told by its number: only where its numbering starts.
  for (const [i, g] of groups.entries()) {
    if (!i || slots[i].first === undefined) continue;
    const same = JSON.stringify(slots[i].first) === JSON.stringify(slots[i].default)
      && (slots[i].even === undefined || JSON.stringify(slots[i].first) === JSON.stringify(slots[i].even));
    if (same) slots[i].first = undefined;
    else if (g.start == null) return { mode: 'once', reason: 'continuing' };
    else if (JSON.stringify([slots[i].first, slots[i].default, slots[i].even]).includes('"field":"pages"')) return { mode: 'once', reason: 'firstPages' };
  }
  return { mode: 'box', heights, slots };
}

/**
 * The print stylesheet for a document of several sections (see above), and how its header and
 * footer print: { css, header, footer ('none', 'every' or 'once'), reasons: { header, footer } }.
 */
function sectionsCss(page, doc, zoom, boxes) {
  const { groups, ranges } = groupsOf(page, doc);
  const evenAndOdd = !!doc.headerFooterFlags?.().evenAndOdd;
  const plans = {
    header: sectionsPlan(page, 'header', doc, groups, zoom, boxes, evenAndOdd),
    footer: sectionsPlan(page, 'footer', doc, groups, zoom, boxes, evenAndOdd),
  };
  const counterStyles = ['@counter-style intact-blank { system: cyclic; symbols: ""; }'];
  let n = 0;
  /** A counter style showing `symbols` (or extending `base`) only for numbers in `range`. */
  const only = (range, { symbols, base }) => {
    const name = `intact-c${n++}`;
    counterStyles.push(`@counter-style ${name} { ${base ? `system: extends ${base};` : `system: cyclic; symbols: ${symbols};`} range: ${range}; fallback: intact-blank; }`);
    return name;
  };
  const runCounter = (g) => `${RUN}${g.run}`;
  /** Tokens as CSS content, for group `g`; `range` limits them to the pages numbered so (a condition). */
  const content = (lines, g, range) => {
    const parts = [];
    const text = (s) => parts.push(range ? `counter(${runCounter(g)}, ${only(range, { symbols: quote(s) })})` : quote(s));
    lines.forEach((tokens, i) => {
      if (i) text('\n');
      for (const t of tokens) {
        if (!t.field) { text(t.text); continue; }
        if (t.field === 'pages') { parts.push('counter(pages)'); continue; }
        const { style, dashes } = styleOf(g.setup.numFmt, t.instr);
        if (dashes) text('- ');
        if (range) parts.push(`counter(${runCounter(g)}, ${only(range, { base: style })})`);
        else parts.push(style === 'decimal' ? `counter(${runCounter(g)})` : `counter(${runCounter(g)}, ${style})`);
        if (dashes) text(' -');
      }
    });
    return parts.length ? parts.join(' ') : '""';
  };
  /** The margin boxes of one rule: `kinds` the slots shown, each with the range it shows on. */
  const boxesFor = (which, g, kinds) => {
    const row = which === 'header' ? 'top' : 'bottom';
    const dist = which === 'header' ? g.headerDist : g.footerDist;
    return ['left', 'center', 'right'].map((slot) => {
      const shown = kinds.filter(([tokens]) => tokens?.[slot]);
      if (!shown.length) return `@${row}-${slot} { content: none; }`;
      const value = shown.map(([tokens, range]) => content(tokens[slot].lines, g, range)).join(' ');
      return `@${row}-${slot} { content: ${value}; white-space: pre; text-align: ${slot}; vertical-align: ${row === 'top' ? 'top' : 'bottom'}; padding-${row}: ${dist}px; ${shown[0][0][slot].style} }`;
    }).join(' ');
  };
  const rules = [];
  const media = [];
  groups.forEach((g, i) => {
    const s = g.setup;
    const px = (twips) => Number(twips) / 15;
    const top = plans.header.mode === 'box' ? Math.max(px(s.top), g.headerDist + plans.header.heights[i]) : px(s.top);
    const bottom = plans.footer.mode === 'box' ? Math.max(px(s.bottom), g.footerDist + plans.footer.heights[i]) : px(s.bottom);
    const size = `size: ${px(s.w)}px ${px(s.h)}px; margin: ${top}px ${px(s.right)}px ${bottom}px ${px(s.left)}px;`;
    if (!i) rules.push(`@page { ${size} }`);
    const kinds = { default: [], first: [], even: [] };
    for (const which of ['header', 'footer']) {
      const plan_ = plans[which];
      if (plan_.mode !== 'box') continue;
      const t = plan_.slots[i];
      const S = g.start;
      if (i && t.first !== undefined) {
        // Its first page by its number S: the first page's text there, the others' elsewhere.
        const not = `infinite ${S - 1}, ${S + 1} infinite`;
        kinds.default.push(boxesFor(which, g, [[t.first, `${S} ${S}`], [t.default, not]]));
        if (t.even !== undefined) kinds.even.push(boxesFor(which, g, [[t.first, `${S} ${S}`], [t.even, not]]));
      } else {
        kinds.default.push(boxesFor(which, g, [[t.default]]));
        if (t.even !== undefined) kinds.even.push(boxesFor(which, g, [[t.even]]));
        if (!i && t.first !== undefined) kinds.first.push(boxesFor(which, g, [[t.first]]));
      }
    }
    rules.push(`@page ${g.name} { ${size} counter-increment: ${runCounter(g)}; ${kinds.default.join(' ')} }`);
    if (kinds.even.length) rules.push(`@page ${g.name}:left { ${kinds.even.join(' ')} }`);
    if (!i) {
      // The document's first page: every run's start added (the first run's with this page's step).
      const runs = groups.filter((x) => x.start != null).map((x) => `${runCounter(x)} ${x.run === g.run ? x.start : x.start - 1}`);
      rules.push(`@page ${g.name}:first { counter-increment: ${runs.join(' ')}; ${kinds.first.join(' ')} }`);
    }
  });
  for (const [from, to, g] of ranges) media.push(`.docx-sheet > .docx-page > :nth-child(n+${from}):nth-child(-n+${to}) { page: ${groups[g].name}; }`);
  for (const which of ['header', 'footer']) if (plans[which].mode === 'box') media.push(`.docx-sheet > .docx-page > .docx-hf-band.${which} { display: none !important; }`);
  const mode = (which) => ({ none: 'none', once: 'once', box: 'every' })[plans[which].mode];
  return {
    css: `${counterStyles.join('\n')}\n${rules.join('\n')}\n@media print {\n${media.join('\n')}\n}`,
    header: mode('header'),
    footer: mode('footer'),
    reasons: { header: plans.header.reason, footer: plans.footer.reason },
  };
}

/** Whether `doc` prints as several sections: more than one run of pages, each on its own paper. */
function severalSections(page, doc) {
  if (!page || !doc.printSections) return false;
  return doc.printSections(page).slice(1).some((s) => s.start !== 'continuous');
}
