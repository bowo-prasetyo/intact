// Where the pages would break, about: the page here is one long sheet (Word decides the real page
// breaks when it opens the file, from its own fonts and layout), but people sending a CV, a letter or
// a form need to know whether it still fits on one page or two. So the text is cut into pages the
// way Word fills them: line by line, each page holding the paper's height less its margins (and a
// header or footer taller than its margin), with Word's widow and orphan control and its headings
// kept with the paragraph after them, and a new page at each page break, "page break before" and
// section break that starts a new page. Like the line numbers and the formatting marks, the result
// is drawn in a layer over the page, never in it, so it can't change the document.

const BLOCKS = 'p, h1, h2, h3, h4, h5, h6';
const SKIP = '.docx-notes, .docx-hf-band, .textbox, .raw-block';
const HEADING = /^H[1-6]$/;

/**
 * The lines of `block` (outside its tables): [{ top, bottom }] in client coordinates, top to bottom,
 * as line boxes: the text's own rectangles are only as tall as its font, so the first line starts at
 * the top of the paragraph's content, the last ends at its bottom, and lines meet halfway between.
 * `zoom` is how much bigger than its size the page is shown.
 */
function linesOf(block, zoom) {
  const out = [];
  const add = (r) => {
    if (!r.height) return;
    const line = out.find((l) => Math.abs(l.top - r.top) < r.height / 2 || (r.top < l.bottom - 2 && r.bottom > l.top + 2));
    if (line) {
      line.top = Math.min(line.top, r.top);
      line.bottom = Math.max(line.bottom, r.bottom);
    } else out.push({ top: r.top, bottom: r.bottom });
  };
  const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT, {
    acceptNode: (n) => (n.nodeType === 1 && n.matches('table, .page-break, .raw-hidden, .floating, .textbox') ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT),
  });
  const range = document.createRange();
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    if (n.nodeType === 3) {
      if (!n.data.replace(/\n/g, '')) continue;
      range.selectNodeContents(n);
      for (const r of range.getClientRects()) add(r);
    } else if (n.matches('img, .picture, .raw-inline, br')) {
      const r = n.getBoundingClientRect();
      if (n.tagName !== 'BR' || !out.length) add(r);
    }
  }
  const r = block.getBoundingClientRect();
  const cs = getComputedStyle(block);
  const top = r.top + (parseFloat(cs.borderTopWidth) + parseFloat(cs.paddingTop)) * zoom;
  const bottom = r.bottom - (parseFloat(cs.borderBottomWidth) + parseFloat(cs.paddingBottom)) * zoom;
  if (!out.length) {
    // An empty paragraph is still a line of the page, as in Word.
    if (bottom > top) out.push({ top, bottom });
    return out;
  }
  out.sort((a, b) => a.top - b.top);
  for (let i = 1; i < out.length; i++) {
    const mid = (out[i - 1].bottom + out[i].top) / 2;
    if (mid > out[i - 1].top && mid < out[i].bottom) out[i - 1].bottom = out[i].top = mid;
  }
  if (top <= out[0].top) out[0].top = top;
  if (bottom >= out[out.length - 1].bottom) out[out.length - 1].bottom = bottom;
  return out;
}

/**
 * The units the pages are filled with, in order: the lines of paragraphs, and table rows (kept
 * whole, as most rows in a CV or a form are short), each { top, bottom, block, index, count }, and
 * the places a new page is forced (client y: a unit starting below one starts a page).
 */
function unitsOf(page, zoom) {
  const units = [];
  const forced = [];
  const rows = new Set();
  for (const el of page.querySelectorAll(UNIT_SELECTOR)) {
    if (el.parentElement.closest(SKIP) || el.closest('.docx-notes, .docx-hf-band, .textbox')) continue;
    if (el.tagName === 'TR') {
      // A row of a nested table belongs to its outer row.
      if (el.parentElement.closest('tr')) continue;
      const r = el.getBoundingClientRect();
      if (r.height) units.push({ top: r.top, bottom: r.bottom, block: el, index: 0, count: 1 });
      rows.add(el);
      continue;
    }
    if (el.closest('tr')) {
      if (el.dataset.pageBreakBefore != null) forced.push(el.closest('tr').getBoundingClientRect().top - 0.5);
      continue; // in a row, which goes whole
    }
    if (!el.getClientRects().length) continue;
    if (el.matches('.raw-block')) {
      const r = el.getBoundingClientRect();
      if (r.height) units.push({ top: r.top, bottom: r.bottom, block: el, index: 0, count: 1 });
      continue;
    }
    const lines = linesOf(el, zoom);
    lines.forEach((l, i) => units.push({ ...l, block: el, index: i, count: lines.length }));
    if (el.dataset.pageBreakBefore != null && lines.length) forced.push(lines[0].top - 0.5);
    for (const pb of el.querySelectorAll('.page-break')) forced.push(pb.getBoundingClientRect().bottom - 0.5);
    const end = el.dataset.sectionEnd;
    if (end && end !== 'continuous' && lines.length) forced.push(lines[lines.length - 1].bottom - 0.5);
  }
  // In document order, which in columns isn't top to bottom: by where they start, side by side
  // lines of columns counting as one line of the page.
  units.sort((a, b) => a.top - b.top);
  return { units, forced: forced.sort((a, b) => a - b) };
}

/** The elements the pages are filled with, in the page's order: for finding one in a copy. */
export const UNIT_SELECTOR = `${BLOCKS}, tr, .raw-block`;

/**
 * The pages of `page` (the editable page element), about: { count, breaks }, where `breaks` are
 * where pages 2, 3… begin: { y, block }, the client y position and the element (paragraph, row) in
 * which it is. `height` is the text's height on a page in client pixels (the paper less its margins,
 * as shown at the current `zoom`).
 */
export function estimatePages(page, height, zoom = 1) {
  const { units, forced } = unitsOf(page, zoom);
  const breaks = [];
  if (!units.length || !(height > 0)) return { count: 1, breaks };
  let start = units[0].top;
  let f = 0;
  let prevBottom = -Infinity;
  for (let i = 0; i < units.length; i++) {
    const u = units[i];
    // A page break, "page break before" or a section break between the last unit and this one.
    let hard = false;
    while (f < forced.length && forced[f] < u.top + 0.5) {
      if (forced[f] >= prevBottom - 1 && i > 0) hard = true;
      f++;
    }
    if (hard) {
      start = u.top;
      breaks.push({ y: u.top, block: u.block });
    } else if (u.bottom - start > height + 0.05 && u.top > start + 0.5) {
      // Doesn't fit: the page ends before this unit, or a little earlier so as not to leave a
      // paragraph's first line alone at the bottom (orphan) or its last alone at the top (widow),
      // and not to leave a heading alone at the bottom, as Word's widow control and its heading
      // styles' "keep with next" do. Never back to the page's own start.
      let j = i;
      let k = u.index;
      const same = (at) => units[at]?.block === u.block;
      if (u.count > 1) {
        if (u.count - k === 1 && same(j - 1)) { j--; k--; } // widow: a line goes with the last
        if (k === 1 && same(j - 1)) { j--; k = 0; } // orphan: the first line goes too
      }
      // A heading just before the page's first paragraph goes to the next page with it.
      const prev = units[j - 1];
      if (k === 0 && prev && HEADING.test(prev.block.tagName) && prev.index === prev.count - 1 && units[j - prev.count]?.block === prev.block) j -= prev.count;
      if (units[j].top <= start + 0.5) j = i;
      start = units[j].top;
      breaks.push({ y: start, block: units[j].block });
      i = j; // go on from the first unit of the new page
      prevBottom = units[j].bottom;
      continue;
    }
    // A unit taller than a page (a big picture or a long row): Word carries on on the next pages.
    while (u.bottom - start > height + 0.05) {
      start += height;
      breaks.push({ y: start, block: u.block });
    }
    prevBottom = u.bottom;
  }
  return { count: breaks.length + 1, breaks };
}

/**
 * Draws the page breaks into `layer` (an empty, absolutely positioned element over the sheet): a
 * dashed line across the sheet with the number of the page that begins there. `zoom` is how much
 * bigger than its size the page is shown; the marks are placed in its own pixels.
 */
export function drawPageBreaks(layer, breaks, zoom = 1) {
  layer.replaceChildren();
  const base = layer.getBoundingClientRect();
  const fragment = document.createDocumentFragment();
  breaks.forEach((y, i) => {
    const mark = document.createElement('div');
    mark.className = 'page-estimate';
    // Just above the first line of the new page, in the gap between the pages' lines.
    mark.style.top = `${(y - base.top) / zoom - 2}px`;
    const label = document.createElement('span');
    label.textContent = `Page ${i + 2}`;
    mark.append(label);
    fragment.append(mark);
  });
  layer.append(fragment);
}
