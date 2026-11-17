// Line numbers in the margin, as Word shows them for a section with line numbering (w:lnNumType):
// legal and legislative drafts, contracts and transcripts cite "page 3, line 12". Like the formatting
// marks, they are drawn in a layer over the page, never in it, so they can't change the document.
//
// Word numbers every line of the text in the section, except in tables, notes, headers, footers,
// text boxes and paragraphs that suppress them (w:suppressLineNumbers), and shows every countBy-th
// number. It restarts the count on each page, each section, or never (w:restart). There are no pages
// here, so "each page" restarts where Word last broke the pages when it saved the file (its
// w:lastRenderedPageBreak marks) and at page breaks: the numbers are the ones on Word's printed
// pages, as long as the edits don't move the page breaks.

const BLOCKS = 'p, h1, h2, h3, h4, h5, h6';
const SKIP = 'table, .docx-notes, .docx-hf-band, .textbox, .raw-block';

/** The lines of `block`: [{ top, bottom, height }] in client coordinates, top to bottom. */
function lines(block) {
  const out = [];
  const add = (r) => {
    if (!r.height) return;
    const line = out.find((l) => Math.abs(l.top - r.top) < r.height / 2 || (r.top < l.bottom - 2 && r.bottom > l.top + 2));
    if (line) {
      line.top = Math.min(line.top, r.top);
      line.bottom = Math.max(line.bottom, r.bottom);
    } else out.push({ top: r.top, bottom: r.bottom });
  };
  const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT);
  const range = document.createRange();
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    if (n.nodeType === 3) {
      if (!n.data.replace(/\n/g, '')) continue;
      range.selectNodeContents(n);
      for (const r of range.getClientRects()) add(r);
    } else if (n.matches('img, .picture:not(.floating), .raw-inline:not(.raw-hidden), br')) {
      const r = n.getBoundingClientRect();
      if (n.tagName !== 'BR' || !out.length) add(r);
    }
  }
  if (!out.length) {
    const r = block.getBoundingClientRect();
    if (r.height) out.push({ top: r.top, bottom: r.top + Math.min(r.height, parseFloat(getComputedStyle(block).lineHeight) || r.height) });
  }
  return out.sort((a, b) => a.top - b.top).map((l) => ({ ...l, height: l.bottom - l.top }));
}

/** The client top of the character at `offset` in `block`'s text, or null. */
function topAt(block, offset) {
  const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
  let n = 0;
  for (let t = walker.nextNode(); t; t = walker.nextNode()) {
    if (offset < n + t.data.length) {
      const range = document.createRange();
      range.setStart(t, offset - n);
      range.setEnd(t, offset - n + 1);
      return range.getClientRects()[0]?.top ?? null;
    }
    n += t.data.length;
  }
  return null;
}

/**
 * Draws the line numbers of `page` into `layer` (an empty, absolutely positioned element over the
 * sheet). `sections` is the document's sections in order, each { countBy, start, distance (px),
 * restart } or null for a section without line numbers.
 */
export function drawLineNumbers(page, layer, sections, zoom = 1) {
  // `zoom`: how much bigger than its size the page is shown; numbers are placed in its own pixels.
  layer.replaceChildren();
  if (!sections.some(Boolean)) return;
  const base = layer.getBoundingClientRect();
  const style = getComputedStyle(page);
  const textLeft = page.getBoundingClientRect().left + parseFloat(style.paddingLeft) * zoom;
  let section = 0;
  let settings = sections[0];
  let n = settings ? settings.start : 0;
  const fragment = document.createDocumentFragment();
  for (const block of page.querySelectorAll(BLOCKS)) {
    if (block.parentElement.closest(SKIP) || !block.getClientRects().length) continue;
    if (settings) {
      const newPage = settings.restart === 'newPage';
      // Where Word's pages began within this paragraph: restart before the line holding them.
      const breaks = [];
      if (newPage && block.dataset.pageBreakBefore != null) breaks.push(-Infinity);
      if (newPage) {
        for (const at of (block.dataset.pageBreaks || '').split(',').filter(Boolean)) {
          const top = topAt(block, Number(at));
          if (top != null) breaks.push(top - 1);
        }
        for (const pb of block.querySelectorAll('.page-break')) breaks.push(pb.getBoundingClientRect().bottom - 1);
      }
      const suppressed = block.dataset.noLineNumbers != null;
      const font = getComputedStyle(block);
      for (const line of lines(block)) {
        while (breaks.length && breaks.some((b) => b < line.top)) {
          breaks.splice(breaks.findIndex((b) => b < line.top), 1);
          n = settings.start;
        }
        if (suppressed) continue;
        n++;
        if (n % settings.countBy) continue;
        const num = document.createElement('span');
        num.className = 'line-number';
        num.textContent = String(n);
        const h = line.height / zoom;
        num.style.cssText = `top: ${(line.top - base.top) / zoom}px; height: ${h}px; line-height: ${h}px; `
          + `right: ${(base.right - textLeft) / zoom + settings.distance}px; font-family: ${font.fontFamily};`;
        fragment.append(num);
      }
    }
    if (block.dataset.sectionEnd != null) {
      section++;
      settings = sections[section] || null;
      if (settings && settings.restart !== 'continuous') n = settings.start;
    }
  }
  layer.append(fragment);
}
