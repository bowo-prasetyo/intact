// "Show formatting marks" (¶), as in Word: · for a space, ° for a non-breaking space, → for a tab,
// ↵ for a line break and ¶ at the end of each paragraph. The marks are drawn in a layer over the
// page, never in it, so showing them can't change the document, its layout or what Undo records.
// Only paragraphs near the visible part of the page are drawn, so long documents stay quick.

const BLOCKS = 'p, h1, h2, h3, h4, h5, h6, li';
const CHAR_MARKS = { ' ': '·', ' ': '°', '\t': '→' };
const MARGIN = 400; // px above and below the visible area that are drawn too

const visible = (el) => el.getClientRects().length > 0;

/** True if a <br> is the browser's placeholder at the end of its paragraph, not a line break. */
function trailing(br, block) {
  for (let n = br; n && n !== block; n = n.parentNode) {
    let s = n.nextSibling;
    while (s && s.nodeType === 3 && !s.data) s = s.nextSibling;
    if (s) return false;
  }
  return true;
}

/** The rectangle of the last character (or item) in a paragraph, for its ¶; null if it is empty. */
function lastRect(block) {
  const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT);
  let last = null;
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    if (n.nodeType === 3 ? n.data.length : n.matches('img, .raw-inline:not(.raw-hidden)')) last = n;
  }
  if (!last) return null;
  if (last.nodeType === 1) return last.getBoundingClientRect();
  const range = document.createRange();
  let end = last.length;
  while (end > 0 && last.data[end - 1] === '\n') end--;
  range.setStart(last, Math.max(0, end - 1));
  range.setEnd(last, Math.max(1, end));
  const rects = range.getClientRects();
  return rects[rects.length - 1] || null;
}

/**
 * Draws the formatting marks of `page` into `layer` (an empty, absolutely positioned element over
 * the page), for the paragraphs within `view` (a DOMRect of the visible area) and a margin.
 */
export function drawMarks(page, layer, view, zoom = 1) {
  // `zoom`: how much bigger than its size the page is shown; marks are placed in its own pixels.
  const base = layer.getBoundingClientRect();
  const marks = [];
  const add = (char, rect, className, x = rect.left) => {
    marks.push({ char, className, x: (x - base.left) / zoom, y: (rect.top - base.top) / zoom, h: rect.height / zoom });
  };
  for (const block of page.querySelectorAll(BLOCKS)) {
    const box = block.getBoundingClientRect();
    if (!visible(block) || box.bottom < view.top - MARGIN || box.top > view.bottom + MARGIN) continue;
    const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT);
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      if (n.nodeType === 1) {
        if (n.tagName === 'BR' && !trailing(n, block)) {
          const range = document.createRange();
          range.setStartBefore(n);
          const rect = range.getClientRects()[0] || n.getBoundingClientRect();
          if (rect.height) add('↵', rect, 'break');
        }
        continue;
      }
      if (n.parentElement.closest(BLOCKS) !== block) continue; // a nested paragraph draws its own
      for (let i = 0; i < n.data.length; i++) {
        const char = CHAR_MARKS[n.data[i]];
        if (!char) continue;
        const range = document.createRange();
        range.setStart(n, i);
        range.setEnd(n, i + 1);
        const rect = range.getBoundingClientRect();
        if (!rect.width && !rect.height) continue;
        if (n.data[i] === '\t') add(char, rect, 'tab');
        else add(char, rect, 'space', rect.left + rect.width / 2);
      }
    }
    const end = lastRect(block);
    if (end) {
      add('¶', end, 'para', end.right + 1);
    } else {
      // An empty paragraph: its ¶ sits where the text would start.
      const style = getComputedStyle(block);
      const css = (v) => (parseFloat(v) || 0) * zoom; // computed lengths are the page's own: shown zoomed
      add('¶', { left: box.left + css(style.paddingLeft) + css(style.textIndent), top: box.top + css(style.paddingTop), height: css(style.lineHeight) || box.height }, 'para');
    }
  }
  layer.replaceChildren(...marks.map((m) => {
    const span = document.createElement('span');
    span.className = `fmark ${m.className}`;
    span.textContent = m.char;
    span.style.left = `${m.x}px`;
    span.style.top = `${m.y}px`;
    span.style.height = `${m.h}px`;
    span.style.lineHeight = `${m.h}px`;
    span.style.fontSize = `${Math.max(9, Math.round(m.h * 0.7))}px`;
    return span;
  }));
  return marks.length;
}
