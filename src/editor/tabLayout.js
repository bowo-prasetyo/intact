// Tabs where Word puts them. The page keeps each tab as a tab character, and a browser puts tab stops
// at multiples of one width (CSS tab-size), measured from where the paragraph's text starts. A tab
// uses the tab-size of the element that holds it, so where a tab is alone in its element (a run of
// its own, as Word writes a table of contents' and Insert answer line writes one), or alone in its
// paragraph, that element is given a tab-size that ends the tab exactly at Word's stop: the next
// stop after the text before it, lined up as the stop says (left; right, centre or at the decimal
// point of the text after it). Its leader (dots, hyphens, a line), or its underline, is drawn across
// it as the element's background, which prints too. Only the page's styles change, never the text:
// saving reads neither (see DocxDocument.inlineCtx).

const LAID = 'tabLaid'; // data-tab-laid: an element given a tab-size or a leader here
const STYLES = ['tab-size', 'background-image', 'background-size', 'background-repeat', 'background-position', 'print-color-adjust', '-webkit-print-color-adjust'];

// Leaders and an underline, as backgrounds of the tab's element: [image, size, position].
const line = (h) => [`linear-gradient(currentColor, currentColor)`, `100% ${h}`, 'left 0 bottom 0.18em'];
const LEADERS = {
  dot: ['radial-gradient(circle, currentColor 0.055em, transparent 0.075em)', '0.33em 0.33em', 'left 0 bottom 0.06em'],
  middleDot: ['radial-gradient(circle, currentColor 0.055em, transparent 0.075em)', '0.33em 0.33em', 'left 0 center'],
  hyphen: ['linear-gradient(to right, currentColor 55%, transparent 55%)', '0.5em 1px', 'left 0 center'],
  underscore: line('1px'),
  heavy: line('2px'),
  underline: line('1px'),
};

/** Every tab in `block`, in order: { node, offset }. */
function tabsIn(block) {
  const out = [];
  const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    for (let i = n.data.indexOf('\t'); i >= 0; i = n.data.indexOf('\t', i + 1)) out.push({ node: n, offset: i });
  }
  return out;
}

const rectOf = (node, start, end) => {
  const r = document.createRange();
  r.setStart(node, start);
  r.setEnd(node, end);
  return r.getBoundingClientRect();
};

/** The width (client px) of the text after a tab on its line, up to the next tab, or to `stopAt` (a decimal point). */
function widthAfter(block, tab, next, stopAt = null) {
  const r = document.createRange();
  r.setStart(tab.node, tab.offset + 1);
  if (next) r.setEnd(next.node, next.offset);
  else r.setEnd(block, block.childNodes.length);
  const text = r.toString();
  if (stopAt) {
    const at = text.search(stopAt);
    if (at >= 0) {
      // Up to the decimal point: a range over that many characters from the tab.
      let left = at;
      const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
      walker.currentNode = tab.node;
      let node = tab.node;
      let start = tab.offset + 1;
      while (node && left > node.length - start) {
        left -= node.length - start;
        node = walker.nextNode();
        start = 0;
      }
      if (node) r.setEnd(node, start + left);
    }
  }
  // The text's own boxes on the tab's line (a range's rects also hold its elements' boxes).
  const top = rectOf(tab.node, tab.offset, tab.offset + 1).top;
  let width = 0;
  const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    if (!r.intersectsNode(n)) continue;
    const s = n === r.startContainer ? r.startOffset : 0;
    const e = n === r.endContainer ? r.endOffset : n.length;
    if (e <= s) continue;
    const t = document.createRange();
    t.setStart(n, s);
    t.setEnd(n, e);
    for (const q of t.getClientRects()) if (Math.abs(q.top - top) < 2 && q.width) width += q.width;
  }
  return width;
}

/** Takes off what layoutTabs set in `block`. */
function reset(block) {
  for (const el of [block, ...block.querySelectorAll('[data-tab-laid]')]) {
    if (el.dataset[LAID] == null) continue;
    for (const p of STYLES) if (!(el === block && p === 'tab-size')) el.style.removeProperty(p);
    delete el.dataset[LAID];
  }
}

/**
 * Lays out the tabs of paragraph `block`, as Word would: `info` is { stops (DocxDocument.tabStops),
 * left (the paragraph's left indent, twips), defaultTab (twips) }. Returns how many tabs were put
 * where Word puts them.
 */
export function layoutTabs(block, { stops, left, defaultTab }) {
  reset(block);
  const tabs = tabsIn(block);
  if (!tabs.length || !block.getClientRects().length) return 0;
  const box = block.getBoundingClientRect();
  const scale = block.offsetWidth ? box.width / block.offsetWidth : 1;
  const cs = getComputedStyle(block);
  const x0 = box.left + (parseFloat(cs.paddingLeft) + parseFloat(cs.borderLeftWidth)) * scale;
  const decimal = /[.,]/;
  let laid = 0;
  tabs.forEach((tab, i) => {
    const parent = tab.node.parentElement;
    // The element whose tab-size this tab takes: its own (alone in it), or the paragraph's (alone in it).
    const holder = parent !== block && parent.textContent === '\t' ? parent : tabs.length === 1 ? block : null;
    if (!holder) return;
    const start = (rectOf(tab.node, tab.offset, tab.offset + 1).left - x0) / scale; // css px from the text's start
    const at = left + start * 15; // twips from the margin
    const stop = stops.find((s) => s.align !== 'bar' && s.pos > at + 7);
    let end;
    let leader = 'none';
    if (stop) {
      const px = (stop.pos - left) / 15;
      const after = stop.align === 'left' ? 0 : widthAfter(block, tab, tabs[i + 1], stop.align === 'decimal' ? decimal : null) / scale;
      end = stop.align === 'right' || stop.align === 'decimal' ? px - after : stop.align === 'center' ? px - after / 2 : px;
      leader = stop.leader || 'none';
    } else {
      // Word's default stops, every `defaultTab` from the margin, after the paragraph's own.
      end = ((Math.floor(at / defaultTab) + 1) * defaultTab - left) / 15;
    }
    end = Math.max(end, start + 1); // text already past the stop: the tab goes on just after it
    holder.style.tabSize = `${end}px`;
    holder.dataset[LAID] = '';
    laid++;
    const underlined = holder !== block && getComputedStyle(holder).textDecorationLine.includes('underline');
    const look = LEADERS[leader !== 'none' ? leader : underlined ? 'underline' : ''];
    if (look && holder !== block) {
      const [image, size, position] = look;
      Object.assign(holder.style, { backgroundImage: image, backgroundSize: size, backgroundRepeat: leader === 'dot' || leader === 'middleDot' || leader === 'hyphen' ? 'repeat-x' : 'no-repeat', backgroundPosition: position });
      holder.style.setProperty('print-color-adjust', 'exact');
      holder.style.setProperty('-webkit-print-color-adjust', 'exact');
    }
  });
  return laid;
}
