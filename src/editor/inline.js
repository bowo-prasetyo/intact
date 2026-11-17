// Character formatting that Intact applies itself, the same in every browser (the browsers' own
// commands differ, and know nothing of Word's rules): superscript and subscript (<sup>, <sub>, which
// exclude each other) and highlight (<mark data-color>, one of Word's named highlight colours).
// DocxDocument writes them as <w:vertAlign> and <w:highlight>.

/** Word's highlight colours, in the order of its palette: [name in the file, label]. */
export const HIGHLIGHT_COLORS = [
  ['yellow', 'Yellow'], ['green', 'Bright green'], ['cyan', 'Turquoise'], ['magenta', 'Pink'],
  ['blue', 'Blue'], ['red', 'Red'], ['darkBlue', 'Dark blue'], ['darkCyan', 'Teal'],
  ['darkGreen', 'Green'], ['darkMagenta', 'Violet'], ['darkRed', 'Dark red'], ['darkYellow', 'Dark yellow'],
  ['darkGray', 'Grey 50%'], ['lightGray', 'Grey 25%'], ['black', 'Black'],
];

const hostOf = (node) => (node.nodeType === 1 ? node : node.parentElement)?.closest('[contenteditable]');

/** The editable text of a range, as [{ node, s, e }] (text in deletions and read-only parts left out). */
export function portions(range) {
  const root = range.commonAncestorContainer;
  const out = [];
  const add = (node) => {
    if (!node.data || hostOf(node)?.getAttribute('contenteditable') !== 'true' || !range.intersectsNode(node)) return;
    const s = node === range.startContainer ? range.startOffset : 0;
    const e = node === range.endContainer ? range.endOffset : node.length;
    if (e > s) out.push({ node, s, e });
  };
  if (root.nodeType === 3) add(root);
  else {
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    for (let n = walker.nextNode(); n; n = walker.nextNode()) add(n);
  }
  return out;
}

/** Splits a text node so that exactly [s, e) is one node, and returns it. */
export function isolateText({ node, s, e }) {
  let part = node;
  if (s > 0) part = part.splitText(s);
  if (part.length > e - s) part.splitText(e - s);
  return part;
}

/** Splits `ancestor` around `node`, so that one copy of it holds only `node`; returns that copy. */
export function isolate(node, ancestor) {
  const before = document.createRange();
  before.setStart(ancestor, 0);
  before.setEndBefore(node);
  if (!before.collapsed) {
    const copy = ancestor.cloneNode(false);
    copy.append(before.extractContents());
    ancestor.before(copy);
  }
  const after = document.createRange();
  after.setStartAfter(node);
  after.setEnd(ancestor, ancestor.childNodes.length);
  if (!after.collapsed) {
    const copy = ancestor.cloneNode(false);
    copy.append(after.extractContents());
    ancestor.after(copy);
  }
  return ancestor;
}

/** Takes `node` out of every `selector` element around it inside its editing host. */
function unwrapAround(node, selector) {
  const host = hostOf(node);
  for (let el = node.parentElement?.closest(selector); el && host?.contains(el) && el !== host; el = node.parentElement?.closest(selector)) {
    isolate(node, el).replaceWith(...el.childNodes);
  }
}

/** Joins neighbouring elements that are the same (tag and colour), as Word joins runs. */
function joinSiblings(el) {
  const same = (o) => o?.nodeType === 1 && o.tagName === el.tagName && o.dataset.color === el.dataset.color;
  if (same(el.previousSibling)) { el.previousSibling.append(...el.childNodes); const p = el.previousSibling; el.remove(); el = p; }
  if (same(el.nextSibling)) { el.append(...el.nextSibling.childNodes); el.nextSibling.remove(); }
  el.normalize();
}

/** Applies `tag` (with `color` for marks) to the text of a range, or removes it if `tag` is null. */
function format(range, selector, make) {
  const nodes = portions(range).map(isolateText);
  for (const node of nodes) {
    unwrapAround(node, selector);
    const el = make?.();
    if (el) {
      node.replaceWith(el);
      el.append(node);
      joinSiblings(el);
    }
  }
  if (!nodes.length) return null;
  const r = document.createRange();
  r.setStart(nodes[0], 0);
  const last = nodes[nodes.length - 1];
  r.setEnd(last, last.length);
  return r;
}

/** Makes the range superscript or subscript (`kind`), or neither (null). Returns the new range. */
export function setVertAlign(range, kind) {
  return format(range, 'sup, sub', kind ? () => document.createElement(kind === 'superscript' ? 'sup' : 'sub') : null);
}

/** Highlights the range in `color` (a name from HIGHLIGHT_COLORS), or removes highlight (null). */
export function setHighlight(range, color) {
  return format(range, 'mark', color ? () => {
    const mark = document.createElement('mark');
    mark.dataset.color = color;
    return mark;
  } : null);
}

/**
 * The formatting of a range's text: { superscript, subscript } are true if all of it is, and
 * `highlight` is the colour all of it has (or null). For a collapsed range, that of the caret.
 */
export function formatState(range) {
  const parts = portions(range);
  const nodes = parts.length ? parts.map((p) => p.node) : [range.startContainer];
  const host = hostOf(range.startContainer);
  const inside = (n, sel) => {
    const el = (n.nodeType === 1 ? n : n.parentElement)?.closest(sel);
    return el && host?.contains(el) && el !== host ? el : null;
  };
  const colors = new Set(nodes.map((n) => inside(n, 'mark')?.dataset.color || null));
  return {
    superscript: nodes.every((n) => inside(n, 'sup')),
    subscript: nodes.every((n) => inside(n, 'sub')),
    highlight: colors.size === 1 ? [...colors][0] : null,
  };
}
