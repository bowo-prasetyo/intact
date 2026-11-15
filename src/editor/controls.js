// DOM helpers for content controls in the editor (see src/docx/contentControls.js).

/** The editable content control (`.sdt` with data-sdt) containing `node`, within `root`. */
export function controlAt(node, root) {
  const el = (node?.nodeType === 1 ? node : node?.parentElement)?.closest('.sdt[data-sdt]');
  return el && root.contains(el) ? el : null;
}

/** The element holding a control's content. */
export function controlBox(el) {
  return el.querySelector(':scope > .sdt-content') || el;
}

/** Replaces a control's text, keeping the formatting of its first run (as Word does). */
export function setControlText(el, text) {
  const box = controlBox(el);
  const walker = document.createTreeWalker(box, NodeFilter.SHOW_TEXT);
  const nodes = [];
  while (walker.nextNode()) nodes.push(walker.currentNode);
  if (nodes.length) {
    nodes[0].data = text;
    for (const n of nodes.slice(1)) n.data = '';
  } else {
    (box.querySelector('p, h1, h2, h3, h4, h5, h6') || box).append(text);
  }
  for (const br of box.querySelectorAll('br')) if (!br.nextSibling) br.remove();
}

/** Once the user has typed over a placeholder, stop showing it as one. */
export function clearPlaceholder(el, isPlaceholderStyle) {
  if (!el.classList.contains('showing-placeholder')) return;
  el.classList.remove('showing-placeholder');
  for (const s of controlBox(el).querySelectorAll('[data-rstyle]')) {
    if (isPlaceholderStyle(s.dataset.rstyle)) delete s.dataset.rstyle;
  }
}

/** The node right before (dir -1) or after (dir 1) a DOM position, looking out of inline wrappers. */
function neighbour(container, offset, dir, root) {
  let node = container;
  let off = offset;
  while (node && node !== root) {
    if (node.nodeType === 3) {
      if (dir > 0 ? off < node.data.length : off > 0) return null; // inside text
    } else {
      const child = node.childNodes[dir > 0 ? off : off - 1];
      if (child) return child;
    }
    if (node.nodeType === 1 && /^(P|H[1-6]|TD|DIV|LI)$/.test(node.tagName)) return null; // don't leave the paragraph
    const parent = node.parentNode;
    off = Array.prototype.indexOf.call(parent.childNodes, node) + (dir > 0 ? 1 : 0);
    node = parent;
  }
  return null;
}

const CHOICE = '.sdt[data-sdt]:is(.sdt-checkbox, .sdt-dropdown, .sdt-date, .sdt-combo)';

/**
 * The check box, drop-down, date or combo box field the selection is on: selected (alone), holding
 * the caret (a combo box's text), or touching the caret on either side.
 */
export function fieldAtSelection(range, root) {
  if (!range) return null;
  const inside = controlAt(range.startContainer, root);
  if (inside?.matches(CHOICE)) return inside;
  if (!range.collapsed) {
    // A selection of fields only (no other text): the first one. Browsers differ in how far
    // Shift+Arrow extends over a non-editable field, so don't insist on an exact match.
    const hits = [...root.querySelectorAll(CHOICE)].filter((el) => range.intersectsNode(el));
    if (!hits.length) return null;
    const rest = range.cloneContents();
    for (const el of rest.querySelectorAll('.sdt')) el.remove();
    return rest.textContent.trim() ? null : hits[0];
  }
  for (const dir of [1, -1]) {
    const n = neighbour(range.startContainer, range.startOffset, dir, root);
    if (n?.nodeType === 1 && n.matches(CHOICE)) return n;
  }
  return null;
}
