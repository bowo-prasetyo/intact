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
