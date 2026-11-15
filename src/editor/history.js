// Snapshot-based undo/redo for the editor page.
//
// The browser's own contenteditable undo stack only records changes made through execCommand.
// Paragraph style changes (and list renumbering) rewrite the DOM directly, so the native stack
// silently skips them and then gets out of step with the document. This history owns undo instead:
// it stores the page's HTML plus the selection after each change, and restores both on undo/redo.
// Every data-raw/-ppr/-rpr index in the HTML points into the document's append-only raw store, so
// an old snapshot stays valid.

export function createHistory(root, { limit = 100, delay = 400 } = {}) {
  const stack = [];
  let index = -1;
  let timer = 0;
  let composing = false;

  function commit() {
    clearTimeout(timer);
    timer = 0;
    if (composing) return; // never touch the DOM in the middle of IME input
    root.normalize(); // merge split text nodes so selection paths survive an innerHTML round-trip
    const snap = { html: root.innerHTML, sel: saveSelection(root) };
    if (stack[index]?.html === snap.html) {
      stack[index].sel = snap.sel;
      return;
    }
    stack.splice(index + 1);
    stack.push(snap);
    if (stack.length > limit) stack.shift();
    index = stack.length - 1;
  }

  /** Records typing after a short pause, so a burst of keystrokes is one undo step. */
  function schedule() {
    clearTimeout(timer);
    timer = setTimeout(commit, delay);
  }

  /** Records any pending typing now (call before a command so it becomes a separate step). */
  function flush() {
    if (timer) commit();
  }

  function restore(snap) {
    root.innerHTML = snap.html;
    restoreSelection(root, snap.sel);
  }

  function undo() {
    flush();
    if (index <= 0) return false;
    restore(stack[--index]);
    return true;
  }

  function redo() {
    flush();
    if (index >= stack.length - 1) return false;
    restore(stack[++index]);
    return true;
  }

  const onCompositionStart = () => { composing = true; };
  const onCompositionEnd = () => { composing = false; schedule(); };
  root.addEventListener('compositionstart', onCompositionStart);
  root.addEventListener('compositionend', onCompositionEnd);

  function dispose() {
    clearTimeout(timer);
    root.removeEventListener('compositionstart', onCompositionStart);
    root.removeEventListener('compositionend', onCompositionEnd);
  }

  return { commit, schedule, flush, undo, redo, dispose };
}

function pathTo(root, node) {
  const path = [];
  while (node && node !== root) {
    const parent = node.parentNode;
    if (!parent) return null;
    path.unshift(Array.prototype.indexOf.call(parent.childNodes, node));
    node = parent;
  }
  return node === root ? path : null;
}

function nodeAt(root, path) {
  let node = root;
  for (const i of path) {
    if (!node.childNodes[i]) break;
    node = node.childNodes[i];
  }
  return node;
}

function saveSelection(root) {
  const sel = window.getSelection();
  if (!sel.rangeCount || !root.contains(sel.anchorNode) || !root.contains(sel.focusNode)) return null;
  const anchor = pathTo(root, sel.anchorNode);
  const focus = pathTo(root, sel.focusNode);
  return anchor && focus ? { anchor, anchorOffset: sel.anchorOffset, focus, focusOffset: sel.focusOffset } : null;
}

function restoreSelection(root, saved) {
  const sel = window.getSelection();
  if (!saved) {
    root.focus();
    return;
  }
  const clamp = (node, offset) => Math.min(offset, node.nodeType === 3 ? node.data.length : node.childNodes.length);
  const a = nodeAt(root, saved.anchor);
  const f = nodeAt(root, saved.focus);
  root.focus();
  sel.setBaseAndExtent(a, clamp(a, saved.anchorOffset), f, clamp(f, saved.focusOffset));
}
