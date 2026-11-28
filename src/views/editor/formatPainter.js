// Format Painter: the formatting of the text or paragraph at the caret, painted on the next selection (once, or until Esc when double-clicked), as Word's.
// Split out of EditorView.vue: see src/views/editor/README.md. The editor's shared state comes in as
// `ctx`, read when it is used, never copied at setup.
import { reactive } from 'vue';
import { elementChildren } from '../../docx/xml.js';
import { isAtBlockStart } from '../../editor/structure.js';
import { portions } from '../../editor/inline.js';
import { restyle } from '../../editor/runs.js';

/**
 * `ctx`: { applyBlockStyle, closestBlock, doc, explainLock, formatSelection, message, page, restoreSelection, savedRange, selectRange, selectedBlocks, selectionInPage, state, textLocked, tracking, wordAt }, as the editor has them.
 * Returns { onPageMouseup, painter, startPainter, togglePainter, paintSelection, stopPainter, stopPainterQuietly }.
 */
export function useFormatPainter(ctx) {
  // Word's Format Painter: a click copies the character formatting where the caret is (or of the
  // selection's first character) and paints it on the text selected next; a double-click keeps it on
  // for several places, until Esc or another click on it. Selecting whole paragraphs paints the
  // paragraph's style too. Ctrl+Shift+C and Ctrl+Shift+V copy and paint without the button.
  const painter = reactive({ active: false, sticky: false, copied: false, rPr: null, block: null });

  /** The text node whose formatting the selection (or the caret) has: the first selected, or beside the caret. */
  function formatSourceNode(range) {
    const first = portions(range)[0];
    if (first) return first.node;
    let node = range.startContainer;
    if (node.nodeType === 3) return node;
    node = node.childNodes[range.startOffset - 1] || node.childNodes[range.startOffset] || node;
    const walker = document.createTreeWalker(ctx.closestBlock(node) || node, NodeFilter.SHOW_TEXT);
    return walker.nextNode();
  }

  /** Copies the formatting at the selection; `active`: and turns the painter on (`sticky`: until stopped). */
  function startPainter(sticky = false, keyboard = false) {
    ctx.restoreSelection();
    const range = ctx.selectionInPage() || ctx.savedRange;
    const node = range && formatSourceNode(range);
    if (!node || !ctx.page.value.contains(node)) {
      ctx.message.value = 'Put the caret in the text whose formatting you want to copy';
      return;
    }
    Object.assign(painter, { copied: true, rPr: ctx.doc.rPrOf(node), block: blockKind(ctx.closestBlock(node)) });
    if (keyboard) {
      ctx.message.value = 'Formatting copied – select text and press Ctrl+Shift+V to paint it on';
      return;
    }
    Object.assign(painter, { active: true, sticky });
    ctx.message.value = sticky
      ? 'Format Painter: select (or click) each place to paint the formatting on – Esc when you’re done'
      : 'Format Painter: select (or click) the text to paint the formatting on';
    ctx.page.value.focus();
  }

  function stopPainterQuietly() {
    painter.active = false;
    painter.sticky = false;
  }

  function stopPainter() {
    painter.active = false;
    painter.sticky = false;
    ctx.message.value = 'Format Painter off';
  }

  function togglePainter() {
    if (painter.active) stopPainter();
    else startPainter(false);
  }

  /** The paragraph style select's value for a block, as state.block reads it. */
  const blockKind = (block) => (block ? ctx.doc.blockValue(block.dataset.style) : null);

  /** After selecting with the painter on: paints the formatting on the selection (or the word clicked). */
  function onPageMouseup() {
    if (!painter.active) return;
    setTimeout(() => {
      if (!painter.active) return;
      let range = ctx.selectionInPage();
      if (!range) return;
      if (range.collapsed) {
        range = ctx.wordAt(range);
        if (!range) return;
        ctx.selectRange(range);
      }
      paintSelection();
      if (!painter.sticky) painter.active = false;
    });
  }

  /**
   * Paints the copied formatting on the selection: its character formatting replaces the text's (a
   * tracked formatting change while tracking), and selecting whole paragraphs gives them the copied
   * paragraph style too.
   */
  function paintSelection() {
    if (!painter.copied) return;
    if (ctx.textLocked.value) return ctx.explainLock();
    ctx.restoreSelection();
    const range = ctx.selectionInPage();
    if (!range || range.collapsed) {
      ctx.message.value = 'Select the text to paint the formatting on';
      return;
    }
    const blocks = ctx.selectedBlocks(range);
    const whole = blocks.length && isAtBlockStart(range, blocks[0]) && atBlockEnd(range, blocks[blocks.length - 1]);
    if (whole && painter.block && blocks.some((b) => blockKind(b) !== painter.block)) ctx.applyBlockStyle(painter.block);
    const src = painter.rPr;
    const painted = ctx.formatSelection((r) => restyle(ctx.doc, r, (rPr) => {
      const change = elementChildren(rPr).find((c) => c.localName === 'rPrChange');
      for (const c of elementChildren(rPr)) if (c !== change) rPr.removeChild(c);
      for (const c of src ? elementChildren(src) : []) if (c.localName !== 'rPrChange') rPr.insertBefore(c.cloneNode(true), change || null);
    }));
    if (painted) ctx.message.value = painter.sticky ? 'Formatting painted – select the next place, or Esc when you’re done' : 'Formatting painted';
  }

  /** True if `range` ends at the end of `block`'s text. */
  function atBlockEnd(range, block) {
    const rest = document.createRange();
    rest.setStart(range.endContainer, range.endOffset);
    rest.setEnd(block, block.childNodes.length);
    return !rest.toString().replace(/​/g, '');
  }

  return { onPageMouseup, painter, startPainter, togglePainter, paintSelection, stopPainter, stopPainterQuietly };
}
