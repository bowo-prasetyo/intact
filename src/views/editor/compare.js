// Compare documents (Review › Compare): the other version read and compared, its changes listed and gone to, and a redline made from them, as Word's Compare makes one.
// Split out of EditorView.vue: see src/views/editor/README.md. The editor's shared state comes in as
// `ctx`, read when it is used, never copied at setup.
import { reactive, ref } from 'vue';
import JSZip from 'jszip';
import { t } from '../../i18n.js';
import { paragraphsOfXml, compareParagraphs, notesOfPackage, compareNotes } from '../../editor/compare.js';
import { redlinePackage } from '../../docx/redline.js';
import { openFileInNewTab } from '../../tabs.js';
import { sniffFile } from '../../docx/package.js';

/**
 * `ctx`: { BLOCK_SELECTOR, doc, page, openPasswordDialog, draftSettled, withHostsUnlocked, message, refocus, placeCaret, author, baseName, fileExt, passwordProtected, count, downloadBlob }, as the editor has them.
 * Returns { compare, compareInput, startCompare, onCompareFile, compareWith, closeCompare, goToChange, sendRedline }.
 */
export function useCompare(ctx) {
  const compare = reactive({ result: null, other: '', blocks: [], otherXml: '', redline: null });
  const compareInput = ref(null);

  function startCompare() {
    compareInput.value.value = '';
    compareInput.value.click();
  }

  /** The page's paragraphs (not headers, footers or notes), with text, as the XML's are read. */
  function pageParagraphs() {
    const blocks = [...ctx.page.value.querySelectorAll(ctx.BLOCK_SELECTOR)].filter((el) => !el.closest('.docx-hf-band, .docx-notes, .textbox') && !el.parentElement.closest(ctx.BLOCK_SELECTOR));
    const texts = blocks.map((el) => {
      const copy = el.cloneNode(true);
      for (const x of copy.querySelectorAll('[data-tc="del"], .raw-hidden, .note-ref')) x.remove();
      for (const x of copy.querySelectorAll('.raw-inline[data-raw]')) {
        if (['del', 'moveFrom'].includes(ctx.doc.raw[Number(x.dataset.raw)]?.localName)) x.remove();
      }
      return copy.textContent;
    });
    const kept = texts.map((tx, i) => [tx, blocks[i]]).filter(([tx]) => tx.replace(/[\s\u00a0]+/g, ''));
    return { texts: kept.map(([tx]) => tx), blocks: kept.map(([, b]) => b) };
  }

  /**
   * The page's paragraph for each of the XML's (`texts`), to go to a change: the same one where the
   * page has as many, else the next whose text is the same (without spaces), else the nearest.
   */
  function blocksFor(texts) {
    const { texts: shown, blocks } = pageParagraphs();
    if (shown.length === texts.length) return blocks;
    const squash = (x) => x.replace(/[\s\u00a0\u2029⧉]+/g, '');
    const out = [];
    let j = 0;
    for (const text of texts) {
      const k = shown.findIndex((x, i) => i >= j && i < j + 8 && squash(x) === squash(text));
      if (k >= 0) j = k;
      out.push(blocks[Math.min(j, blocks.length - 1)]);
      if (k >= 0) j += 1;
    }
    return out;
  }

  async function onCompareFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    const bytes = new Uint8Array(await file.arrayBuffer());
    // Encrypted with a password: asked for (see the password dialog), then compared.
    if (sniffFile(bytes)?.kind === 'password') {
      ctx.openPasswordDialog('compare', { name: file.name, bytes });
      return;
    }
    await compareWith(file.name, bytes);
  }

  /**
   * Compares the other version (`bytes`, a .docx package, named `name`) with this document. Both are
   * read from their XML the same way (this one as Download would write it), so a document compared
   * with itself is the same.
   */
  async function compareWith(name, bytes) {
    const file = { name };
    try {
      // A .doc, or a file encrypted by rights management: say which, as opening one does.
      const unreadable = sniffFile(bytes);
      if (unreadable) throw new Error(t('{name} can’t be compared: {why}', { name: file.name, why: unreadable.message }));
      const zip = await JSZip.loadAsync(bytes);
      const main = Object.keys(zip.files).find((f) => /^word\/document\d*\.xml$/i.test(f));
      if (!main) throw new Error(t('{name} isn’t a Word document (.docx)', { name: file.name }));
      const otherXml = await zip.file(main).async('string');
      const before = paragraphsOfXml(otherXml);
      await ctx.draftSettled();
      await ctx.withHostsUnlocked(() => ctx.doc.write(ctx.page.value));
      const thisXml = await ctx.doc.zip.file(ctx.doc.mainPath).async('string');
      const after = paragraphsOfXml(thisXml);
      // Footnotes and endnotes too, each kind in the order the text refers to them, numbered as shown.
      const thisNotes = await notesOfPackage(ctx.doc.zip, ctx.doc.mainPath, thisXml);
      const notes = compareNotes(await notesOfPackage(zip, main, otherXml), thisNotes).map((c) => ({
        ...c, n: c.at < thisNotes[c.note].length ? ctx.doc.noteNumber(c.note, c.at + 1) : null,
      }));
      compare.blocks = blocksFor(after);
      compare.other = file.name;
      compare.otherXml = otherXml;
      compare.result = { ...compareParagraphs(before, after), notes };
      // The redline is made now, while the result is read, so that opening it in a new tab follows
      // the click at once (a browser opens a tab only straight after a click).
      compare.redline = compare.result.changes.length ? makeRedline().catch((e) => e) : null;
    } catch (err) {
      ctx.message.value = /zip|central directory|End of data/i.test(err.message)
        ? t('{name} can’t be compared: it isn’t a .docx (an older .doc can’t be read here)', { name: file.name })
        : err.message;
    }
  }

  function closeCompare() {
    compare.result = null;
    compare.redline = null;
    ctx.refocus();
  }

  /** Goes to a change in this document: its paragraph (or, for text taken out, where it was). */
  function goToChange(change) {
    if (change.note) {
      goToNoteChange(change);
      return;
    }
    const block = compare.blocks[Math.min(change.at, compare.blocks.length - 1)];
    compare.result = null;
    compare.redline = null;
    if (!block?.isConnected) return;
    ctx.page.value.focus();
    ctx.placeCaret(block, change.at < compare.blocks.length);
    block.scrollIntoView({ block: 'center' });
    ctx.message.value = change.kind === 'removed' ? t('Text taken out was here') : change.kind === 'added' ? t('Added in this document') : t('Changed in this document');
  }

  /** Goes to a change in a footnote or endnote: the note (or, for one taken out, the note after it, or the last). */
  function goToNoteChange(change) {
    compare.result = null;
    compare.redline = null;
    const notes = [...ctx.page.value.querySelectorAll(`.docx-note[data-note="${change.note}"]:not([hidden])`)];
    const el = (change.n && notes.find((x) => x.dataset.n === change.n)) || notes[notes.length - 1];
    if (!el) return;
    const block = el.querySelector(ctx.BLOCK_SELECTOR) || el;
    el.focus();
    ctx.placeCaret(block, change.kind !== 'removed' || !!change.n);
    el.scrollIntoView({ block: 'center' });
    const name = change.note === 'footnote' ? t('footnote') : t('endnote');
    ctx.message.value = change.kind === 'removed' ? t('A {note} taken out was here', { note: name })
      : change.kind === 'added' ? t('Added in this document: {note} {n}', { note: name, n: change.n }) : t('Changed in this document: {note} {n}', { note: name, n: change.n });
  }

  /**
   * The redline, as Word's Compare makes one: a new document, this one as Download would write it,
   * with the changes since the other version as tracked changes in your name (see docx/redline.js).
   * Encrypted with this document's password, if it has one. { blob, name, accepted }.
   */
  async function makeRedline() {
    await ctx.draftSettled();
    const saved = await ctx.withHostsUnlocked(() => ctx.doc.save(ctx.page.value, { plain: true }));
    const { zip, accepted } = await redlinePackage(JSZip, await saved.arrayBuffer(), ctx.doc.mainPath, compare.otherXml, { author: ctx.author.value, date: new Date() });
    const mime = ctx.doc.format.mime;
    const plain = await zip.generateAsync({ type: 'blob', mimeType: mime, compression: 'DEFLATE' });
    const blob = await ctx.doc.sealed(plain, true);
    const base = ctx.baseName.value.trim() || 'document';
    return { blob: new Blob([blob], { type: mime }), name: `${base} (redline).${ctx.fileExt.value}`, accepted };
  }

  /** The redline: opened in a new tab beside this document (`how` 'open'), or downloaded. */
  async function sendRedline(how) {
    const made = await compare.redline;
    if (!made || made instanceof Error) {
      ctx.message.value = `The redline couldn’t be made: ${made?.message || 'nothing changed'}`;
      return;
    }
    const accepted = made.accepted ? ` – the ${ctx.count(made.accepted, 'tracked change')} already in this document ${made.accepted === 1 ? 'is' : 'are'} accepted in it, as Word’s Compare does` : '';
    const locked = ctx.passwordProtected.value ? ', encrypted with this document’s password' : '';
    if (how === 'open') {
      const result = openFileInNewTab(new File([made.blob], made.name, { type: made.blob.type }));
      ctx.message.value = {
        opened: `The redline is opening in a new tab${locked}${accepted}`,
        blocked: 'The browser didn’t open a new tab: allow pop-ups for this site, or download the redline',
        unsupported: `A new tab is open: download the redline and choose it there${locked}`,
      }[result];
      if (result === 'opened') closeCompare();
      return;
    }
    ctx.downloadBlob(made.blob, made.name);
    ctx.message.value = `Redline downloaded: ${made.name}${locked}${accepted}`;
    closeCompare();
  }

  return { compare, compareInput, startCompare, onCompareFile, compareWith, closeCompare, goToChange, sendRedline };
}
