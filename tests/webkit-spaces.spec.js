// Safari's engine (WebKit) writes an ordinary space as a no-break space (U+00A0) where editing
// leaves one at a paragraph's edge: before the caret when Enter splits a paragraph, or typed at the
// end of one. In the file that would be another character than the one typed, so the editor puts the
// space back (see restoreRewrittenSpaces in EditorView.vue). A no-break space typed on purpose, or
// already in the document, is never touched. In WebKit (the webkit project) the first test is the
// real thing; elsewhere, the rest run the same code with Safari's user agent, rewriting a space as
// WebKit does.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const para = (text) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

/** Puts the caret `offset` characters into the text node holding `text`. */
const caretAt = (page, text, offset) => page.evaluate(([t, o]) => {
  const w = document.createTreeWalker(document.querySelector('.docx-page'), NodeFilter.SHOW_TEXT);
  for (let n = w.nextNode(); n; n = w.nextNode()) {
    if (!n.data.includes(t)) continue;
    document.querySelector('.docx-page').focus();
    getSelection().collapse(n, n.data.indexOf(t) + o);
    return;
  }
  throw new Error(`no text "${t}"`);
}, [text, offset]);

test('Enter after a space, and a space typed at the end, are saved as ordinary spaces', async ({ page }) => {
  await openInEditor(page, await buildDocx(para('Second paragraph.') + para('Last')));
  await caretAt(page, 'Second paragraph.', 7);
  await page.keyboard.press('Enter');
  await caretAt(page, 'Last', 4);
  await page.keyboard.type(' ');
  const { xml } = await download(page);
  expect(xml).toContain('>Second </w:t>');
  expect(xml).toContain('>paragraph.</w:t>');
  expect(xml).toContain('>Last </w:t>');
  expect(xml).not.toContain('\u00a0');
});

test.describe('with WebKit\'s rewriting', () => {
  test.use({ userAgent: 'Mozilla/5.0 (Macintosh; Intel Mac OS X 14_6) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/18.0 Safari/605.1.15' });

  /** Does to the text what WebKit does, then sends the input event the browser would. */
  const rewrite = (page, find, at, inputType, data = null) => page.evaluate(([f, i, type, d]) => {
    const w = document.createTreeWalker(document.querySelector('.docx-page'), NodeFilter.SHOW_TEXT);
    for (let n = w.nextNode(); n; n = w.nextNode()) {
      if (!n.data.includes(f)) continue;
      n.replaceData(n.data.indexOf(f) + i, 1, '\u00a0');
      document.querySelector('.docx-page').dispatchEvent(new InputEvent('input', { inputType: type, data: d, bubbles: true }));
      return n.data;
    }
    throw new Error(`no text "${f}"`);
  }, [find, at, inputType, data]);

  test('a space WebKit made U+00A0 goes back to a space when the input comes', async ({ page }) => {
    await openInEditor(page, await buildDocx(para('Second paragraph.')));
    expect(await rewrite(page, 'Second ', 6, 'insertParagraph')).toBe('Second paragraph.');
  });

  /** Replaces the text node holding `find` with new nodes (`parts`), as WebKit splits text, then sends the input. */
  const replaceNode = (page, find, parts, inputType, data = null) => page.evaluate(([f, ps, type, d]) => {
    const w = document.createTreeWalker(document.querySelector('.docx-page'), NodeFilter.SHOW_TEXT);
    for (let n = w.nextNode(); n; n = w.nextNode()) {
      if (!n.data.includes(f)) continue;
      const made = ps.map((x) => document.createTextNode(x));
      n.replaceWith(...made);
      document.querySelector('.docx-page').dispatchEvent(new InputEvent('input', { inputType: type, data: d, bubbles: true }));
      return made.map((m) => m.data);
    }
    throw new Error(`no text "${f}"`);
  }, [find, parts, inputType, data]);

  test('a split into new text nodes: the space WebKit made U+00A0 in the new node goes back', async ({ page }) => {
    await openInEditor(page, await buildDocx(para('Second paragraph.')));
    expect(await replaceNode(page, 'Second paragraph.', ['Second\u00a0', 'paragraph.'], 'insertParagraph')).toEqual(['Second ', 'paragraph.']);
  });

  test('a space typed into a node of its own goes back to a space; a no-break space typed so stays', async ({ page }) => {
    await openInEditor(page, await buildDocx(para('Last') + para('Page')));
    expect(await replaceNode(page, 'Last', ['Last', '\u00a0'], 'insertText', ' ')).toEqual(['Last', ' ']);
    expect(await replaceNode(page, 'Page', ['Page', '\u00a0'], 'insertText', '\u00a0')).toEqual(['Page', '\u00a0']);
  });

  test('a no-break space typed beside a space WebKit also rewrote: the typed one stays, the other goes back', async ({ page }) => {
    await openInEditor(page, await buildDocx(para('Alpha paragraph')));
    await caretAt(page, 'Alpha paragraph', 5);
    // As WebKit does when U+00A0 is typed before a space: both are U+00A0, the caret after the first.
    const data = await page.evaluate(() => {
      const n = getSelection().anchorNode;
      n.replaceData(5, 1, '\u00a0\u00a0');
      getSelection().collapse(n, 6);
      document.querySelector('.docx-page').dispatchEvent(new InputEvent('input', { inputType: 'insertText', data: '\u00a0', bubbles: true }));
      return n.data;
    });
    expect(data).toBe('Alpha\u00a0 paragraph');
  });

  test('a no-break space already in the document stays one, also when the text around it is edited', async ({ page }) => {
    await openInEditor(page, await buildDocx(para('Page\u00a012 of the report')));
    await caretAt(page, 'report', 6);
    await page.keyboard.type('s');
    const { xml } = await download(page);
    expect(xml).toContain('Page\u00a012 of the reports');
  });
});
