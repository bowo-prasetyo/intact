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
  expect(xml).not.toContain(' ');
});

test.describe('with WebKit\'s rewriting', () => {
  test.use({ userAgent: 'Mozilla/5.0 (Macintosh; Intel Mac OS X 14_6) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/18.0 Safari/605.1.15' });

  /** Does to the text what WebKit does, then sends the input event the browser would. */
  const rewrite = (page, find, at, inputType, data = null) => page.evaluate(([f, i, type, d]) => {
    const w = document.createTreeWalker(document.querySelector('.docx-page'), NodeFilter.SHOW_TEXT);
    for (let n = w.nextNode(); n; n = w.nextNode()) {
      if (!n.data.includes(f)) continue;
      n.replaceData(n.data.indexOf(f) + i, 1, ' ');
      document.querySelector('.docx-page').dispatchEvent(new InputEvent('input', { inputType: type, data: d, bubbles: true }));
      return n.data;
    }
    throw new Error(`no text "${f}"`);
  }, [find, at, inputType, data]);

  test('a space WebKit made U+00A0 goes back to a space when the input comes', async ({ page }) => {
    await openInEditor(page, await buildDocx(para('Second paragraph.')));
    expect(await rewrite(page, 'Second ', 6, 'insertParagraph')).toBe('Second paragraph.');
  });

  test('a no-break space already in the document stays one, also when the text around it is edited', async ({ page }) => {
    await openInEditor(page, await buildDocx(para('Page 12 of the report')));
    await caretAt(page, 'report', 6);
    await page.keyboard.type('s');
    const { xml } = await download(page);
    expect(xml).toContain('Page 12 of the reports');
  });
});
