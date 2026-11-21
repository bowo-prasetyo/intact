// Comments in the printout and PDF, for marked work returned as a PDF: asked for in Save as PDF
// (remembered, and followed by Print too), the comments are listed after the document, each with
// its number as the text shows it, author, date, the words it is on and its replies. The list is
// beside the page, never in it: the document saved is the same. Not asked for, nothing of the
// comments prints, their numbers in the text neither.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, W_NS } from './helpers.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const anchored = (id, before, on, after) => `<w:p><w:r><w:t xml:space="preserve">${before}</w:t></w:r><w:commentRangeStart w:id="${id}"/><w:r><w:t xml:space="preserve">${on}</w:t></w:r><w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r><w:r><w:t xml:space="preserve">${after}</w:t></w:r></w:p>`;
const BODY = anchored(1, 'The results ', 'clearly', ' show it.') + anchored(2, 'Encapsulation is ', 'very important', '.');
const comment = (id, author, text, para) => `<w:comment w:id="${id}" w:author="${author}" w:date="2026-11-03T10:00:00Z" w:initials="X"><w:p w14:paraId="${para}"><w:r><w:t>${text}</w:t></w:r></w:p></w:comment>`;
const COMMENTS = `${XML}<w:comments ${W_NS}>${comment(1, 'Prof. Reviewer', 'Avoid intensifiers.', '0000000A')}${comment(2, 'Prof. Reviewer', 'Quantify this.', '0000000B')}${comment(3, 'A. Student', 'Fixed in the next draft.', '0000000C')}</w:comments>`;
const EXTENDED = `${XML}<w15:commentsEx xmlns:w15="${W15}"><w15:commentEx w15:paraId="0000000A" w15:done="0"/><w15:commentEx w15:paraId="0000000B" w15:done="1"/><w15:commentEx w15:paraId="0000000C" w15:paraIdParent="0000000A" w15:done="0"/></w15:commentsEx>`;
const OPTS = {
  parts: { 'word/comments.xml': COMMENTS, 'word/commentsExtended.xml': EXTENDED },
  rels: `<Relationship Id="rIdC" Type="${REL}/comments" Target="comments.xml"/><Relationship Id="rIdX" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>`,
};

async function openPdfDialog(page) {
  await page.evaluate(() => { window.printed = 0; window.print = () => { window.printed++; }; });
  await page.getByRole('button', { name: 'More ways to save' }).click();
  await page.getByRole('menu', { name: 'Save' }).getByRole('menuitem', { name: /PDF/ }).click();
  return page.getByRole('dialog', { name: 'Save as PDF' });
}
const beforePrint = (page) => page.evaluate(() => window.dispatchEvent(new Event('beforeprint')));
const afterPrint = (page) => page.evaluate(() => window.dispatchEvent(new Event('afterprint')));
const list = (page) => page.locator('.docx-sheet > .print-comments');

test('asked for in Save as PDF: the comments print after the document, numbered as in the text, with replies', async ({ page, browserName }) => {
  await openInEditor(page, await buildDocx(BODY, OPTS));
  const dialog = await openPdfDialog(page);
  const option = dialog.getByRole('checkbox', { name: /Print comments/ });
  await expect(option).not.toBeChecked();
  await expect(dialog).toContainText('Comments and the marks drawn over the page here aren’t in the PDF.');
  await option.check();
  await expect(dialog).toContainText('The marks drawn over the page here aren’t in the PDF.');
  await dialog.getByRole('button', { name: 'Open print window' }).click();
  await expect.poll(() => page.evaluate(() => window.printed)).toBe(1);

  await beforePrint(page);
  await expect(list(page)).toHaveCount(1);
  await page.emulateMedia({ media: 'print' });
  await expect(list(page)).toBeVisible();
  await expect(list(page).locator('h2')).toHaveText('Comments');
  const items = list(page).locator('li');
  await expect(items).toHaveCount(2);
  await expect(items.nth(0).locator('.n')).toHaveText('1');
  await expect(items.nth(0)).toContainText('Prof. Reviewer');
  await expect(items.nth(0).locator('.on')).toHaveText('On “clearly”');
  await expect(items.nth(0).locator('.text').first()).toHaveText('Avoid intensifiers.');
  await expect(items.nth(0).locator('.reply')).toContainText('Reply from A. Student');
  await expect(items.nth(0).locator('.reply')).toContainText('Fixed in the next draft.');
  await expect(items.nth(1).locator('.on')).toHaveText('On “very important” · resolved');
  // The numbers in the text print too, as the list refers to them.
  const marker = await page.locator('.docx-page .comment-ref').first().evaluate((el) => getComputedStyle(el, '::before').content);
  expect(marker).not.toBe('none');
  // Starts on a page of its own, after the document.
  expect(await list(page).evaluate((el) => getComputedStyle(el).breakBefore)).toBe('page');
  // A PDF is made with it (Playwright makes PDFs in Chromium only).
  if (browserName === 'chromium') expect(Buffer.byteLength(await page.pdf({ preferCSSPageSize: true }))).toBeGreaterThan(0);

  // Gone after printing; the document saved is the same as without it.
  await page.emulateMedia({ media: 'screen' });
  await afterPrint(page);
  await expect(list(page)).toHaveCount(0);
  const { xml } = await download(page);
  expect(xml).not.toContain('Avoid intensifiers');
  expect(xml).not.toContain('print-comments');
  // Remembered for the next print, Ctrl+P too.
  expect(await page.evaluate(() => localStorage.getItem('intact.printComments'))).toBe('1');
});

test('not asked for, nothing of the comments prints, their numbers in the text neither', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY, OPTS));
  await beforePrint(page);
  await expect(list(page)).toHaveCount(0);
  await page.emulateMedia({ media: 'print' });
  const marker = await page.locator('.docx-page .comment-ref').first().evaluate((el) => getComputedStyle(el, '::before').content);
  expect(marker).toBe('none');
  await page.emulateMedia({ media: 'screen' });
  await afterPrint(page);
});

test('a document without comments offers no such option', async ({ page }) => {
  await openInEditor(page, await buildDocx('<w:p><w:r><w:t>No comments here.</w:t></w:r></w:p>'));
  const dialog = await openPdfDialog(page);
  await expect(dialog.getByRole('checkbox', { name: /Print comments/ })).toHaveCount(0);
  await expect(dialog).toContainText('Comments and the marks drawn over the page here aren’t in the PDF.');
});
