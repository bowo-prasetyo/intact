// The caret never stays at the edge of read-only text where typing does nothing: Firefox puts it
// inside someone's tracked change that ends (or starts) a paragraph, with End or Home or a click
// beside it; it is moved just outside, so typing goes in next to the change, untouched.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor } from './helpers.js';

const REV = 'w:author="Ann" w:date="2026-01-01T00:00:00Z"';
const BODY = `<w:p><w:ins w:id="1" ${REV}><w:r><w:t xml:space="preserve">Inserted by Ann, </w:t></w:r></w:ins><w:r><w:t xml:space="preserve">kept </w:t></w:r><w:ins w:id="2" ${REV}><w:r><w:t>and added by Ann</w:t></w:r></w:ins></w:p>`;
const para = (page) => page.locator('.docx-page p').first();

/** Puts the caret inside the n-th tracked insertion, `at` its start or end, as Firefox does. */
const caretInside = (page, n, at) => page.evaluate(([i, end]) => {
  const ins = document.querySelectorAll('.docx-page .raw-inline.inserted')[i];
  const walker = document.createTreeWalker(ins, NodeFilter.SHOW_TEXT);
  const texts = [];
  for (let t = walker.nextNode(); t; t = walker.nextNode()) texts.push(t);
  const t = end ? texts[texts.length - 1] : texts[0];
  document.querySelector('.docx-page').focus();
  getSelection().setBaseAndExtent(t, end ? t.length : 0, t, end ? t.length : 0);
}, [n, at === 'end']);

test('a caret at the end of a tracked insertion ending the paragraph is moved after it: typing goes in', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await caretInside(page, 1, 'end');
  await page.keyboard.type('!');
  await expect(para(page)).toHaveText('Inserted by Ann, kept and added by Ann!');
  await expect(page.locator('.docx-page .raw-inline.inserted').nth(1)).toHaveText('and added by Ann'); // Ann's change untouched
});

test('a caret at the start of a tracked insertion starting the paragraph is moved before it', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await caretInside(page, 0, 'start');
  await page.keyboard.type('Note: ');
  await expect(para(page)).toHaveText('Note: Inserted by Ann, kept and added by Ann');
  await expect(page.locator('.docx-page .raw-inline.inserted').first()).toHaveText('Inserted by Ann, ');
});

test('End then typing, in a paragraph ending with someone’s tracked insertion, adds the text after it', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await para(page).getByText('kept').click();
  await page.keyboard.press('End');
  await page.keyboard.type('!');
  await expect(para(page)).toHaveText('Inserted by Ann, kept and added by Ann!');
});

test('in dark mode, the status bar’s “Tracking changes as …” is readable (7.5:1)', async ({ page }) => {
  await page.emulateMedia({ colorScheme: 'dark' });
  await openInEditor(page, await buildDocx(BODY));
  await page.getByRole('button', { name: 'Track changes' }).click();
  const dialog = page.getByRole('dialog', { name: 'Your name on tracked changes' });
  await dialog.getByRole('textbox').fill('Rita');
  await dialog.getByRole('button', { name: 'Start tracking' }).click();
  await expect(page.locator('.statusbar .tracking-status')).toHaveCSS('color', 'rgb(196, 163, 255)');
});
