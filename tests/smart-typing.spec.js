// Quotes, dashes and ellipses become typographic as you type, as Word's AutoFormat: curly quotes in
// the text's language, an em dash for "--" between words, an en dash for " - ", an ellipsis for
// "...". It can be turned off (remembered in this browser), and is tracked while tracking. The
// status bar counts characters as well as words.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const para = (page) => page.locator('.docx-page > p').first();
const lang = (code) => `<w:p><w:r><w:rPr><w:lang w:val="${code}"/></w:rPr><w:t>Start</w:t></w:r></w:p>`;

async function typeAtEnd(page, text) {
  await para(page).click();
  await page.keyboard.press('End');
  await page.keyboard.type(text);
}

test('quotes become curly, opening or closing as in Word; apostrophes too', async ({ page }) => {
  await openInEditor(page, await buildDocx(lang('en-GB')));
  await typeAtEnd(page, ': "It\'s here," she said (\'quietly\').');
  await expect(para(page)).toHaveText('Start: “It’s here,” she said (‘quietly’).');
  const { xml } = await download(page);
  expect(xml).toContain('Start: “It’s here,” she said (‘quietly’).');
});

test('dashes and an ellipsis: word--word, word - word and ...', async ({ page }) => {
  await openInEditor(page, await buildDocx(lang('en-US')));
  await typeAtEnd(page, ' well--almost. Pages 4 - 6 and so on...');
  await expect(para(page)).toHaveText('Start well—almost. Pages 4 – 6 and so on…');
});

test('the quotes follow the language of the text: German and French', async ({ page }) => {
  await openInEditor(page, await buildDocx(lang('de-DE')));
  await typeAtEnd(page, ' "Hallo"');
  await expect(para(page)).toHaveText('Start „Hallo“');
  await page.close();
});

test('French quotes come with no-break spaces inside', async ({ page }) => {
  await openInEditor(page, await buildDocx(lang('fr-FR')));
  await typeAtEnd(page, ' "Bonjour"');
  await expect(para(page)).toHaveText('Start « Bonjour »');
});

test('turned off, quotes and dashes are typed as they are; the choice is remembered', async ({ page }) => {
  await openInEditor(page, await buildDocx(lang('en-GB')));
  const toggle = page.getByRole('button', { name: /^Smart quotes and dashes: on$/ });
  await toggle.click();
  await expect(page.getByRole('button', { name: /^Smart quotes and dashes: off$/ })).toHaveAttribute('aria-pressed', 'false');
  await typeAtEnd(page, ' "plain" -- text');
  await expect(para(page)).toHaveText('Start "plain" -- text');
  await page.reload();
  await openInEditor(page, await buildDocx(lang('en-GB')));
  await expect(page.getByRole('button', { name: /^Smart quotes and dashes: off$/ })).toBeVisible();
});

test('while tracking, a curly quote and a dash are tracked insertions', async ({ page }) => {
  await openInEditor(page, await buildDocx(lang('en-GB')));
  await page.getByRole('button', { name: 'Track changes' }).click();
  const dialog = page.getByRole('dialog', { name: 'Your name on tracked changes' });
  await dialog.getByRole('textbox').fill('Rita');
  await dialog.getByRole('button', { name: 'Start tracking' }).click();
  await typeAtEnd(page, ' "a--b"');
  await expect(para(page)).toHaveText('Start “a—b”');
  const { xml } = await download(page);
  // Typed in one go, it is one tracked insertion: the hyphens it replaced were the session's own.
  expect(xml).toMatch(/<w:ins w:id="\d+" w:author="Rita" w:date="[^"]+"><w:r><w:rPr><w:lang w:val="en-GB"\/><\/w:rPr><w:t xml:space="preserve"> “a—b”<\/w:t><\/w:r><\/w:ins>/);
  expect(xml).not.toContain('<w:del ');
});

test('the status bar counts characters, with and without spaces, in the text and the selection', async ({ page }) => {
  await openInEditor(page, await buildDocx('<w:p><w:r><w:t>Two words.</w:t></w:r></w:p><w:p><w:r><w:t>And three more.</w:t></w:r></w:p>'));
  const count = page.locator('.statusbar > span').first();
  await expect(count).toHaveText('5 words · 25 characters');
  await expect(count).toHaveAttribute('title', '25 characters with spaces, 22 without');
  await page.evaluate(() => {
    const t = document.querySelector('.docx-page p').firstChild;
    getSelection().setBaseAndExtent(t, 0, t, 3);
  });
  await expect(count).toHaveText('1 of 5 words · 3 of 25 characters');
});
