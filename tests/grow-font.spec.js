// Grow Font and Shrink Font, as in Word: Ctrl+Shift+> and Ctrl+Shift+< (and the toolbar's buttons)
// step the selected text through Word's sizes (… 10, 11, 12, 14, 16 …); Ctrl+] and Ctrl+[ by 1 pt.
// Text in several sizes changes each. Tracked while tracking.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const DOC = '<w:p><w:r><w:t xml:space="preserve">Body </w:t></w:r><w:r><w:rPr><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:t>big</w:t></w:r></w:p>'
  + '<w:p><w:r><w:t>Other</w:t></w:r></w:p>';

async function selectFirstParagraph(page) {
  await page.evaluate(() => {
    const p = document.querySelector('.docx-page p');
    const range = document.createRange();
    range.selectNodeContents(p);
    document.querySelector('.docx-page').focus();
    getSelection().removeAllRanges();
    getSelection().addRange(range);
  });
}
const firstParagraph = (xml) => xml.match(/<w:body><w:p>(.*?)<\/w:p>/)[1];
const run = (sz, text) => `<w:r><w:rPr><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;

test('Ctrl+Shift+> grows each run to Word\'s next size; Ctrl+Shift+< shrinks it', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await selectFirstParagraph(page);
  await page.keyboard.press('ControlOrMeta+Shift+Period'); // 11 → 12, 14 → 16
  let { xml } = await download(page);
  expect(firstParagraph(xml)).toBe(run(24, 'Body ') + run(32, 'big'));
  await selectFirstParagraph(page);
  await page.keyboard.press('ControlOrMeta+Shift+Comma');
  await page.keyboard.press('ControlOrMeta+Shift+Comma'); // 12 → 11 → 10, 16 → 14 → 12
  ({ xml } = await download(page));
  expect(firstParagraph(xml)).toBe(run(20, 'Body ') + run(24, 'big'));
});

test('Ctrl+] and Ctrl+[ change the size by 1 pt; the toolbar buttons grow and shrink', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await selectFirstParagraph(page);
  await page.keyboard.press('ControlOrMeta+BracketRight');
  await page.keyboard.press('ControlOrMeta+BracketRight');
  await page.keyboard.press('ControlOrMeta+BracketLeft'); // +1 pt in all
  let { xml } = await download(page);
  expect(firstParagraph(xml)).toBe(run(24, 'Body ') + run(30, 'big'));
  await selectFirstParagraph(page);
  await page.getByRole('button', { name: 'Grow font' }).click(); // 12 → 14, 15 → 16
  ({ xml } = await download(page));
  expect(firstParagraph(xml)).toBe(run(28, 'Body ') + run(32, 'big'));
  await selectFirstParagraph(page);
  await page.getByRole('button', { name: 'Shrink font' }).click();
  ({ xml } = await download(page));
  expect(firstParagraph(xml)).toBe(run(24, 'Body ') + run(28, 'big'));
});

test('with nothing selected, it says to select the text', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await page.locator('.docx-page p', { hasText: 'Other' }).click();
  await page.keyboard.press('ControlOrMeta+Shift+Period');
  await expect(page.locator('.statusbar [role=status]')).toHaveText('Select the text first');
});

test('while tracking, growing the font is a tracked formatting change', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await page.getByRole('button', { name: 'Track changes' }).click();
  const dialog = page.getByRole('dialog', { name: 'Your name on tracked changes' });
  await dialog.getByRole('textbox').fill('Rita');
  await dialog.getByRole('button', { name: 'Start tracking' }).click();
  await selectFirstParagraph(page);
  await page.keyboard.press('ControlOrMeta+Shift+Period');
  const { xml } = await download(page);
  expect(firstParagraph(xml)).toMatch(/^<w:r><w:rPr><w:sz w:val="24"\/><w:szCs w:val="24"\/><w:rPrChange w:id="\d+" w:author="Rita" w:date="[^"]+"><w:rPr\/><\/w:rPrChange><\/w:rPr><w:t xml:space="preserve">Body <\/w:t><\/w:r>/);
});
