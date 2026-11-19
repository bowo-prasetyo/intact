// Lengths in centimetres or inches, as Word shows them in the unit of the place it is used: inches
// for a document on North American paper (Letter, Legal, Executive), centimetres for the others, or
// the unit chosen in the Page setup or Paragraph dialog, remembered in this browser. What is written
// is the same either way (twips), and a length left as it was is never rewritten.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const BODY = '<w:p><w:r><w:t>Smith, Jane. A Book. Press, 2020.</w:t></w:r></w:p>';
const setup = (page) => page.getByRole('dialog', { name: 'Page setup' });
const para = (page) => page.getByRole('dialog', { name: 'Paragraph' });

async function openSetup(page) {
  await page.locator('.docx-page p').first().click();
  await page.getByRole('button', { name: 'Page setup' }).click();
  await expect(setup(page)).toBeVisible();
}
async function openParagraph(page) {
  await page.locator('.docx-page p').first().click();
  await page.getByRole('button', { name: 'Line and paragraph spacing' }).click();
  await page.getByRole('menuitem', { name: 'Paragraph…' }).click();
  await expect(para(page)).toBeVisible();
}

test('A4 shows centimetres; on Letter paper the margins and indents are in inches', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await openSetup(page);
  await expect(setup(page).getByLabel('Units')).toHaveValue('cm');
  await expect(setup(page).getByLabel('Top (cm)')).toHaveValue('2.54');
  await setup(page).getByLabel('Size').selectOption('Letter');
  await setup(page).getByRole('button', { name: 'OK' }).click();
  await openSetup(page);
  await expect(setup(page).getByLabel('Units')).toHaveValue('in');
  await expect(setup(page).getByLabel('Top (in)')).toHaveValue('1');
  await expect(setup(page).getByLabel('Preset')).toHaveValue('Normal');
  await setup(page).getByLabel('Preset').selectOption('Narrow');
  await expect(setup(page).getByLabel('Left (in)')).toHaveValue('0.5');
  await setup(page).getByRole('button', { name: 'Cancel' }).click();
  // A hanging indent of half an inch, as a works-cited list has.
  await openParagraph(page);
  await expect(para(page).getByLabel('Left (in)')).toHaveValue('0');
  await para(page).getByLabel('Special').selectOption('hanging');
  await para(page).getByLabel('By (in)').fill('0.5');
  await para(page).getByRole('button', { name: 'OK' }).click();
  const { xml } = await download(page);
  expect(xml).toContain('<w:pPr><w:ind w:hanging="720"/></w:pPr><w:r><w:t xml:space="preserve">Smith, Jane.');
  expect(xml).toContain('<w:pgSz w:w="12240" w:h="15840"/>');
});

test('choosing inches converts what the dialog shows, writes only what changed, and is remembered', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await openSetup(page);
  await setup(page).getByLabel('Left (cm)').fill('3');
  await setup(page).getByLabel('Units').selectOption('in');
  await expect(setup(page).getByLabel('Top (in)')).toHaveValue('1'); // left as it was: 1 in exactly
  await expect(setup(page).getByLabel('Left (in)')).toHaveValue('1.18'); // changed: 3 cm, converted
  await setup(page).getByRole('button', { name: 'OK' }).click();
  let { xml } = await download(page);
  // Only the left margin changed (1.18 in = 1699 twips); the others stay as the file had them.
  expect(xml).toContain('<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1699" w:header="708" w:footer="708" w:gutter="0"/>');
  // Remembered: the Paragraph dialog and the tab stops are in inches too, on this A4 document.
  await openParagraph(page);
  await expect(para(page).getByLabel('Units')).toHaveValue('in');
  await expect(para(page).getByText('every 0.5 in')).toBeVisible();
  await para(page).getByLabel('Tab stop position (in)').fill('1');
  await para(page).getByRole('button', { name: 'Set' }).click();
  await expect(para(page).getByRole('listitem')).toContainText('1 in');
  // Back to centimetres while the dialog is open: the tab stop shows in centimetres.
  await para(page).getByLabel('Units').selectOption('cm');
  await expect(para(page).getByRole('listitem')).toContainText('2.54 cm');
  await para(page).getByRole('button', { name: 'OK' }).click();
  ({ xml } = await download(page));
  expect(xml).toContain('<w:tabs><w:tab w:val="left" w:pos="1440"/></w:tabs>');
  expect(xml).not.toContain('<w:ind');
  await page.reload();
  expect(await page.evaluate(() => localStorage.getItem('intact.unit'))).toBe('cm');
});
