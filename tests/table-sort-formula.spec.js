// Sorting a table and table formulas, as Word's Table Layout > Sort and Formula: the rows below the
// heading row sorted by a column's text, number or date, as one undoable step; and Word's = field
// (=SUM(ABOVE), =B2*C2) in a cell, worked out with the document's own separators, written as Word
// writes it, and worked out again by Update fields (F9).
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';
import JSZip from 'jszip';

const cell = (text) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:tc>`;
const row = (...texts) => `<w:tr>${texts.map(cell).join('')}</w:tr>`;
const table = (...rows) => `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>${rows.join('')}</w:tbl><w:p/>`;
const BUDGET = table(
  row('Kegiatan', 'Tanggal', 'Biaya'),
  row('Rapat koordinasi', '17/08/2026', 'Rp 1.250.000'),
  row('Bimbingan teknis', '02/03/2026', 'Rp 12.500.000'),
  row('Sosialisasi', '25/12/2025', 'Rp 750.000'),
  row('Jumlah', '', ''),
);

/** The document as Indonesian: its default language id-ID, so numbers are written 1.250.000,50. */
async function indonesian(body) {
  const zip = await JSZip.loadAsync(await buildDocx(body));
  const styles = await zip.file('word/styles.xml').async('string');
  zip.file('word/styles.xml', styles.replace('<w:rPr><w:sz w:val="22"/></w:rPr>', '<w:rPr><w:sz w:val="22"/><w:lang w:val="id-ID"/></w:rPr>'));
  return zip.generateAsync({ type: 'nodebuffer' });
}
const firstColumn = (page) => page.locator('.docx-page table tr td:first-child').allTextContents();
const status = (page) => page.locator('.statusbar [role=status]');

async function tableMenu(page, item) {
  await page.getByTitle('Table layout: insert or delete rows and columns').click();
  await page.getByRole('menu', { name: 'Table layout' }).getByRole('menuitem', { name: item }).click();
}

test('Sort: by number, by date and by text, the heading row kept on top; Ctrl+Z puts the rows back', async ({ page }) => {
  await openInEditor(page, await indonesian(BUDGET.replace(row('Jumlah', '', ''), '')));
  await page.locator('.docx-page td', { hasText: 'Rp 750.000' }).click();
  await tableMenu(page, 'Sort…');
  const dialog = page.getByRole('dialog', { name: 'Sort' });
  // The column with the caret, by its heading; numbers sorted as numbers.
  await expect(dialog.getByRole('combobox', { name: 'Sort by' })).toHaveValue('2');
  await expect(dialog.getByRole('option', { name: 'Biaya' })).toHaveCount(1);
  await expect(dialog.getByRole('combobox', { name: 'Type' })).toHaveValue('number');
  await dialog.getByLabel('Descending').check();
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(status(page)).toHaveText('Sorted 3 rows by Biaya, descending');
  expect(await firstColumn(page)).toEqual(['Kegiatan', 'Bimbingan teknis', 'Rapat koordinasi', 'Sosialisasi']);
  // By date, oldest first (day/month/year, as the document's language writes dates).
  await page.locator('.docx-page td', { hasText: 'Sosialisasi' }).click();
  await tableMenu(page, 'Sort…');
  await dialog.getByRole('combobox', { name: 'Sort by' }).selectOption({ label: 'Tanggal' });
  await dialog.getByRole('combobox', { name: 'Type' }).selectOption('date');
  await dialog.getByRole('button', { name: 'OK' }).click();
  expect(await firstColumn(page)).toEqual(['Kegiatan', 'Sosialisasi', 'Bimbingan teknis', 'Rapat koordinasi']);
  // By text, A to Z.
  await page.locator('.docx-page td', { hasText: 'Sosialisasi' }).click();
  await tableMenu(page, 'Sort…');
  await dialog.getByRole('combobox', { name: 'Sort by' }).selectOption({ label: 'Kegiatan' });
  await dialog.getByRole('button', { name: 'OK' }).click();
  expect(await firstColumn(page)).toEqual(['Kegiatan', 'Bimbingan teknis', 'Rapat koordinasi', 'Sosialisasi']);
  const { xml } = await download(page);
  expect(xml.indexOf('Bimbingan teknis')).toBeLessThan(xml.indexOf('Rapat koordinasi'));
  expect(xml.indexOf('Kegiatan')).toBeLessThan(xml.indexOf('Bimbingan teknis'));
  await page.locator('.docx-page td', { hasText: 'Sosialisasi' }).click();
  await page.keyboard.press('ControlOrMeta+z');
  expect(await firstColumn(page)).toEqual(['Kegiatan', 'Sosialisasi', 'Bimbingan teknis', 'Rapat koordinasi']);
});

test('Sort isn\'t tracked: while tracking it says so and changes nothing', async ({ page }) => {
  await openInEditor(page, await indonesian(BUDGET));
  await page.locator('.docx-page td', { hasText: 'Sosialisasi' }).click();
  await page.getByRole('button', { name: 'Track changes' }).click();
  const name = page.getByRole('dialog', { name: 'Your name on tracked changes' });
  await name.getByRole('textbox').fill('Rita');
  await name.getByRole('button', { name: 'Start tracking' }).click();
  await expect(page.locator('.statusbar .tracking-status')).toBeVisible();
  await page.locator('.docx-page td', { hasText: 'Sosialisasi' }).click();
  await tableMenu(page, 'Sort…');
  await page.getByRole('dialog', { name: 'Sort' }).getByRole('button', { name: 'OK' }).click();
  await expect(status(page)).toHaveText(/^Sorting isn’t tracked/);
  expect(await firstColumn(page)).toEqual(['Kegiatan', 'Rapat koordinasi', 'Bimbingan teknis', 'Sosialisasi', 'Jumlah']);
});

test('Formula: =SUM(ABOVE) in Rupiah, written as Word\'s = field, worked out again by Update fields', async ({ page }) => {
  await openInEditor(page, await indonesian(BUDGET));
  const total = page.locator('.docx-page tr').last().locator('td').nth(2);
  await total.click();
  await tableMenu(page, 'Formula…');
  const dialog = page.getByRole('dialog', { name: 'Formula' });
  await expect(dialog.getByRole('textbox', { name: 'Formula' })).toHaveValue('=SUM(ABOVE)');
  await dialog.getByRole('combobox', { name: 'Number format' }).selectOption('Rp #.##0');
  await expect(dialog.getByRole('status')).toHaveText('Result: Rp 14.500.000');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(total).toHaveText('Rp 14.500.000');
  await expect(total.locator('.field')).toHaveAttribute('data-instr', /^\s*=SUM\(ABOVE\) \\# "Rp #\.##0"\s*$/);
  // A cost changed: F9 works the total out again.
  const cost = page.locator('.docx-page tr').nth(3).locator('td').nth(2);
  await cost.click();
  await page.keyboard.press('End');
  await page.keyboard.press('Backspace');
  await page.keyboard.press('Backspace');
  await page.keyboard.press('Backspace');
  await page.keyboard.press('Backspace');
  await page.keyboard.press('Backspace');
  await page.keyboard.press('Backspace');
  await page.keyboard.press('Backspace');
  await page.keyboard.type('250.000');
  await expect(cost).toHaveText('Rp 250.000');
  await page.keyboard.press('F9');
  await expect(status(page)).toContainText('1 formula');
  await expect(total).toHaveText('Rp 14.000.000');
  const { xml } = await download(page);
  expect(xml).toContain('<w:instrText xml:space="preserve"> =SUM(ABOVE) \\# "Rp #.##0" </w:instrText>');
  expect(xml).toMatch(/<w:fldChar w:fldCharType="separate"\/><\/w:r><w:r><w:t[^>]*>Rp 14\.000\.000<\/w:t>/);
});

test('Formula: cell references and arithmetic, a function pasted in, and one Word wouldn\'t understand refused', async ({ page }) => {
  await openInEditor(page, await buildDocx(table(row('Qty', 'Price', 'Amount'), row('3', '2.50', ''))));
  const amount = page.locator('.docx-page tr').last().locator('td').nth(2);
  await amount.click();
  await tableMenu(page, 'Formula…');
  const dialog = page.getByRole('dialog', { name: 'Formula' });
  // Nothing above it is a number; to the left there are: Word offers SUM(LEFT).
  await expect(dialog.getByRole('textbox', { name: 'Formula' })).toHaveValue('=SUM(LEFT)');
  await dialog.getByRole('textbox', { name: 'Formula' }).fill('=A2*B2*(1+11%)');
  await dialog.getByRole('combobox', { name: 'Number format' }).selectOption('#,##0.00');
  await expect(dialog.getByRole('status')).toHaveText('Result: 8.33');
  await dialog.getByRole('textbox', { name: 'Formula' }).fill('=A2*');
  await expect(dialog.getByRole('status')).toHaveText(/^Word wouldn’t understand this formula/);
  await expect(dialog.getByRole('button', { name: 'OK' })).toBeDisabled();
  await dialog.getByRole('textbox', { name: 'Formula' }).fill('=');
  await dialog.getByRole('combobox', { name: 'Paste function' }).selectOption('MAX');
  await expect(dialog.getByRole('textbox', { name: 'Formula' })).toHaveValue('=MAX()');
  await dialog.getByRole('textbox', { name: 'Formula' }).fill('=MAX(LEFT)');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(amount).toHaveText('3.00'); // in the number format chosen
});
