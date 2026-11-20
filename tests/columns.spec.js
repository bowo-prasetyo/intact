// Columns, as Word's Layout > Columns: in Page setup, one, two or three equal columns, the space
// between them and a line between them, for the whole document, a section, this point forward or
// the selected text. Columns alone are set off by continuous section breaks (on the same page), as
// Word's Columns dialog makes them; the page shows the columns at once, and Ctrl+Z takes them back.
import fs from 'node:fs';
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, downloadFile } from './helpers.js';

const para = (text) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const BODY = para('Quiz title.') + para('Name and class.') + para('Question one.') + para('Question two.') + para('Question three.');
const dialog = (page) => page.getByRole('dialog', { name: 'Page setup' });
const status = (page) => page.locator('.statusbar [role=status]');
const sections = (xml) => xml.match(/<w:sectPr\b.*?<\/w:sectPr>/gs);
const p = (page, text) => page.locator('.docx-page p', { hasText: text });

async function pageSetup(page) {
  await page.getByRole('button', { name: 'Page setup' }).click();
  await expect(dialog(page)).toBeVisible();
  return dialog(page);
}

test('two columns with a line between, for the whole document: shown at once, written as Word writes them, undone', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await p(page, 'Question one.').click();
  const d = await pageSetup(page);
  await expect(d.getByLabel('Number of columns')).toHaveValue('1');
  await expect(d.getByLabel('Line between')).toHaveCount(0); // only for more than one column
  await d.getByLabel('Number of columns').selectOption('2');
  await expect(d.getByRole('spinbutton', { name: 'Spacing between columns (cm)' })).toHaveValue('1.27');
  await d.getByLabel('Line between').check();
  await d.getByRole('button', { name: 'OK' }).click();
  await expect(status(page)).toHaveText('Page setup changed for the whole document');

  const columns = page.locator('.docx-page > .docx-columns');
  await expect(columns).toHaveCount(1);
  await expect(columns).toHaveAttribute('data-cols', '2');
  await expect(columns).toHaveCSS('column-count', '2');
  await expect(columns).toHaveCSS('column-rule-style', 'solid');
  await expect(columns.locator('p')).toHaveCount(5);
  // Typing goes on in the columns.
  await p(page, 'Question three.').click();
  await page.keyboard.press('End');
  await page.keyboard.type(' Typed');
  await expect(columns).toContainText('Question three. Typed');

  const { xml } = await download(page);
  const [only] = sections(xml);
  expect(only).toContain('<w:cols w:num="2" w:space="720" w:sep="1"/>');
  expect(xml).toContain('Question three. Typed');

  // Ctrl+Z (twice: the typing, then the columns) takes them back.
  await page.locator('.docx-page').focus();
  await page.keyboard.press('Control+z');
  await page.keyboard.press('Control+z');
  await expect(page.locator('.docx-page .docx-columns')).toHaveCount(0);
  const back = await download(page);
  expect(sections(back.xml)[0]).not.toContain('w:cols');
});

test('columns from this point forward: a continuous break, the text before stays in one column on the same page', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await p(page, 'Question one.').click();
  const d = await pageSetup(page);
  await d.getByRole('combobox', { name: 'Apply paper and margins to' }).selectOption('forward');
  await expect(d).toContainText('A section break (next page) goes before the paragraph with the caret');
  await d.getByLabel('Number of columns').selectOption('2');
  await expect(d).toContainText('A continuous section break goes before the paragraph with the caret');
  await d.getByRole('button', { name: 'OK' }).click();
  await expect(status(page)).toHaveText('Page setup changed from this point forward (1 section break added)');

  // The questions in two columns; the title and name above them, not in columns; no new page.
  const columns = page.locator('.docx-page > .docx-columns');
  await expect(columns).toHaveCount(1);
  await expect(columns.locator('p')).toHaveText(['Question one.', 'Question two.', 'Question three.']);
  await expect(p(page, 'Name and class.')).toHaveAttribute('data-section-end', 'continuous');
  await expect(p(page, 'Name and class.')).toHaveAttribute('data-section-note', 'Section 1: A4 portrait · section break (continuous)');
  await expect(p(page, 'Name and class.')).toHaveCSS('break-after', 'auto');
  await expect(page.locator('.page-estimate')).toHaveCount(0); // still one page

  const path = await downloadFile(page);
  const { xml } = await download(page);
  const [first, last] = sections(xml);
  // The first section keeps how it started (it says nothing) and its one column; the new one starts
  // on the same page, in two columns.
  expect(first).not.toContain('<w:type');
  expect(first).not.toContain('w:num=');
  expect(last).toContain('<w:type w:val="continuous"/>');
  expect(last).toContain('<w:cols w:num="2" w:space="720"/>');
  expect(xml).toMatch(/<w:p><w:pPr><w:sectPr>(?:(?!<\/w:p>).)*<\/w:sectPr><\/w:pPr><w:r><w:t[^>]*>Name and class\.<\/w:t>/s);

  // Opened again, it is shown the same: the break is continuous, the questions in columns.
  await page.goto('./');
  await page.setInputFiles('input[type=file]', { name: 'Quiz.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: fs.readFileSync(path) });
  await expect(page.locator('.docx-page > .docx-columns p')).toHaveText(['Question one.', 'Question two.', 'Question three.']);
  await expect(p(page, 'Name and class.')).toHaveAttribute('data-section-end', 'continuous');

  // The dialog shows the caret's section's columns; One takes them away there.
  await p(page, 'Question two.').click();
  const again = await pageSetup(page);
  await expect(again.getByRole('combobox', { name: 'Apply paper and margins to' })).toHaveValue('section');
  await expect(again.getByLabel('Number of columns')).toHaveValue('2');
  await again.getByLabel('Number of columns').selectOption('1');
  await again.getByRole('button', { name: 'OK' }).click();
  await expect(page.locator('.docx-page .docx-columns')).toHaveCount(0);
  const one = await download(page);
  expect(sections(one.xml)[1]).toContain('<w:cols w:space="720"/>');
});

test('the selected paragraphs in three columns: continuous breaks before and after them', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await page.evaluate(() => {
    const ps = [...document.querySelectorAll('.docx-page > p')];
    const start = ps.find((x) => x.textContent === 'Question one.').firstChild;
    const end = ps.find((x) => x.textContent === 'Question two.').firstChild;
    document.querySelector('.docx-page').focus();
    getSelection().setBaseAndExtent(start, 0, end, end.length);
  });
  const d = await pageSetup(page);
  await d.getByRole('combobox', { name: 'Apply paper and margins to' }).selectOption('selected');
  await d.getByLabel('Number of columns').selectOption('3');
  await expect(d).toContainText('Continuous section breaks go before and after the selected paragraphs');
  await d.getByRole('button', { name: 'OK' }).click();
  await expect(status(page)).toHaveText('Page setup changed for the selected text (2 section breaks added)');
  await expect(page.locator('.docx-page > .docx-columns p')).toHaveText(['Question one.', 'Question two.']);
  await expect(page.locator('.docx-page > .docx-columns')).toHaveAttribute('data-cols', '3');
  await expect(p(page, 'Question two.')).toHaveAttribute('data-section-note', 'Section 2: A4 portrait · 3 columns · section break (continuous)');
  const { xml } = await download(page);
  const [, middle, last] = sections(xml);
  expect(middle).toContain('<w:type w:val="continuous"/>');
  expect(middle).toContain('<w:cols w:num="3" w:space="720"/>');
  expect(last).toContain('<w:type w:val="continuous"/>');
  expect(last).not.toContain('w:num=');
});

test('columns too narrow for the page are refused, and nothing changes', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await p(page, 'Question one.').click();
  const d = await pageSetup(page);
  await d.getByLabel('Number of columns').selectOption('3');
  await d.getByRole('spinbutton', { name: 'Spacing between columns (cm)' }).fill('7');
  await d.getByRole('button', { name: 'OK' }).click();
  await expect(status(page)).toHaveText('3 columns with that spacing leave less than 1.27 cm for each: nothing was changed');
  await expect(page.locator('.docx-page .docx-columns')).toHaveCount(0);
});

test('a document in columns of its own widths: the dialog says they become equal when changed', async ({ page }) => {
  const own = '<w:p><w:r><w:t>Narrow and wide.</w:t></w:r></w:p>';
  const buf = await buildDocx(own);
  // The final section in two columns of their own widths.
  const JSZip = (await import('jszip')).default;
  const zip = await JSZip.loadAsync(buf);
  const doc = (await zip.file('word/document.xml').async('string')).replace(/<\/w:sectPr>/, '<w:cols w:num="2" w:space="720" w:equalWidth="0"><w:col w:w="3000" w:space="720"/><w:col w:w="5306"/></w:cols></w:sectPr>');
  zip.file('word/document.xml', doc);
  await openInEditor(page, await zip.generateAsync({ type: 'nodebuffer' }));
  await expect(page.locator('.docx-page > .docx-columns')).toHaveAttribute('data-cols', '2');
  await p(page, 'Narrow and wide.').click();
  const d = await pageSetup(page);
  await expect(d.getByLabel('Number of columns')).toHaveValue('2');
  await expect(d).toContainText('These columns have widths of their own: changed here, they become equal.');
  // Cancelled or unchanged, they stay as they are.
  await d.getByRole('button', { name: 'OK' }).click();
  let { xml } = await download(page);
  expect(xml).toContain('<w:col w:w="3000" w:space="720"/>');
  // Changed (a line between), they become equal.
  const again = await pageSetup(page);
  await again.getByLabel('Line between').check();
  await again.getByRole('button', { name: 'OK' }).click();
  ({ xml } = await download(page));
  expect(xml).toContain('<w:cols w:num="2" w:space="720" w:sep="1"/>');
  expect(xml).not.toContain('<w:col ');
});
