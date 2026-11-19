// Mail merge from a CSV, as Word's Mailings tab does it with a list: each row fills the document's
// placeholders, giving one document of every row's copy (each on a new page) or a .zip of a
// document for each row. The list stays on the device and the document itself isn't changed.
import fs from 'node:fs';
import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import { buildDocx, openInEditor } from './helpers.js';

const LETTER = '<w:p><w:r><w:t>[Recipient name]</w:t></w:r></w:p>'
  + '<w:p><w:r><w:t>[Address]</w:t></w:r></w:p>'
  // A placeholder across two runs, the first bold: the value takes the first run's formatting.
  + '<w:p><w:r><w:t xml:space="preserve">Dear </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>[Recip</w:t></w:r><w:r><w:t>ient name],</w:t></w:r></w:p>'
  + '<w:p><w:bookmarkStart w:id="3" w:name="Body"/><w:ins w:id="7" w:author="Ana" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Your order [Order] is ready.</w:t></w:r></w:ins><w:bookmarkEnd w:id="3"/></w:p>';
const CSV = 'Recipient name,address,ORDER\r\nBudi Santoso,"Jl. Merdeka 1\r\nJakarta",A-17\r\n"Ana, S.E.",Bandung,B-2\r\n';

async function openMerge(page) {
  await page.locator('.app-header').getByRole('button', { name: 'More ways to save' }).click();
  await page.getByRole('menuitem', { name: /^Mail merge…/ }).click();
  return page.getByRole('dialog', { name: 'Mail merge' });
}

async function chooseList(dialog, csv, name = 'list.csv') {
  await dialog.locator('input[type=file]').setInputFiles({ name, mimeType: 'text/csv', buffer: Buffer.isBuffer(csv) ? csv : Buffer.from(csv) });
}

async function merge(page, dialog) {
  const [dl] = await Promise.all([page.waitForEvent('download'), dialog.getByRole('button', { name: /^Merge \d+ rows?$/ }).click()]);
  return { name: dl.suggestedFilename(), bytes: fs.readFileSync(await dl.path()) };
}

test('a list fills the letter once for each row, in one document, each copy on a new page', async ({ page }) => {
  await openInEditor(page, await buildDocx(LETTER), 'letter.docx');
  const dialog = await openMerge(page);
  await chooseList(dialog, CSV);
  await expect(dialog.getByRole('status')).toHaveText('list.csv: 2 rows, 3 columns');
  // Matched by name, whatever the case.
  await expect(dialog.getByRole('combobox', { name: 'Column for [Recipient name]' })).toHaveValue('0');
  await expect(dialog.getByRole('combobox', { name: 'Column for [Address]' })).toHaveValue('1');
  await expect(dialog.getByRole('combobox', { name: 'Column for [Order]' })).toHaveValue('2');
  await expect(dialog.locator('.map .value').first()).toHaveText('Budi Santoso');
  await dialog.getByRole('button', { name: 'Next row' }).click();
  await expect(dialog.locator('.map .value').first()).toHaveText('Ana, S.E.');

  const out = await merge(page, dialog);
  expect(out.name).toBe('letter (merged).docx');
  await expect(dialog).toHaveCount(0);
  await expect(page.locator('.statusbar [role=status]')).toHaveText('Merged 2 rows: letter (merged).docx');
  // The document itself is as it was.
  await expect(page.locator('.docx-page p').first()).toHaveText('[Recipient name]');

  const xml = await (await JSZip.loadAsync(out.bytes)).file('word/document.xml').async('string');
  expect(xml).not.toContain('[Recipient name]');
  expect(xml.match(/Budi Santoso/g)).toHaveLength(2);
  expect(xml.match(/Ana, S\.E\./g)).toHaveLength(2);
  // The second copy starts on a new page; the first doesn't.
  expect(xml.match(/<w:pageBreakBefore\/>/g)).toHaveLength(1);
  // An address's line break is Word's.
  expect(xml).toContain('<w:t xml:space="preserve">Jl. Merdeka 1</w:t><w:br/><w:t xml:space="preserve">Jakarta</w:t>');
  // Across runs: the name in the first run's formatting (bold), the rest of the second run kept.
  expect(xml).toMatch(/<w:rPr><w:b\/><\/w:rPr><w:t xml:space="preserve">Budi Santoso<\/w:t><\/w:r><w:r><w:t>,<\/w:t>/);
  // Ids that must be unique in a document are, and bookmarks' names too.
  // Each copy's ids are above every id of the document's (bookmark 3, insertion 7): 3, 7 and 11, 15.
  expect([...xml.matchAll(/<w:(?:ins|bookmarkStart) w:id="(\d+)"/g)].map((m) => Number(m[1])).sort((a, b) => a - b)).toEqual([3, 7, 11, 15]);
  expect(xml).toContain('w:name="Body"');
  expect(xml).toContain('w:name="Body_2"');
  expect(xml).toContain('Your order A-17 is ready.');
  expect(xml).toContain('Your order B-2 is ready.');
});

test('a separate document for each row comes in a .zip, named by a column', async ({ page }) => {
  await openInEditor(page, await buildDocx(LETTER), 'letter.docx');
  const dialog = await openMerge(page);
  await chooseList(dialog, CSV);
  await dialog.getByRole('radio', { name: /separate document for each row/ }).check();
  await expect(dialog.getByRole('combobox', { name: 'Name each file by' })).toHaveValue('0'); // "Recipient name"
  const out = await merge(page, dialog);
  expect(out.name).toBe('letter (merged).zip');
  const zip = await JSZip.loadAsync(out.bytes);
  expect(Object.keys(zip.files).sort()).toEqual(['Ana, S.E..docx', 'Budi Santoso.docx']);
  const one = await JSZip.loadAsync(await zip.file('Budi Santoso.docx').async('uint8array'));
  const xml = await one.file('word/document.xml').async('string');
  expect(xml).toContain('Your order A-17 is ready.');
  expect(xml).not.toContain('S.E.');
  expect(xml).not.toContain('pageBreakBefore');
});

test('a placeholder left without a column stays as it is; an Excel list with semicolons in Windows-1252 is read', async ({ page }) => {
  await openInEditor(page, await buildDocx(LETTER), 'letter.docx');
  const dialog = await openMerge(page);
  // "Recipient name;Order" then "José Niño;C-9", in Windows-1252 (é = E9, ñ = F1).
  const bytes = Buffer.concat([Buffer.from('Recipient name;Order\r\nJos'), Buffer.from([0xe9]), Buffer.from(' Ni'), Buffer.from([0xf1]), Buffer.from('o;C-9\r\n')]);
  await chooseList(dialog, bytes, 'excel.csv');
  await expect(dialog.getByRole('combobox', { name: 'Column for [Address]' })).toHaveValue('-1');
  await expect(dialog.getByText('1 placeholder is left as it is')).toBeVisible();
  const out = await merge(page, dialog);
  const xml = await (await JSZip.loadAsync(out.bytes)).file('word/document.xml').async('string');
  expect(xml).toContain('José Niño');
  expect(xml).toContain('[Address]');
});

test('Fill in offers the list too; a document without placeholders says how to add them', async ({ page }) => {
  await openInEditor(page, await buildDocx(LETTER));
  await page.getByRole('button', { name: 'Fill in placeholders' }).click();
  await page.getByRole('dialog', { name: 'Fill in placeholders' }).getByRole('button', { name: 'From a list…' }).click();
  await expect(page.getByRole('dialog', { name: 'Mail merge' })).toBeVisible();
  await page.keyboard.press('Escape');

  await openInEditor(page, await buildDocx('<w:p><w:r><w:t>No brackets here.</w:t></w:r></w:p>'));
  const dialog = await openMerge(page);
  await expect(dialog).toContainText('This document has no placeholders');
});

test('a document with footnotes is merged into separate documents only', async ({ page }) => {
  const notes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    + '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p><w:r><w:t>A note.</w:t></w:r></w:p></w:footnote></w:footnotes>';
  await openInEditor(page, await buildDocx('<w:p><w:r><w:t xml:space="preserve">[Recipient name]</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>', {
    parts: { 'word/footnotes.xml': notes },
    rels: '<Relationship Id="rIdFn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>',
  }));
  const dialog = await openMerge(page);
  await chooseList(dialog, CSV);
  await expect(dialog.getByRole('radio', { name: /One document/ })).toBeDisabled();
  await expect(dialog.getByRole('radio', { name: /separate document/ })).toBeChecked();
  await expect(dialog.getByText('Not for this document: it has footnotes or endnotes')).toBeVisible();
});

test('a starter letter merges with a list whose columns are its placeholders', async ({ page }) => {
  await page.goto('./');
  await page.getByRole('button', { name: /^Letter:/ }).click();
  await page.getByRole('dialog', { name: /fill in/i }).getByRole('button', { name: 'From a list…' }).click();
  const dialog = page.getByRole('dialog', { name: 'Mail merge' });
  await chooseList(dialog, 'Recipient name,Company name,Company address\nMs Rina,PT Sinar,Surabaya\nMr Joko,CV Terang,Medan\n');
  await expect(dialog.getByRole('combobox', { name: 'Column for [Company name]' })).toHaveValue('1');
  const out = await merge(page, dialog);
  const xml = await (await JSZip.loadAsync(out.bytes)).file('word/document.xml').async('string');
  expect(xml).toContain('Dear Ms Rina,');
  expect(xml).toContain('Dear Mr Joko,');
  expect(xml).toContain('[Your name]'); // not in the list: kept to fill in
});
