// Changing a table's layout from the Table layout menu, as on Word's Layout tab: rows inserted
// above or below (copies of the row, inside vertical merges where they fall in one), columns
// inserted left or right (the grid and widths follow), rows, columns and tables deleted. Rows are
// tracked while tracking; column changes can't be, and the status bar says so.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, canonicalize, documentXmlOf } from './helpers.js';

const CELL = (text, props = '<w:tcW w:w="3000" w:type="dxa"/>') => `<w:tc><w:tcPr>${props}</w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const GRID = '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>';
const TABLE = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>${GRID}`
  + `<w:tr><w:trPr><w:tblHeader/></w:trPr>${CELL('Item')}${CELL('Price', '<w:tcW w:w="6000" w:type="dxa"/><w:gridSpan w:val="2"/>')}</w:tr>`
  + `<w:tr>${CELL('Paper')}${CELL('5', '<w:tcW w:w="3000" w:type="dxa"/><w:vMerge w:val="restart"/>')}${CELL('EUR')}</w:tr>`
  + `<w:tr>${CELL('Pens')}<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:vMerge/></w:tcPr><w:p/></w:tc>${CELL('USD')}</w:tr>`
  + '</w:tbl><w:p><w:r><w:t>Below the table.</w:t></w:r></w:p>';

const status = (page) => page.locator('.statusbar [role=status]');
const rows = (xml) => [...xml.matchAll(/<w:tr>.*?<\/w:tr>|<w:tr [^>]*>.*?<\/w:tr>/g)].map((m) => m[0]);
const cellTexts = (page) => page.locator('table tr').evaluateAll((trs) => trs.map((tr) => [...tr.cells].map((c) => c.textContent)));

async function layout(page, item) {
  await page.getByTitle('Table layout: insert or delete rows and columns').click();
  await page.getByRole('menu', { name: 'Table layout' }).getByRole('menuitem', { name: item }).click();
}

async function startTracking(page) {
  await page.getByRole('button', { name: 'Track changes' }).click();
  await page.getByRole('dialog', { name: 'Your name on tracked changes' }).getByRole('textbox').fill('Rita');
  await page.getByRole('button', { name: 'Start tracking' }).click();
}

test('Table layout is only for tables; a row inserted below is a copy, with the caret in it', async ({ page }) => {
  await openInEditor(page, await buildDocx(TABLE));
  await page.getByText('Below the table.').click();
  await expect(page.getByTitle('Table layout: insert or delete rows and columns')).toBeDisabled();
  await page.locator('td', { hasText: 'Item' }).click();
  await layout(page, 'Insert row below');
  await expect(status(page)).toHaveText('Row inserted below');
  await page.keyboard.type('New');
  await expect.poll(() => cellTexts(page)).toEqual([['Item', 'Price'], ['New', ''], ['Paper', '5', 'EUR'], ['Pens', '', 'USD']]);
  const { xml } = await download(page);
  // Below a header row: not a header itself (only rows within the header rows repeat), same cells.
  expect(rows(xml)[1]).toBe('<w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">New</w:t></w:r></w:p></w:tc>'
    + '<w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/><w:gridSpan w:val="2"/></w:tcPr><w:p/></w:tc></w:tr>');
});

test('above a header row the new row repeats too; inside a vertical merge it is merged', async ({ page }) => {
  await openInEditor(page, await buildDocx(TABLE));
  await page.locator('td', { hasText: 'Item' }).click();
  await layout(page, 'Insert row above');
  await page.locator('td', { hasText: 'Paper' }).click();
  await layout(page, 'Insert row below');
  const { xml } = await download(page);
  expect(rows(xml)[0]).toMatch(/^<w:tr><w:trPr><w:tblHeader\/><\/w:trPr>/);
  // Between "5" (merge start) and the merged cell below: the new cell carries the merge on.
  expect(rows(xml)[3]).toBe('<w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p/></w:tc>'
    + '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:vMerge/></w:tcPr><w:p/></w:tc>'
    + '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p/></w:tc></w:tr>');
});

test('with two rows selected, two rows are inserted; Undo takes them back exactly', async ({ page }) => {
  const docx = await buildDocx(TABLE);
  await openInEditor(page, docx);
  await page.evaluate(() => {
    const range = document.createRange();
    range.setStart(document.querySelector('tr:nth-child(2) td').firstChild, 0);
    range.setEnd(document.querySelector('tr:nth-child(3) td').firstChild, 1);
    document.querySelector('.docx-page').focus();
    getSelection().removeAllRanges();
    getSelection().addRange(range);
  });
  await layout(page, 'Insert 2 rows above');
  await expect(status(page)).toHaveText('2 rows inserted above');
  await expect(page.locator('table tr')).toHaveCount(5);
  await page.keyboard.press('ControlOrMeta+z');
  await expect(page.locator('table tr')).toHaveCount(3);
  const { xml } = await download(page);
  expect(await canonicalize(page, xml)).toEqual(await canonicalize(page, await documentXmlOf(docx)));
});

test('a column inserted to the right: a cell in every row, spans widened, the table fitted to the page', async ({ page }) => {
  await openInEditor(page, await buildDocx(TABLE));
  await page.locator('td', { hasText: 'Paper' }).click();
  await layout(page, 'Insert column to the right');
  await expect(status(page)).toHaveText('Column inserted to the right');
  await page.keyboard.type('Qty');
  await expect.poll(() => cellTexts(page)).toEqual([['Item', '', 'Price'], ['Paper', 'Qty', '5', 'EUR'], ['Pens', '', '', 'USD']]);
  const { xml } = await download(page);
  // Four columns of 3000 would overflow the 9026 twips between the margins: all narrow alike.
  expect(xml).toContain('<w:tblGrid><w:gridCol w:w="2256"/><w:gridCol w:w="2256"/><w:gridCol w:w="2256"/><w:gridCol w:w="2256"/></w:tblGrid>');
  expect(rows(xml)[0]).toContain('<w:tcW w:w="4512" w:type="dxa"/><w:gridSpan w:val="2"/>');
  expect(rows(xml)[1]).toContain('<w:tc><w:tcPr><w:tcW w:w="2256" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">Qty</w:t></w:r></w:p></w:tc>');
});

test('a column inserted to the left of a spanned cell goes before it in every row', async ({ page }) => {
  await openInEditor(page, await buildDocx(TABLE));
  await page.locator('td', { hasText: 'Price' }).click();
  await layout(page, 'Insert column to the left');
  await expect.poll(() => cellTexts(page)).toEqual([['Item', '', 'Price'], ['Paper', '', '5', 'EUR'], ['Pens', '', '', 'USD']]);
});

test('deleting a column: cells in it go, cells spanning it narrow, merges stay right', async ({ page }) => {
  await openInEditor(page, await buildDocx(TABLE));
  await page.locator('td', { hasText: 'EUR' }).click();
  await layout(page, 'Delete column');
  await expect(status(page)).toHaveText('Column deleted – Undo (Ctrl+Z) brings it back');
  await expect.poll(() => cellTexts(page)).toEqual([['Item', 'Price'], ['Paper', '5'], ['Pens', '']]);
  const { xml } = await download(page);
  expect(xml).toContain('<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>');
  expect(rows(xml)[0]).toContain('<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">Price</w:t></w:r></w:p></w:tc>');
  expect(xml).not.toContain('EUR');
});

test('deleting the row where a merge starts starts it in the next row', async ({ page }) => {
  await openInEditor(page, await buildDocx(TABLE));
  await page.locator('td', { hasText: 'Paper' }).click();
  await layout(page, 'Delete row');
  await expect(status(page)).toHaveText('Row deleted – Undo (Ctrl+Z) brings it back');
  await expect(page.locator('table tr')).toHaveCount(2);
  const { xml } = await download(page);
  expect(rows(xml)[1]).toContain('<w:vMerge w:val="restart"/>');
  expect(xml).not.toContain('Paper');
});

test('Delete table removes it, and the caret goes to the paragraph after it', async ({ page }) => {
  await openInEditor(page, await buildDocx(`<w:p><w:r><w:t>Above.</w:t></w:r></w:p>${TABLE}`));
  await page.locator('td', { hasText: 'Pens' }).click();
  await layout(page, 'Delete table');
  await expect(page.locator('table')).toHaveCount(0);
  await expect(status(page)).toHaveText('Table deleted – Undo (Ctrl+Z) brings it back');
  await page.keyboard.type('X');
  await expect(page.locator('.docx-page p').nth(1)).toHaveText('XBelow the table.');
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  await expect(page.getByRole('dialog', { name: 'Review your changes' }).locator('.notes')).toContainText('1 table removed.');
});

test('while tracking, inserted rows are tracked; a deleted row is marked, and Reject brings it back', async ({ page }) => {
  await openInEditor(page, await buildDocx(TABLE));
  await startTracking(page);
  await page.locator('td', { hasText: 'Paper' }).click();
  await layout(page, 'Insert row above');
  await expect(page.locator('tr[data-rev-row="ins"]')).toHaveCount(1);
  await page.locator('td', { hasText: 'EUR' }).click();
  await layout(page, 'Delete row');
  await expect(status(page)).toHaveText('Row marked as deleted (tracked)');
  const deleted = page.locator('tr[data-rev-row="del"]');
  await expect(deleted).toHaveCount(1);
  await expect(page.locator('.changes-group .badge')).toHaveText('2'); // the inserted row, and the deleted one
  let { xml } = await download(page);
  expect(rows(xml)[1]).toMatch(/^<w:tr><w:trPr><w:ins w:id="\d+" w:author="Rita" w:date="[^"]+"\/><\/w:trPr>/);
  expect(rows(xml)[2]).toMatch(/^<w:tr><w:trPr><w:del w:id="\d+" w:author="Rita" w:date="[^"]+"\/><\/w:trPr>/);
  expect(rows(xml)[2]).toMatch(/<w:del w:id="\d+" w:author="Rita" w:date="[^"]+"><w:r><w:delText xml:space="preserve">Paper<\/w:delText><\/w:r><\/w:del>/);
  const ids = [...xml.matchAll(/w:id="(\d+)"/g)].map((m) => m[1]);
  expect(new Set(ids).size).toBe(ids.length);

  // The deleted row is one change: rejecting it brings its text back.
  await page.locator('td', { hasText: 'Pens' }).click();
  await page.getByTitle('Previous change').click();
  await page.getByTitle('Reject the change (or the changes selected) and go to the next').click();
  await expect(page.locator('tr[data-rev-row="del"]')).toHaveCount(0);
  ({ xml } = await download(page));
  expect(rows(xml)[2]).toContain('<w:r><w:t xml:space="preserve">Paper</w:t></w:r>');
  expect(rows(xml)[2]).not.toContain('w:del');
});

test('deleting, while tracking, a row you inserted removes it', async ({ page }) => {
  await openInEditor(page, await buildDocx(TABLE));
  await startTracking(page);
  await page.locator('td', { hasText: 'Pens' }).click();
  await layout(page, 'Insert row below');
  await page.keyboard.type('Ink');
  await layout(page, 'Delete row');
  await expect(page.locator('table tr')).toHaveCount(3);
  await expect(page.locator('tr[data-rev-row]')).toHaveCount(0);
});

test('rows holding someone\'s tracked change aren\'t deleted untracked', async ({ page }) => {
  const table = TABLE.replace('<w:r><w:t>Pens</w:t></w:r>', '<w:ins w:id="90" w:author="Ann"><w:r><w:t>Pens</w:t></w:r></w:ins>');
  await openInEditor(page, await buildDocx(table));
  await page.locator('td', { hasText: 'USD' }).click();
  await layout(page, 'Delete row');
  await expect(status(page)).toHaveText('The rows hold tracked changes: accept or reject them first, or turn on Track changes');
  await expect(page.locator('table tr')).toHaveCount(3);
});
