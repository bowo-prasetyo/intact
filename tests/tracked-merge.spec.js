// Merging and splitting cells while tracking, in the ways a Word document can record them: cells
// merged down a column (<w:cellMerge>, the text moved shown as inserted in the merged cell and
// deleted where it was), cells split into columns (inserted cells, the old grid kept so rejecting
// gives the widths back) and into rows (inserted rows). Each is one change to accept or reject.
// Cells merged across a row can't be recorded: that part is done directly, and the status bar says so.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, canonicalize, documentXmlOf } from './helpers.js';

const CELL = (text) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const TABLE = '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>'
  + `<w:tr>${CELL('A1')}${CELL('B1')}</w:tr><w:tr>${CELL('A2')}${CELL('B2')}</w:tr></w:tbl><w:p><w:r><w:t>After.</w:t></w:r></w:p>`;

const status = (page) => page.locator('.statusbar [role=status]');
const group = (page) => page.getByRole('group', { name: 'Tracked changes' });
const texts = (page) => page.locator('table tr').evaluateAll((trs) => trs.map((tr) => [...tr.cells].map((c) => (c.classList.contains('vmerge') ? '^' : '') + c.textContent)));

async function startTracking(page) {
  await page.getByRole('button', { name: 'Track changes' }).click();
  await page.getByRole('dialog', { name: 'Your name on tracked changes' }).getByRole('textbox').fill('Rita');
  await page.getByRole('button', { name: 'Start tracking' }).click();
}

async function selectCells(page, from, to) {
  await page.evaluate(([a, b]) => {
    const cell = (t) => [...document.querySelectorAll('td')].find((td) => td.textContent === t);
    const range = document.createRange();
    range.setStart(cell(a).querySelector('p'), 0);
    const end = cell(b).querySelector('p');
    range.setEnd(end, end.childNodes.length);
    document.querySelector('.docx-page').focus();
    getSelection().removeAllRanges();
    getSelection().addRange(range);
  }, [from, to]);
}

async function layout(page, item) {
  await page.getByTitle('Table layout: insert or delete rows and columns').click();
  await page.getByRole('menu', { name: 'Table layout' }).getByRole('menuitem', { name: item }).click();
}

async function split(page, cols, rows) {
  await layout(page, 'Split cells…');
  const dialog = page.getByRole('dialog', { name: 'Split cells' });
  await dialog.getByRole('spinbutton', { name: 'Number of columns' }).fill(String(cols));
  await dialog.getByRole('spinbutton', { name: 'Number of rows' }).fill(String(rows));
  await dialog.getByRole('button', { name: 'Split' }).click();
}

test('cells merged down a column while tracking: one change, written as Word records it', async ({ page }) => {
  await openInEditor(page, await buildDocx(TABLE));
  await startTracking(page);
  await selectCells(page, 'A1', 'A2');
  await layout(page, 'Merge cells');
  await expect(status(page)).toHaveText('Cells merged (tracked)');
  await expect(group(page).locator('.badge')).toHaveText('1');
  await expect(page.locator('td[data-rev-merge]')).toHaveCount(2);
  const { xml } = await download(page);
  expect(xml).toMatch(/<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"\/><w:vMerge w:val="restart"\/><w:cellMerge w:id="\d+" w:author="Rita" w:date="[^"]+" w:vMerge="rest"\/><\/w:tcPr><w:p><w:r><w:t xml:space="preserve">A1<\/w:t><\/w:r><\/w:p><w:p><w:pPr><w:rPr><w:ins w:id="\d+" w:author="Rita" w:date="[^"]+"\/><\/w:rPr><\/w:pPr><w:ins w:id="\d+" w:author="Rita" w:date="[^"]+"><w:r><w:t xml:space="preserve">A2<\/w:t><\/w:r><\/w:ins><\/w:p><\/w:tc>/);
  expect(xml).toMatch(/<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"\/><w:vMerge\/><w:cellMerge w:id="\d+" w:author="Rita" w:date="[^"]+" w:vMerge="cont"\/><\/w:tcPr><w:p><w:del w:id="\d+" w:author="Rita" w:date="[^"]+"><w:r><w:delText xml:space="preserve">A2<\/w:delText><\/w:r><\/w:del><\/w:p><\/w:tc>/);
});

test('rejecting a tracked merge puts the cells and their text back exactly', async ({ page }) => {
  const docx = await buildDocx(TABLE);
  await openInEditor(page, docx);
  await startTracking(page);
  await selectCells(page, 'B1', 'B2');
  await layout(page, 'Merge cells');
  await page.locator('td', { hasText: 'A2' }).click();
  await group(page).getByTitle('Previous change').click();
  await group(page).getByRole('button', { name: 'Reject' }).first().click();
  await expect(status(page)).toContainText('Rejected the merged cells');
  await expect.poll(() => texts(page)).toEqual([['A1', 'B1'], ['A2', 'B2']]);
  const { xml } = await download(page);
  expect(await canonicalize(page, xml)).toEqual(await canonicalize(page, await documentXmlOf(docx)));
});

test('accepting a tracked merge leaves the merged cell with all the text', async ({ page }) => {
  await openInEditor(page, await buildDocx(TABLE));
  await startTracking(page);
  await selectCells(page, 'A1', 'A2');
  await layout(page, 'Merge cells');
  await page.getByTitle('More ways to accept').click();
  await page.getByRole('menuitem', { name: 'Accept all changes' }).click();
  await expect.poll(() => texts(page)).toEqual([['A1A2', 'B1'], ['^', 'B2']]);
  const { xml } = await download(page);
  expect(xml).not.toMatch(/cellMerge|w:ins |w:del /);
  expect(xml).toContain('<w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">A1</w:t></w:r></w:p><w:p><w:r><w:t xml:space="preserve">A2</w:t></w:r></w:p></w:tc>');
});

test('merging across a row while tracking is done directly, and the status bar says why', async ({ page }) => {
  await openInEditor(page, await buildDocx(TABLE));
  await startTracking(page);
  await selectCells(page, 'A1', 'B1');
  await layout(page, 'Merge cells');
  await expect(status(page)).toHaveText('Cells merged – merging across a row can’t be tracked in a Word document, so it isn’t marked');
  await expect(page.locator('td[data-rev-merge]')).toHaveCount(0);
});

test('a cell split into columns while tracking: inserted cells; rejecting gives the old widths back', async ({ page }) => {
  const docx = await buildDocx(TABLE);
  await openInEditor(page, docx);
  await startTracking(page);
  await page.locator('td', { hasText: 'B2' }).click();
  await split(page, 2, 1);
  await expect(status(page)).toHaveText('Cell split into 2 × 1 (tracked)');
  await expect(page.locator('td[data-rev-cell="ins"]')).toHaveCount(1);
  await expect(group(page).locator('.badge')).toHaveText('1'); // the grid change is part of it
  let { xml } = await download(page);
  expect(xml).toMatch(/<w:tblGrid><w:gridCol w:w="3000"\/><w:gridCol w:w="1500"\/><w:gridCol w:w="1500"\/><w:tblGridChange w:id="\d+" w:author="Rita" w:date="[^"]+"><w:tblGrid><w:gridCol w:w="3000"\/><w:gridCol w:w="3000"\/><\/w:tblGrid><\/w:tblGridChange><\/w:tblGrid>/);
  expect(xml).toMatch(/<w:cellIns w:id="\d+" w:author="Rita"/);

  await page.locator('td[data-rev-cell="ins"]').click();
  await group(page).getByRole('button', { name: 'Reject' }).first().click();
  ({ xml } = await download(page));
  expect(await canonicalize(page, xml)).toEqual(await canonicalize(page, await documentXmlOf(docx)));
});

test('a cell split into rows while tracking: the new rows are tracked insertions', async ({ page }) => {
  await openInEditor(page, await buildDocx(TABLE));
  await startTracking(page);
  await page.locator('td', { hasText: 'A1' }).click();
  await split(page, 1, 2);
  await expect(page.locator('tr[data-rev-row="ins"]')).toHaveCount(1);
  await page.getByTitle('More ways to reject').click();
  await page.getByRole('menuitem', { name: 'Reject all changes' }).click();
  await expect.poll(() => texts(page)).toEqual([['A1', 'B1'], ['A2', 'B2']]);
});

test('a merged cell split back into rows while tracking is a tracked split, undone by Reject', async ({ page }) => {
  const merged = TABLE.replace(`${CELL('A1')}`, '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>')
    .replace(`${CELL('A2')}`, '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:vMerge/></w:tcPr><w:p/></w:tc>');
  const docx = await buildDocx(merged);
  await openInEditor(page, docx);
  await startTracking(page);
  await page.locator('td', { hasText: 'A1' }).click();
  await split(page, 1, 2);
  await expect.poll(() => texts(page)).toEqual([['A1', 'B1'], ['', 'B2']]);
  const { xml } = await download(page);
  expect(xml).toMatch(/<w:cellMerge w:id="\d+" w:author="Rita" w:date="[^"]+" w:vMerge="rest" w:vMergeOrig="cont"\/>/);
  await page.locator('td', { hasText: 'B2' }).click();
  await group(page).getByTitle('Previous change').click();
  await expect(status(page)).toContainText('The split cells');
  await group(page).getByRole('button', { name: 'Reject' }).first().click();
  await expect.poll(() => texts(page)).toEqual([['A1', 'B1'], ['^', 'B2']]);
  const back = await download(page);
  expect(await canonicalize(page, back.xml)).toEqual(await canonicalize(page, await documentXmlOf(docx)));
});
