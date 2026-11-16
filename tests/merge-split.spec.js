// Merging and splitting table cells, as Word's Layout tab does. Merging: the selected cells become
// one (gridSpan across, vMerge down), their content one after another in it. Splitting: a cell into
// columns of equal width (new grid columns where needed) and rows (the row's other cells merged
// down). (Tracked merging and splitting: tracked-merge.spec.js.)
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, canonicalize, documentXmlOf } from './helpers.js';

const CELL = (text, props = '<w:tcW w:w="3000" w:type="dxa"/>') => `<w:tc><w:tcPr>${props}</w:tcPr><w:p>${text ? `<w:r><w:t>${text}</w:t></w:r>` : ''}</w:p></w:tc>`;
const ROW = (...texts) => `<w:tr>${texts.map((t) => CELL(t)).join('')}</w:tr>`;
const TABLE = '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>'
  + `${ROW('A1', 'B1', 'C1')}${ROW('A2', 'B2', 'C2')}${ROW('A3', '', 'C3')}</w:tbl><w:p><w:r><w:t>After.</w:t></w:r></w:p>`;

const status = (page) => page.locator('.statusbar [role=status]');
const rows = (xml) => [...xml.matchAll(/<w:tr>.*?<\/w:tr>|<w:tr [^>]*>.*?<\/w:tr>/g)].map((m) => m[0]);
const texts = (page) => page.locator('table tr').evaluateAll((trs) => trs.map((tr) => [...tr.cells].map((c) => (c.classList.contains('vmerge') ? '^' : '') + [...c.querySelectorAll('p')].map((p) => p.textContent).join('|'))));

async function selectCells(page, from, to) {
  await page.evaluate(([a, b]) => {
    // A cell by its text, or by [row, cell] for an empty one.
    const cell = (t) => (Array.isArray(t) ? document.querySelectorAll('tr')[t[0]].cells[t[1]] : [...document.querySelectorAll('td')].find((td) => td.textContent === t));
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

async function split(page, cols, rowsCount) {
  await layout(page, 'Split cells…');
  const dialog = page.getByRole('dialog', { name: 'Split cells' });
  await dialog.getByRole('spinbutton', { name: 'Number of columns' }).fill(String(cols));
  await dialog.getByRole('spinbutton', { name: 'Number of rows' }).fill(String(rowsCount));
  await dialog.getByRole('button', { name: 'Split' }).click();
}

test('merging cells across a row: one cell spanning them, their text in it', async ({ page }) => {
  await openInEditor(page, await buildDocx(TABLE));
  await page.locator('td', { hasText: 'A1' }).click();
  await page.getByTitle('Table layout: insert or delete rows and columns').click();
  await expect(page.getByRole('menuitem', { name: 'Merge cells' })).toBeDisabled(); // one cell
  await page.keyboard.press('Escape');
  await selectCells(page, 'A1', 'B1');
  await layout(page, 'Merge cells');
  await expect(status(page)).toHaveText('Cells merged');
  await expect.poll(() => texts(page)).toEqual([['A1|B1', 'C1'], ['A2', 'B2', 'C2'], ['A3', '', 'C3']]);
  const { xml } = await download(page);
  expect(rows(xml)[0]).toBe('<w:tr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">A1</w:t></w:r></w:p><w:p><w:r><w:t xml:space="preserve">B1</w:t></w:r></w:p></w:tc>'
    + '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">C1</w:t></w:r></w:p></w:tc></w:tr>');
});

test('merging a block of cells: across and down, empty cells adding nothing; Undo takes it back', async ({ page }) => {
  const docx = await buildDocx(TABLE);
  await openInEditor(page, docx);
  await selectCells(page, 'A2', [2, 1]);
  await layout(page, 'Merge cells');
  await expect.poll(() => texts(page)).toEqual([['A1', 'B1', 'C1'], ['A2|B2|A3', 'C2'], ['^', 'C3']]);
  const { xml } = await download(page);
  expect(rows(xml)[1]).toContain('<w:tcW w:w="6000" w:type="dxa"/><w:gridSpan w:val="2"/><w:vMerge w:val="restart"/>');
  expect(rows(xml)[2]).toMatch(/^<w:tr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"\/><w:gridSpan w:val="2"\/><w:vMerge\/><\/w:tcPr><w:p\/><\/w:tc>/);

  await page.keyboard.press('ControlOrMeta+z');
  await expect.poll(() => texts(page)).toEqual([['A1', 'B1', 'C1'], ['A2', 'B2', 'C2'], ['A3', '', 'C3']]);
  const back = await download(page);
  expect(await canonicalize(page, back.xml)).toEqual(await canonicalize(page, await documentXmlOf(docx)));
});

test('a selection cutting through a merged cell grows to take all of it, as in Word', async ({ page }) => {
  await openInEditor(page, await buildDocx(TABLE.replace(`${CELL('A1')}${CELL('B1')}`, CELL('A1', '<w:tcW w:w="6000" w:type="dxa"/><w:gridSpan w:val="2"/>'))));
  await selectCells(page, 'A1', 'A2'); // A1 spans two columns: B2 comes in too
  await layout(page, 'Merge cells');
  await expect.poll(() => texts(page)).toEqual([['A1|A2|B2', 'C1'], ['^', 'C2'], ['A3', '', 'C3']]);
});

test('splitting a cell into columns adds grid columns that other rows span', async ({ page }) => {
  await openInEditor(page, await buildDocx(TABLE));
  await page.locator('td', { hasText: 'B2' }).click();
  await split(page, 2, 1);
  await expect(status(page)).toHaveText('Cell split into 2 × 1');
  await expect.poll(() => texts(page)).toEqual([['A1', 'B1', 'C1'], ['A2', 'B2', '', 'C2'], ['A3', '', 'C3']]);
  await page.keyboard.type('X'); // the caret stays in the first part
  const { xml } = await download(page);
  expect(xml).toContain('<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/><w:gridCol w:w="3000"/></w:tblGrid>');
  expect(rows(xml)[0]).toContain('<w:tcW w:w="3000" w:type="dxa"/><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">B1');
  expect(rows(xml)[1]).toContain('<w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">XB2</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr><w:p/></w:tc>');
});

test('splitting a merged cell back uses the columns it spans', async ({ page }) => {
  await openInEditor(page, await buildDocx(TABLE.replace(`${CELL('A1')}${CELL('B1')}`, CELL('A1', '<w:tcW w:w="6000" w:type="dxa"/><w:gridSpan w:val="2"/>'))));
  await page.locator('td', { hasText: 'A1' }).click();
  await split(page, 2, 1);
  const { xml } = await download(page);
  expect(xml).toContain('<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>');
  expect(rows(xml)[0]).not.toContain('gridSpan');
});

test('splitting a cell into rows merges the row\'s other cells down; a merge splits back', async ({ page }) => {
  await openInEditor(page, await buildDocx(TABLE));
  await page.locator('td', { hasText: 'B1' }).click();
  await split(page, 1, 2);
  await expect.poll(() => texts(page)).toEqual([['A1', 'B1', 'C1'], ['^', '', '^'], ['A2', 'B2', 'C2'], ['A3', '', 'C3']]);
  let { xml } = await download(page);
  expect(rows(xml)[0]).toContain('<w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">A1');
  expect(rows(xml)[1]).toMatch(/^<w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"\/><w:vMerge\/><\/w:tcPr><w:p\/><\/w:tc><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"\/><\/w:tcPr><w:p\/><\/w:tc>/);

  // The A cell is now merged across two rows: split into 2 rows, the merge is undone.
  await page.locator('td', { hasText: 'A1' }).click();
  await split(page, 1, 2);
  await expect.poll(() => texts(page)).toEqual([['A1', 'B1', 'C1'], ['', '', '^'], ['A2', 'B2', 'C2'], ['A3', '', 'C3']]);
  await page.locator('td', { hasText: 'C1' }).click();
  await split(page, 1, 3);
  await expect(status(page)).toHaveText('This cell is merged across 2 rows: split it into 1 or 2 rows');
});
