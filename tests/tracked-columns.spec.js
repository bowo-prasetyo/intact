// Columns inserted and deleted while tracking, as Word tracks them: each cell of the column is
// marked (<w:cellIns>, <w:cellDel>; a deleted cell's text in <w:del>), and the column is one change
// to accept or reject. Rejecting an inserted column, or accepting a deleted one, takes its cells and
// grid column out; cells spanning it narrow.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, canonicalize, documentXmlOf } from './helpers.js';

const CELL = (text, props = '<w:tcW w:w="3000" w:type="dxa"/>') => `<w:tc><w:tcPr>${props}</w:tcPr><w:p>${text ? `<w:r><w:t>${text}</w:t></w:r>` : ''}</w:p></w:tc>`;
const TABLE = '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>'
  + `<w:tr>${CELL('Title', '<w:tcW w:w="6000" w:type="dxa"/><w:gridSpan w:val="2"/>')}</w:tr>`
  + `<w:tr>${CELL('A2')}${CELL('B2')}</w:tr><w:tr>${CELL('A3')}${CELL('B3')}</w:tr></w:tbl><w:p><w:r><w:t>After.</w:t></w:r></w:p>`;

const status = (page) => page.locator('.statusbar [role=status]');
const rows = (xml) => [...xml.matchAll(/<w:tr>.*?<\/w:tr>|<w:tr [^>]*>.*?<\/w:tr>/g)].map((m) => m[0]);
const texts = (page) => page.locator('table tr').evaluateAll((trs) => trs.map((tr) => [...tr.cells].map((c) => c.textContent)));

async function startTracking(page) {
  await page.getByRole('button', { name: 'Track changes' }).click();
  await page.getByRole('dialog', { name: 'Your name on tracked changes' }).getByRole('textbox').fill('Rita');
  await page.getByRole('button', { name: 'Start tracking' }).click();
}

async function layout(page, item) {
  await page.getByTitle('Table layout: insert or delete rows and columns').click();
  await page.getByRole('menu', { name: 'Table layout' }).getByRole('menuitem', { name: item }).click();
}

test('a column inserted while tracking: each new cell is a tracked insertion, the column one change', async ({ page }) => {
  await openInEditor(page, await buildDocx(TABLE));
  await startTracking(page);
  await page.locator('td', { hasText: 'A2' }).click();
  await layout(page, 'Insert column to the right');
  await expect(status(page)).toHaveText('Column inserted to the right (tracked)');
  await expect(page.locator('td[data-rev-cell="ins"]')).toHaveCount(2); // the title cell widens instead
  await expect(page.locator('.changes-group .badge')).toHaveText('1');
  const { xml } = await download(page);
  expect(rows(xml)[1]).toMatch(/<w:tc><w:tcPr><w:tcW w:w="\d+" w:type="dxa"\/><w:cellIns w:id="\d+" w:author="Rita" w:date="[^"]+"\/><\/w:tcPr><w:p\/><\/w:tc>/);
  expect(rows(xml)[0]).toContain('<w:gridSpan w:val="3"/>');

  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  await expect(page.getByRole('dialog', { name: 'Review your changes' }).locator('.notes')).toContainText('Recorded as tracked changes (2 table cells)');
});

test('rejecting an inserted column takes it out again, exactly', async ({ page }) => {
  const docx = await buildDocx(TABLE);
  await openInEditor(page, docx);
  await startTracking(page);
  await page.locator('td', { hasText: 'B3' }).click();
  await layout(page, 'Insert column to the left');
  await page.locator('td[data-rev-cell="ins"]').first().click();
  await page.getByTitle('Reject the change (or the changes selected) and go to the next').click();
  await expect(status(page)).toContainText('Rejected the inserted column');
  await expect.poll(() => texts(page)).toEqual([['Title'], ['A2', 'B2'], ['A3', 'B3']]);
  const { xml } = await download(page);
  expect(await canonicalize(page, xml)).toEqual(await canonicalize(page, await documentXmlOf(docx)));
});

test('a column deleted while tracking is marked; accepting it removes the cells, the title narrows', async ({ page }) => {
  await openInEditor(page, await buildDocx(TABLE));
  await startTracking(page);
  await page.locator('td', { hasText: 'B2' }).click();
  await layout(page, 'Delete column');
  await expect(status(page)).toHaveText('Column marked as deleted (tracked)');
  await expect(page.locator('td[data-rev-cell="del"]')).toHaveCount(2);
  let { xml } = await download(page);
  expect(rows(xml)[1]).toMatch(/<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"\/><w:cellDel w:id="\d+" w:author="Rita" w:date="[^"]+"\/><\/w:tcPr><w:p><w:del w:id="\d+" w:author="Rita" w:date="[^"]+"><w:r><w:delText xml:space="preserve">B2<\/w:delText><\/w:r><\/w:del><\/w:p><\/w:tc>/);
  const ids = [...xml.matchAll(/w:id="(\d+)"/g)].map((m) => m[1]);
  expect(new Set(ids).size).toBe(ids.length);

  await page.locator('td', { hasText: 'A2' }).click();
  await page.getByTitle('Next change').click();
  await page.getByTitle('Accept the change (or the changes selected) and go to the next').click();
  await expect.poll(() => texts(page)).toEqual([['Title'], ['A2'], ['A3']]);
  ({ xml } = await download(page));
  expect(xml).toContain('<w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>');
  expect(rows(xml)[0]).toContain('<w:tcW w:w="3000" w:type="dxa"/></w:tcPr>');
  expect(rows(xml)[0]).not.toContain('gridSpan');
});

test('rejecting a deleted column brings its text back', async ({ page }) => {
  await openInEditor(page, await buildDocx(TABLE));
  await startTracking(page);
  await page.locator('td', { hasText: 'A3' }).click();
  await layout(page, 'Delete column');
  await page.getByTitle('More ways to reject').click();
  await page.getByRole('menuitem', { name: 'Reject all changes' }).click();
  await expect(page.locator('td[data-rev-cell]')).toHaveCount(0);
  const { xml } = await download(page);
  expect(rows(xml)[2]).toContain('<w:r><w:t>A3</w:t></w:r>'); // all rejected: the table as it came
  expect(xml).not.toContain('cellDel');
  expect(xml).not.toContain('w:delText');
});

test('deleting, while tracking, a column you inserted removes it', async ({ page }) => {
  await openInEditor(page, await buildDocx(TABLE));
  await startTracking(page);
  await page.locator('td', { hasText: 'B2' }).click();
  await layout(page, 'Insert column to the right');
  await layout(page, 'Delete column');
  await expect(page.locator('td[data-rev-cell]')).toHaveCount(0);
  await expect.poll(() => texts(page)).toEqual([['Title'], ['A2', 'B2'], ['A3', 'B3']]);
});

test('a column Word deleted with tracking is one change, accepted here', async ({ page }) => {
  const ann = '<w:cellDel w:id="91" w:author="Ann" w:date="2026-01-05T10:00:00Z"/>';
  await openInEditor(page, await buildDocx(TABLE.replace(CELL('B2'), CELL('B2', `<w:tcW w:w="3000" w:type="dxa"/>${ann}`)).replace(CELL('B3'), CELL('B3', `<w:tcW w:w="3000" w:type="dxa"/>${ann.replace('91', '92')}`))));
  await expect(page.locator('.changes-group .badge')).toHaveText('1');
  await page.locator('td', { hasText: 'B3' }).click();
  await page.getByTitle('Accept the change (or the changes selected) and go to the next').click();
  await expect(status(page)).toContainText('Accepted the deleted column by Ann');
  await expect.poll(() => texts(page)).toEqual([['Title'], ['A2'], ['A3']]);
});
