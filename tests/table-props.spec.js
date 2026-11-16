// Tracked changes to a table's, row's or cell's properties (as Word records a change of table
// style, row height or cell shading): each is a change to accept (its record goes) or reject (the
// old properties come back), like any other.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const by = (local, id, inner) => `<w:${local} w:id="${id}" w:author="Gus" w:date="2026-02-0${id}T00:00:00Z">${inner}</w:${local}>`;
const TABLE = '<w:tbl>'
  + `<w:tblPr><w:tblStyle w:val="Fancy"/><w:tblW w:w="5000" w:type="pct"/>${by('tblPrChange', 1, '<w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>')}</w:tblPr>`
  + `<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="2000"/>${by('tblGridChange', 2, '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>')}</w:tblGrid>`
  + `<w:tr><w:trPr><w:trHeight w:val="800"/>${by('trPrChange', 3, '<w:trPr/>')}</w:trPr>`
  + `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>${by('tcPrChange', 4, '<w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>')}</w:tcPr><w:p><w:r><w:t>Yellow</w:t></w:r></w:p></w:tc>`
  + '<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Plain</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>';

const status = (page) => page.locator('.statusbar [role=status]');
const group = (page) => page.getByRole('group', { name: 'Tracked changes' });

test('table, row and cell formatting changes are listed and shown; none is left for Word', async ({ page }) => {
  await openInEditor(page, await buildDocx(TABLE));
  await expect(group(page).locator('.badge')).toHaveText('3');
  await expect(group(page).locator('.badge')).toHaveAttribute('title', '3 tracked changes to accept or reject here');
  await expect(page.locator('table')).toHaveAttribute('title', 'Table formatting changed by Gus on 2026-02-01 – tracked change: Accept or Reject on the toolbar');
  await expect(page.locator('tr')).toHaveAttribute('data-fmt-row', '1');
  await expect(page.locator('td', { hasText: 'Yellow' })).toHaveAttribute('title', 'Cell formatting changed by Gus on 2026-02-04 – tracked change: Accept or Reject on the toolbar');
  await page.locator('.docx-page > p').last().click();
  await group(page).getByTitle('Next change').click();
  await expect(status(page)).toHaveText('The table formatting change by Gus, 2026-02-01');
  await group(page).getByTitle('Next change').click();
  await expect(status(page)).toHaveText('The cell formatting change by Gus, 2026-02-04');
  await group(page).getByTitle('Next change').click();
  await expect(status(page)).toHaveText('The row formatting change by Gus, 2026-02-03');
});

test('rejecting a cell\'s change brings its old properties back; accepting a row\'s keeps the new', async ({ page }) => {
  await openInEditor(page, await buildDocx(TABLE));
  await page.locator('td', { hasText: 'Yellow' }).click();
  await group(page).getByRole('button', { name: 'Reject' }).first().click();
  await expect(status(page)).toContainText('Rejected the cell formatting change by Gus');
  await expect(page.locator('td', { hasText: 'Yellow' })).not.toHaveAttribute('data-fmt-cell', '1');
  expect(await page.locator('td', { hasText: 'Yellow' }).evaluate((td) => td.style.backgroundColor)).toBe('');
  await page.locator('td', { hasText: 'Plain' }).click(); // in the row: its change
  await group(page).getByRole('button', { name: 'Accept' }).first().click();
  await expect(status(page)).toContainText('Accepted the row formatting change by Gus');
  const { xml } = await download(page);
  expect(xml).toContain('<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">Yellow</w:t>');
  expect(xml).toContain('<w:tr><w:trPr><w:trHeight w:val="800"/></w:trPr>');
  expect(xml).toContain('tblPrChange'); // still to decide
});

test('rejecting a table\'s change brings back its old properties and grid', async ({ page }) => {
  await openInEditor(page, await buildDocx(TABLE));
  await page.getByTitle('More ways to reject').click();
  await page.getByRole('menuitem', { name: 'Reject all changes' }).click();
  await expect(status(page)).toHaveText('Rejected all 3 changes');
  await expect(page.locator('table')).not.toHaveAttribute('data-tstyle', /./);
  const { xml } = await download(page);
  expect(xml).toContain('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:trPr/>');
  expect(xml).not.toMatch(/PrChange|GridChange/);
});
