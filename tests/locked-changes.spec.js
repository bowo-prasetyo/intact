// Tracked changes inside content Intact keeps as it is (locked fields, fields linked to data, locked
// cells, fields like PAGE): they can't be shown one by one, so each such piece is one change, and
// Accept or Reject resolves every change in it at once, in its XML, as Word's Accept All in it would.
// The piece stays as it was otherwise: still locked, still linked.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const by = (local, id, inner) => `<w:${local} w:id="${id}" w:author="Ann" w:date="2026-03-0${id}T00:00:00Z">${inner}</w:${local}>`;
const run = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const delRun = (t) => `<w:r><w:delText xml:space="preserve">${t}</w:delText></w:r>`;
const LOCKED = (body, extra = '<w:lock w:val="contentLocked"/>') => `<w:sdt><w:sdtPr><w:alias w:val="Clause"/><w:id w:val="5"/>${extra}</w:sdtPr><w:sdtContent>${body}</w:sdtContent></w:sdt>`;
const CLAUSE = LOCKED(`<w:p>${run('The fee is ')}${by('del', 1, delRun('EUR 1,000'))}${by('ins', 2, run('EUR 1,200'))}${run(' a month.')}</w:p>`);
const DOC = `<w:p>${run('Before.')}</w:p>${CLAUSE}<w:p>${run('After.')}</w:p>`;

const status = (page) => page.locator('.statusbar [role=status]');
const group = (page) => page.getByRole('group', { name: 'Tracked changes' });

test('a locked field with tracked changes is one change; Reject undoes them all, the lock stays', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await expect(group(page).locator('.badge')).toHaveText('1');
  await page.getByText('Before.').click();
  await group(page).getByTitle('Next change').click();
  await expect(status(page)).toHaveText('The 2 tracked changes in the locked field “Clause” by Ann');
  await group(page).getByRole('button', { name: 'Reject' }).first().click();
  await expect(status(page)).toHaveText('Rejected the 2 tracked changes in the locked field “Clause” by Ann. No more tracked changes');
  await expect(page.locator('.docx-page .sdt')).toContainText('The fee is EUR 1,000 a month.');
  const { xml } = await download(page);
  expect(xml).toContain('<w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t xml:space="preserve">The fee is </w:t></w:r><w:r><w:t xml:space="preserve">EUR 1,000</w:t></w:r><w:r><w:t xml:space="preserve"> a month.</w:t></w:r></w:p></w:sdtContent>');
});

test('accepting changes in a field linked to data keeps the link; Undo brings the changes back', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC.replace('<w:lock w:val="contentLocked"/>', '<w:dataBinding w:xpath="/root/clause" w:storeItemID="{00000000-0000-0000-0000-000000000001}"/>')));
  await page.getByText('Before.').click();
  await group(page).getByTitle('Next change').click();
  await expect(status(page)).toHaveText('The 2 tracked changes in the field linked to data “Clause” by Ann');
  await group(page).getByRole('button', { name: 'Accept' }).first().click();
  let { xml } = await download(page);
  expect(xml).toContain('<w:dataBinding w:xpath="/root/clause"');
  expect(xml).toContain('EUR 1,200');
  expect(xml).not.toContain('EUR 1,000');
  await page.keyboard.press('ControlOrMeta+z');
  await expect(group(page).locator('.badge')).toHaveText('1');
  ({ xml } = await download(page));
  expect(xml).toContain('<w:delText xml:space="preserve">EUR 1,000</w:delText>');
});

test('a deleted paragraph break in a locked field: accepted, the paragraphs join', async ({ page }) => {
  const field = LOCKED(`<w:p><w:pPr><w:rPr>${by('del', 3, '')}</w:rPr></w:pPr>${run('One ')}</w:p><w:p>${run('two.')}</w:p>`);
  await openInEditor(page, await buildDocx(`${field}<w:p/>`));
  await page.getByTitle('More ways to accept').click();
  await page.getByRole('menuitem', { name: 'Accept all changes' }).click();
  const { xml } = await download(page);
  expect(xml).toContain('<w:sdtContent><w:p><w:r><w:t xml:space="preserve">One </w:t></w:r><w:r><w:t xml:space="preserve">two.</w:t></w:r></w:p></w:sdtContent>');
});

test('a locked cell with tracked changes is one change too', async ({ page }) => {
  const cell = `<w:sdt><w:sdtPr><w:alias w:val="Total"/><w:lock w:val="sdtContentLocked"/><w:id w:val="9"/></w:sdtPr><w:sdtContent><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p>${by('ins', 4, run('42'))}</w:p></w:tc></w:sdtContent></w:sdt>`;
  await openInEditor(page, await buildDocx(`<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p>${run('Sum')}</w:p></w:tc>${cell}</w:tr></w:tbl><w:p/>`));
  await expect(group(page).locator('.badge')).toHaveText('1');
  await page.locator('td', { hasText: 'Sum' }).click();
  await group(page).getByTitle('Next change').click();
  await expect(status(page)).toHaveText('The tracked change in a locked cell by Ann');
  await group(page).getByRole('button', { name: 'Reject' }).first().click();
  await expect(page.locator('td').nth(1)).toHaveText('');
  const { xml } = await download(page);
  expect(xml).toContain('<w:lock w:val="sdtContentLocked"/><w:id w:val="9"/></w:sdtPr><w:sdtContent><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p/></w:tc></w:sdtContent>');
});

test('a field kept as it is (a simple PAGE field) with a tracked insertion is resolved the same way', async ({ page }) => {
  await openInEditor(page, await buildDocx(`<w:p>${run('Page ')}<w:fldSimple w:instr=" PAGE ">${by('ins', 5, run('1'))}</w:fldSimple></w:p>`));
  await expect(group(page).locator('.badge')).toHaveText('1');
  await page.getByTitle('More ways to accept').click();
  await page.getByRole('menuitem', { name: 'Accept all changes' }).click();
  const { xml } = await download(page);
  expect(xml).toContain('<w:fldSimple w:instr=" PAGE "><w:r><w:t xml:space="preserve">1</w:t></w:r></w:fldSimple>');
});
