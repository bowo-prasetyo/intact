// Tables whose rows or cells are fields (content controls), as forms and templates have them: rows
// and columns change through the fields. A row added in a repeating section is a new item of it (as
// Word's + adds one), a row added in a plain field stays in it, a field left empty goes, and new
// cells next to a field are ordinary cells. Locked fields and fields linked to data stay as they are.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const W15 = 'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"';
const CELL = (text) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const GRID = '<w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>';
const item = (id, row) => `<w:sdt><w:sdtPr><w:id w:val="${id}"/><w15:repeatingSectionItem ${W15}/></w:sdtPr><w:sdtContent>${row}</w:sdtContent></w:sdt>`;
const REPEATING = `<w:tbl>${GRID}<w:tr>${CELL('Item')}${CELL('Qty')}</w:tr>`
  + `<w:sdt><w:sdtPr><w:alias w:val="Lines"/><w:id w:val="10"/><w15:repeatingSection ${W15}/></w:sdtPr><w:sdtContent>`
  + `${item(11, `<w:tr>${CELL('Paper')}${CELL('5')}</w:tr>`)}${item(12, `<w:tr>${CELL('Pens')}${CELL('2')}</w:tr>`)}`
  + '</w:sdtContent></w:sdt></w:tbl><w:p/>';
const field = (inner, extra = '') => `<w:sdt><w:sdtPr><w:alias w:val="Rows"/><w:id w:val="20"/>${extra}</w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

const status = (page) => page.locator('.statusbar [role=status]');
const texts = (page) => page.locator('table tr').evaluateAll((trs) => trs.map((tr) => [...tr.cells].map((c) => c.textContent)));

async function layout(page, item2) {
  await page.getByTitle('Table layout: insert or delete rows and columns').click();
  await page.getByRole('menu', { name: 'Table layout' }).getByRole('menuitem', { name: item2 }).click();
}

test('a row inserted in a repeating section is a new item of it, with its own id', async ({ page }) => {
  await openInEditor(page, await buildDocx(REPEATING));
  await page.locator('td', { hasText: 'Paper' }).click();
  await layout(page, 'Insert row below');
  await expect(status(page)).toHaveText('Row inserted below');
  await page.keyboard.type('Ink');
  await expect.poll(() => texts(page)).toEqual([['Item', 'Qty'], ['Paper', '5'], ['Ink', ''], ['Pens', '2']]);
  const { xml } = await download(page);
  const items = [...xml.matchAll(/<w:sdt><w:sdtPr><w:id w:val="(-?\d+)"\/><w15:repeatingSectionItem[^>]*\/><\/w:sdtPr><w:sdtContent><w:tr>.*?<\/w:tr><\/w:sdtContent><\/w:sdt>/g)];
  expect(items).toHaveLength(3);
  expect(items[1][0]).toContain('Ink');
  expect(new Set(items.map((m) => m[1])).size).toBe(3);
});

test('deleting a repeating section\'s row deletes its item; Tab in the last cell adds one', async ({ page }) => {
  await openInEditor(page, await buildDocx(REPEATING));
  await page.locator('td', { hasText: 'Paper' }).click();
  await layout(page, 'Delete row');
  await expect.poll(() => texts(page)).toEqual([['Item', 'Qty'], ['Pens', '2']]);
  await page.locator('td', { hasText: '2' }).click();
  await page.keyboard.press('End');
  await page.keyboard.press('Tab');
  await page.keyboard.type('Clips');
  const { xml } = await download(page);
  expect(xml).not.toContain('Paper');
  expect([...xml.matchAll(/repeatingSectionItem/g)]).toHaveLength(2);
  expect(xml).toMatch(/<w15:repeatingSectionItem[^>]*\/><\/w:sdtPr><w:sdtContent><w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"\/><\/w:tcPr><w:p><w:r><w:t xml:space="preserve">Clips/);
});

test('a row added to rows in a plain field stays in the field; a field left empty goes', async ({ page }) => {
  await openInEditor(page, await buildDocx(`<w:tbl>${GRID}<w:tr>${CELL('Head')}${CELL('Er')}</w:tr>${field(`<w:tr>${CELL('Line')}${CELL('One')}</w:tr>`)}</w:tbl><w:p/>`));
  await page.locator('td', { hasText: 'One' }).click();
  await page.keyboard.press('End');
  await page.keyboard.press('Tab');
  await page.keyboard.type('Two');
  let { xml } = await download(page);
  expect(xml).toMatch(/<w:sdtContent><w:tr>.*?Line.*?<\/w:tr><w:tr>.*?Two.*?<\/w:tr><\/w:sdtContent><\/w:sdt><\/w:tbl>/);

  for (const text of ['Line', 'Two']) {
    await page.locator('td', { hasText: text }).click();
    await layout(page, 'Delete row');
  }
  ({ xml } = await download(page));
  expect(xml).not.toContain('<w:sdt>');
  expect(xml).toContain('Head');
});

test('columns: a new cell next to a field cell is an ordinary cell; a deleted field cell takes its field', async ({ page }) => {
  const cellField = `<w:sdt><w:sdtPr><w:alias w:val="Price"/><w:id w:val="30"/><w:text/></w:sdtPr><w:sdtContent>${CELL('9.99')}</w:sdtContent></w:sdt>`;
  await openInEditor(page, await buildDocx(`<w:tbl>${GRID}<w:tr>${CELL('Pen')}${cellField}</w:tr><w:tr>${CELL('Ink')}${CELL('3.50')}</w:tr></w:tbl><w:p/>`));
  await page.locator('td', { hasText: 'Pen' }).click();
  await layout(page, 'Insert column to the right');
  await expect.poll(() => texts(page)).toEqual([['Pen', '', '9.99'], ['Ink', '', '3.50']]);
  let { xml } = await download(page);
  expect(xml).toMatch(/Pen<\/w:t><\/w:r><\/w:p><\/w:tc><w:tc><w:tcPr><w:tcW w:w="\d+" w:type="dxa"\/><\/w:tcPr><w:p\/><\/w:tc><w:sdt>/);

  await page.locator('td', { hasText: '9.99' }).click();
  await layout(page, 'Delete column');
  await expect.poll(() => texts(page)).toEqual([['Pen', ''], ['Ink', '']]);
  ({ xml } = await download(page));
  expect(xml).not.toContain('<w:sdt>');
});

test('rows in a locked field, or one linked to data, stay as they are', async ({ page }) => {
  await openInEditor(page, await buildDocx(`<w:tbl>${GRID}${field(`<w:tr>${CELL('Fixed')}${CELL('Row')}</w:tr>`, '<w:lock w:val="sdtContentLocked"/>')}</w:tbl><w:p/>`));
  await page.locator('td', { hasText: 'Fixed' }).click();
  await layout(page, 'Insert row below');
  await expect(status(page)).toHaveText('Read-only: Content control “Rows” – locked in Word');
  await expect(page.locator('table tr')).toHaveCount(1);
});

// Only the rows and cells a change touches count: a locked row (a form's total, say) or a cell
// linked to data doesn't stop changes elsewhere in the table.
const cellField = (name, inner, extra) => `<w:sdt><w:sdtPr><w:alias w:val="${name}"/><w:id w:val="30"/>${extra}</w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
const PARTLY = `<w:tbl>${GRID}<w:tr>${CELL('Item')}${CELL('Qty')}</w:tr><w:tr>${CELL('Paper')}${CELL('5')}</w:tr>`
  + `${field(`<w:tr>${CELL('Total')}${CELL('5')}</w:tr>`, '<w:lock w:val="sdtContentLocked"/>')}</w:tbl><w:p/>`;

test('a locked row stops only changes to it: rows elsewhere are inserted and deleted', async ({ page }) => {
  await openInEditor(page, await buildDocx(PARTLY));
  await page.locator('td', { hasText: 'Paper' }).click();
  await layout(page, 'Insert row below');
  await expect(status(page)).toHaveText('Row inserted below');
  await expect.poll(() => texts(page)).toEqual([['Item', 'Qty'], ['Paper', '5'], ['', ''], ['Total', '5']]);
  await page.locator('td', { hasText: 'Item' }).click();
  await layout(page, 'Delete row');
  await expect(status(page)).toHaveText('Row deleted – Undo (Ctrl+Z) brings it back');
  await expect.poll(() => texts(page)).toEqual([['Paper', '5'], ['', ''], ['Total', '5']]);
  const { xml } = await download(page);
  expect(xml).toContain('<w:lock w:val="sdtContentLocked"/></w:sdtPr><w:sdtContent><w:tr>');
  expect(xml).toContain('Total');
});

test('a column would add a cell to the locked row, so it is refused, naming the field', async ({ page }) => {
  await openInEditor(page, await buildDocx(PARTLY));
  await page.locator('td', { hasText: 'Paper' }).click();
  await layout(page, 'Insert column to the right');
  await expect(status(page)).toHaveText('This would change the field “Rows”, which is locked: change it in Word');
  await expect.poll(() => texts(page)).toEqual([['Item', 'Qty'], ['Paper', '5'], ['Total', '5']]);
});

test('a cell linked to data: columns beside it change, its own column does not', async ({ page }) => {
  const bound = '<w:dataBinding w:xpath="/order/qty" w:storeItemID="{11111111-1111-1111-1111-111111111111}"/>';
  await openInEditor(page, await buildDocx(`<w:tbl>${GRID}<w:tr>${CELL('Item')}${CELL('Qty')}</w:tr>`
    + `<w:tr>${CELL('Paper')}${cellField('Quantity', CELL('5'), bound)}</w:tr></w:tbl><w:p/>`));
  await page.locator('td', { hasText: 'Item' }).click();
  await layout(page, 'Insert column to the left');
  await expect.poll(() => texts(page)).toEqual([['', 'Item', 'Qty'], ['', 'Paper', '5']]);
  await page.locator('td', { hasText: 'Qty' }).click();
  await layout(page, 'Delete column');
  await expect(status(page)).toHaveText('This would change the field “Quantity”, which is linked to the document’s data: change it in Word');
  await page.locator('td', { hasText: 'Item' }).click();
  await layout(page, 'Delete column');
  await expect.poll(() => texts(page)).toEqual([['', 'Qty'], ['', '5']]);
  const { xml } = await download(page);
  expect(xml).toContain(bound);
});
