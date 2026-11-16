// Tab in the last cell of a table adds a row, as in Word: a copy of the last row's cells, empty,
// with the caret in its first cell; tracked while tracking; one Undo step.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, canonicalize, documentXmlOf } from './helpers.js';

const CELL = (text, props = '<w:tcW w:w="3000" w:type="dxa"/>', pPr = '') => `<w:tc><w:tcPr>${props}</w:tcPr><w:p>${pPr}<w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const TABLE = '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>'
  + `<w:tr><w:trPr><w:tblHeader/></w:trPr>${CELL('Item')}${CELL('Price', '<w:tcW w:w="6000" w:type="dxa"/><w:gridSpan w:val="2"/>')}</w:tr>`
  + `<w:tr>${CELL('Paper', '<w:tcW w:w="3000" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="EEEEEE"/>')}${CELL('5', '<w:tcW w:w="3000" w:type="dxa"/><w:vMerge w:val="restart"/>', '<w:pPr><w:jc w:val="right"/></w:pPr>')}${CELL('EUR')}</w:tr>`
  + '</w:tbl><w:p><w:r><w:t>Below the table.</w:t></w:r></w:p>';

const lastRow = (xml) => [...xml.matchAll(/<w:tr>.*?<\/w:tr>|<w:tr [^>]*>.*?<\/w:tr>/g)].pop()[0];

test('Tab in the last cell adds a row like the last one, and the caret moves into it', async ({ page }) => {
  await openInEditor(page, await buildDocx(TABLE));
  await page.locator('td', { hasText: 'EUR' }).click();
  await page.keyboard.press('End');
  await page.keyboard.press('Tab');
  await expect(page.locator('tr')).toHaveCount(3);
  await page.keyboard.type('Pens');
  await page.keyboard.press('Tab');
  await page.keyboard.type('2');
  await expect(page.locator('tr').nth(2).locator('td')).toHaveText(['Pens', '2', '']);
  const { xml } = await download(page);
  // Same cells (width, shading, alignment), no vertical merge, the typed text in it.
  expect(lastRow(xml)).toBe('<w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="EEEEEE"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">Pens</w:t></w:r></w:p></w:tc>'
    + '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t xml:space="preserve">2</w:t></w:r></w:p></w:tc>'
    + '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p/></w:tc></w:tr>');
});

test('a header row\'s repeat isn\'t copied, and Undo takes the row back', async ({ page }) => {
  const docx = await buildDocx(TABLE.replace(/<w:tr>(?!<w:trPr>)[\s\S]*?<\/w:tr>(?=<\/w:tbl>)/, ''));
  await openInEditor(page, docx);
  await page.locator('td', { hasText: 'Price' }).click();
  await page.keyboard.press('Tab');
  await expect(page.locator('tr')).toHaveCount(2);
  let { xml } = await download(page);
  expect(lastRow(xml)).toMatch(/^<w:tr><w:tc>/); // no <w:trPr><w:tblHeader/>
  expect(lastRow(xml)).toContain('<w:gridSpan w:val="2"/>');

  await page.keyboard.press('ControlOrMeta+z');
  await expect(page.locator('tr')).toHaveCount(1);
  ({ xml } = await download(page));
  expect(await canonicalize(page, xml)).toEqual(await canonicalize(page, await documentXmlOf(docx)));
});

test('while tracking, the new row is a tracked insertion with your name', async ({ page }) => {
  await openInEditor(page, await buildDocx(TABLE));
  await page.getByRole('button', { name: 'Track changes' }).click();
  await page.getByRole('dialog', { name: 'Your name on tracked changes' }).getByRole('textbox').fill('Rita');
  await page.getByRole('button', { name: 'Start tracking' }).click();
  await page.locator('td', { hasText: 'EUR' }).click();
  await page.keyboard.press('End');
  await page.keyboard.press('Tab');
  await page.keyboard.type('Ink');
  const { xml } = await download(page);
  expect(lastRow(xml)).toMatch(/^<w:tr><w:trPr><w:ins w:id="\d+" w:author="Rita" w:date="[^"]+"\/><\/w:trPr><w:tc>/);
  expect(lastRow(xml)).toMatch(/<w:ins w:id="\d+" w:author="Rita" w:date="[^"]+"><w:r><w:t xml:space="preserve">Ink<\/w:t><\/w:r><\/w:ins>/);
  // Ids are unique across the document.
  const ids = [...xml.matchAll(/w:id="(\d+)"/g)].map((m) => m[1]);
  expect(new Set(ids).size).toBe(ids.length);
});
