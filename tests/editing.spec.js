// Editing behaviour a user meets in the first minute: list numbering and undo/redo.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const NUM = '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>';
const LIST = `<w:p>${NUM}<w:r><w:t>Alpha</w:t></w:r></w:p><w:p>${NUM}<w:r><w:t>Beta</w:t></w:r></w:p><w:p><w:r><w:t>After the list</w:t></w:r></w:p>`;

const markers = (page) => page.locator('.docx-page .list').evaluateAll((els) => els.map((e) => `${e.dataset.marker} ${e.textContent}`));

test('a new list item gets the next number', async ({ page }) => {
  await openInEditor(page, await buildDocx(LIST));
  await page.locator('.docx-page p', { hasText: 'Beta' }).click();
  await page.keyboard.press('End');
  await page.keyboard.press('Enter');
  await page.keyboard.type('Gamma');
  await expect.poll(() => markers(page)).toEqual(['1. Alpha', '2. Beta', '3. Gamma']);
});

test('deleting a list item renumbers the rest', async ({ page }) => {
  await openInEditor(page, await buildDocx(`<w:p>${NUM}<w:r><w:t>A</w:t></w:r></w:p>${LIST}`));
  // Select the whole first item and its paragraph break, then delete it.
  await page.locator('.docx-page p', { hasText: /^A$/ }).click();
  await page.keyboard.press('Home');
  await page.keyboard.press('Shift+ArrowDown');
  await page.keyboard.press('Delete');
  await expect.poll(() => markers(page)).toEqual(['1. Alpha', '2. Beta']);
});

test('undo and redo cover paragraph style changes', async ({ page }) => {
  await openInEditor(page, await buildDocx(LIST));
  const first = page.locator('.docx-page > *').first();
  await page.locator('.docx-page p', { hasText: 'After the list' }).click();
  await page.getByLabel('Paragraph style').selectOption('h1');
  const block = page.locator('.docx-page > *', { hasText: 'After the list' });
  await expect(block).toHaveJSProperty('tagName', 'H1');

  await page.keyboard.press('ControlOrMeta+z');
  await expect(block).toHaveJSProperty('tagName', 'P');
  await expect(page.getByLabel('Paragraph style')).toHaveValue('normal');

  await page.keyboard.press('ControlOrMeta+y');
  await expect(block).toHaveJSProperty('tagName', 'H1');
  await expect(first).toContainText('Alpha');

  const { xml } = await download(page);
  expect(xml).toMatch(/<w:pStyle w:val="Heading1"\/><\/w:pPr><w:r><w:t xml:space="preserve">After the list/);
});

test('undo reverts typing and formatting, and the toolbar buttons work too', async ({ page }) => {
  await openInEditor(page, await buildDocx(LIST));
  const last = page.locator('.docx-page p', { hasText: 'After the list' });
  await last.click();
  await page.keyboard.press('End');
  await page.keyboard.type(' typed');
  await expect(last).toHaveText('After the list typed');

  await page.keyboard.press('Shift+Home');
  await page.getByTitle('Bold (Ctrl+B)').click();
  await expect(last.locator('b, strong')).toHaveText('After the list typed');

  await page.getByTitle('Undo (Ctrl+Z)').click();
  await expect(last.locator('b, strong')).toHaveCount(0);
  await page.getByTitle('Undo (Ctrl+Z)').click();
  await expect(last).toHaveText('After the list');
  await page.getByTitle('Redo (Ctrl+Y)').click();
  await expect(last).toHaveText('After the list typed');
});

test('typed spaces are saved as ordinary spaces', async ({ page }) => {
  await openInEditor(page, await buildDocx('<w:p><w:r><w:t>End</w:t></w:r></w:p>'));
  await page.locator('.docx-page p').click();
  await page.keyboard.press('End');
  await page.keyboard.type('  two  spaces ');
  const { xml } = await download(page);
  expect(xml).toContain('>End  two  spaces </w:t>');
  expect(xml).not.toContain('\u00a0');
});
