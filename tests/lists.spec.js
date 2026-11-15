// Lists and tables respond to Enter, Backspace and Tab the way they do in Word, and the Bullets and
// Numbering buttons start and end lists. Each change is written to the saved file.
import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import fs from 'node:fs';
import { buildDocx, openInEditor, download } from './helpers.js';

const NUM = '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>';
const LIST = `<w:p>${NUM}<w:r><w:t>Alpha</w:t></w:r></w:p><w:p>${NUM}<w:r><w:t>Beta</w:t></w:r></w:p><w:p><w:r><w:t>After the list</w:t></w:r></w:p>`;
const TABLE = `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>
  <w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr>
  <w:tr><w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc></w:tr>
</w:tbl><w:p><w:r><w:t>Below</w:t></w:r></w:p>`;

const markers = (page) => page.locator('.docx-page .list').evaluateAll((els) => els.map((e) => `${e.dataset.marker} ${e.textContent}`));
const para = (page, text) => page.locator('.docx-page :is(p, h1, h2, h3)', { hasText: text });
/** The saved <w:p> whose text contains `text`. */
const savedParagraph = (xml, text) => xml.match(/<w:p[ >].*?<\/w:p>/gs).find((p) => p.includes(`>${text}<`));

async function caretAt(page, text, where = 'End') {
  await para(page, text).click();
  await page.keyboard.press(where);
}

test('Enter on an empty list item ends the list', async ({ page }) => {
  await openInEditor(page, await buildDocx(LIST));
  await caretAt(page, 'Beta');
  await page.keyboard.press('Enter');
  await expect.poll(() => markers(page)).toEqual(['1. Alpha', '2. Beta', '3. ']);
  await page.keyboard.press('Enter');
  await page.keyboard.type('Plain text');
  await expect.poll(() => markers(page)).toEqual(['1. Alpha', '2. Beta']);
  await expect(page.getByTitle('Numbering')).not.toHaveClass(/on/);
  await expect(para(page, 'Plain text')).toHaveCSS('padding-left', '0px');

  const { xml } = await download(page);
  expect(savedParagraph(xml, 'Plain text')).not.toContain('numPr');
  expect(savedParagraph(xml, 'Beta')).toContain('<w:numId w:val="1"/>');
});

test('Enter on an empty sub-item moves it up a level before ending the list', async ({ page }) => {
  const SUB = '<w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr>';
  await openInEditor(page, await buildDocx(`<w:p>${NUM}<w:r><w:t>Alpha</w:t></w:r></w:p><w:p>${SUB}<w:r><w:t>Sub</w:t></w:r></w:p>`));
  await caretAt(page, 'Sub');
  await page.keyboard.press('Enter');
  await page.keyboard.press('Enter');
  await page.keyboard.type('Next');
  await expect.poll(() => markers(page)).toEqual(['1. Alpha', 'a. Sub', '2. Next']);
});

test('Tab at the start of a list item changes its level, Shift+Tab changes it back', async ({ page }) => {
  await openInEditor(page, await buildDocx(LIST));
  await caretAt(page, 'Beta', 'Home');
  await page.keyboard.press('Tab');
  await expect.poll(() => markers(page)).toEqual(['1. Alpha', 'a. Beta']);
  await expect(para(page, 'Beta')).toHaveText('Beta'); // no tab character typed

  let { xml } = await download(page);
  expect(savedParagraph(xml, 'Beta')).toContain('<w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr>');
  expect(xml).not.toContain('<w:tab/>');

  await caretAt(page, 'Beta', 'Home'); // Download took the focus
  await page.keyboard.press('Shift+Tab');
  await expect.poll(() => markers(page)).toEqual(['1. Alpha', '2. Beta']);
  await expect(para(page, 'Beta')).toHaveCSS('padding-left', '48px'); // 36pt: back to level 1's indent
  ({ xml } = await download(page));
  expect(savedParagraph(xml, 'Beta')).toContain('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>');

  await page.getByTitle('Undo (Ctrl+Z)').click();
  await expect.poll(() => markers(page)).toEqual(['1. Alpha', 'a. Beta']);
  await expect(para(page, 'Beta')).toHaveCSS('padding-left', '96px'); // 72pt: level 2's indent
});

test('Tab inside the text of a list item types a tab, as in Word', async ({ page }) => {
  await openInEditor(page, await buildDocx(LIST));
  await caretAt(page, 'Beta');
  await page.keyboard.press('Tab');
  await expect(para(page, 'Beta')).toHaveText('Beta\t');
  await expect.poll(() => markers(page)).toEqual(['1. Alpha', '2. Beta\t']);
});

test('the level buttons work on selected list items', async ({ page }) => {
  await openInEditor(page, await buildDocx(LIST));
  await caretAt(page, 'After the list');
  await expect(page.getByTitle('Increase list level (Tab)')).toBeDisabled();
  await caretAt(page, 'Beta');
  await page.getByTitle('Increase list level (Tab)').click();
  await expect.poll(() => markers(page)).toEqual(['1. Alpha', 'a. Beta']);
  await page.getByTitle('Decrease list level (Shift+Tab)').click();
  await expect.poll(() => markers(page)).toEqual(['1. Alpha', '2. Beta']);
});

test('a list with a single level is not given levels it does not define', async ({ page }) => {
  // Word's built-in "List Bullet" style uses a one-level list: Word and LibreOffice would each
  // render a second level differently, so Intact doesn't create one.
  const STYLE = '<w:pPr><w:pStyle w:val="ListBullet"/></w:pPr>';
  await openInEditor(page, await buildDocx(`<w:p>${STYLE}<w:r><w:t>One</w:t></w:r></w:p><w:p>${STYLE}<w:r><w:t>Two</w:t></w:r></w:p>`));
  await caretAt(page, 'Two', 'Home');
  await page.keyboard.press('Tab');
  await expect(page.locator('.statusbar')).toContainText('This list has no further levels');
  await expect.poll(() => markers(page)).toEqual(['• One', '• Two']);
  await expect(para(page, 'Two')).toHaveText('Two');
});

test('Backspace at the start of a list item removes its number but keeps the text', async ({ page }) => {
  await openInEditor(page, await buildDocx(LIST));
  await caretAt(page, 'Beta', 'Home');
  await page.keyboard.press('Backspace');
  await expect.poll(() => markers(page)).toEqual(['1. Alpha']);
  await expect(para(page, 'Beta')).toHaveText('Beta');
  const { xml } = await download(page);
  expect(savedParagraph(xml, 'Beta')).not.toContain('numPr');
});

test('ending a list made by a list style returns the paragraph to Normal', async ({ page }) => {
  const STYLE = '<w:pPr><w:pStyle w:val="ListBullet"/></w:pPr>';
  await openInEditor(page, await buildDocx(`<w:p>${STYLE}<w:r><w:t>Styled bullet</w:t></w:r></w:p>`));
  await caretAt(page, 'Styled bullet');
  await page.keyboard.press('Enter');
  await expect.poll(() => markers(page)).toEqual(['• Styled bullet', '• ']);
  await page.keyboard.press('Enter');
  await page.keyboard.type('Normal again');
  await expect.poll(() => markers(page)).toEqual(['• Styled bullet']);
  await expect(page.getByLabel('Paragraph style')).toHaveValue('normal');
  const { xml } = await download(page);
  expect(savedParagraph(xml, 'Normal again')).not.toContain('<w:pPr>');
  expect(savedParagraph(xml, 'Styled bullet')).toContain('<w:pStyle w:val="ListBullet"/>');
});

test('Bullets and Numbering start and end lists, adding numbering to a document without any', async ({ page }) => {
  page.on('dialog', (d) => d.accept());
  await page.goto('./');
  await page.getByRole('button', { name: /Blank document/ }).click();
  const doc = page.locator('.docx-page');
  await doc.click();
  await page.keyboard.type('One');
  await page.keyboard.press('Enter');
  await page.keyboard.type('Two');
  await page.keyboard.press('ControlOrMeta+a');
  await page.getByTitle('Numbering').click();
  await expect.poll(() => markers(page)).toEqual(['1. One', '2. Two']);
  await expect(page.getByTitle('Numbering')).toHaveClass(/on/);

  await page.getByTitle('Bullets').click();
  await expect.poll(() => markers(page)).toEqual(['• One', '• Two']);

  const [dl] = await Promise.all([page.waitForEvent('download'), page.getByRole('button', { name: 'Download' }).click()]);
  const zip = await JSZip.loadAsync(fs.readFileSync(await dl.path()));
  const xml = await zip.file('word/document.xml').async('string');
  const numbering = await zip.file('word/numbering.xml')?.async('string');
  expect(numbering).toContain('<w:numFmt w:val="bullet"/>');
  expect(await zip.file('word/_rels/document.xml.rels').async('string')).toContain('relationships/numbering" Target="numbering.xml"');
  expect(await zip.file('[Content_Types].xml').async('string')).toContain('PartName="/word/numbering.xml"');
  const numId = /<w:numId w:val="(\d+)"\/>/.exec(savedParagraph(xml, 'One'))[1];
  expect(numbering).toContain(`<w:num w:numId="${numId}">`);
  expect(numbering.indexOf('<w:abstractNum ')).toBeLessThan(numbering.indexOf('<w:num '));

  await page.getByTitle('Bullets').click();
  await expect.poll(() => markers(page)).toEqual([]);
});

test('Numbering next to a numbered list continues it; a separate one starts at 1', async ({ page }) => {
  await openInEditor(page, await buildDocx(`${LIST}<w:p><w:r><w:t>Gap</w:t></w:r></w:p><w:p><w:r><w:t>Far away</w:t></w:r></w:p>`));
  await caretAt(page, 'After the list');
  await page.getByTitle('Numbering').click();
  await caretAt(page, 'Far away');
  await page.getByTitle('Numbering').click();
  await expect.poll(() => markers(page)).toEqual(['1. Alpha', '2. Beta', '3. After the list', '1. Far away']);
  const { xml } = await download(page);
  expect(savedParagraph(xml, 'After the list')).toContain('<w:numId w:val="1"/>');
  expect(savedParagraph(xml, 'Far away')).not.toContain('<w:numId w:val="1"/>');
});

test('Tab and Shift+Tab move between table cells instead of typing a tab', async ({ page }) => {
  await openInEditor(page, await buildDocx(TABLE));
  await page.locator('td', { hasText: 'A1' }).click();
  await page.keyboard.press('Tab');
  await page.keyboard.type('X'); // the next cell's text is selected, as in Word
  await page.keyboard.press('Tab');
  await page.keyboard.press('Shift+Tab');
  await page.keyboard.press('Shift+Tab');
  await page.keyboard.press('End');
  await page.keyboard.type('!');
  await expect(page.locator('td')).toHaveText(['A1!', 'X', 'A2', 'B2']);

  await page.locator('td', { hasText: 'B2' }).click();
  await page.keyboard.press('Tab');
  await expect(page.locator('.statusbar')).toContainText('Adding table rows isn’t supported yet');
  const { xml } = await download(page);
  expect(xml).not.toContain('<w:tab/>');
});
