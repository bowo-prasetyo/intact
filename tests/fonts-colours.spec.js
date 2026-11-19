// Text formatting from what the document already uses: its theme and own fonts, its sizes (and
// Word's usual ones), its theme and own text colours, and its character styles. Clear formatting
// (Ctrl+Space) takes direct formatting off, as Word does; highlight stays. Each change is written
// as Word writes it, on the selected text only.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const THEME = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Studio"><a:themeElements><a:clrScheme name="Studio"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F2A44"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="0E6467"/></a:accent1><a:accent2><a:srgbClr val="C55A11"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Studio"><a:majorFont><a:latin typeface="Georgia"/></a:majorFont><a:minorFont><a:latin typeface="Verdana"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>';
const BODY = '<w:p><w:r><w:t xml:space="preserve">Plain words to style here.</w:t></w:r></w:p>'
  + '<w:p><w:r><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/><w:color w:val="7030A0"/><w:sz w:val="28"/></w:rPr><w:t>Brand purple</w:t></w:r>'
  + '<w:r><w:t xml:space="preserve"> and </w:t></w:r><w:r><w:rPr><w:rStyle w:val="Strong"/></w:rPr><w:t>strong</w:t></w:r></w:p>'
  + '<w:p><w:r><w:rPr><w:b/><w:i/><w:color w:val="FF0000"/><w:sz w:val="32"/><w:highlight w:val="yellow"/></w:rPr><w:t>Loud text</w:t></w:r></w:p>';
const OPTS = { parts: { 'word/theme/theme1.xml': THEME }, rels: `<Relationship Id="rIdT" Type="${REL}/theme" Target="theme/theme1.xml"/>` };

const status = (page) => page.locator('.statusbar [role=status]');
const para = (page, n) => page.locator('.docx-page > p').nth(n);
const paragraphXml = (xml, n) => [...xml.matchAll(/<w:p>(.*?)<\/w:p>/g)][n][1];

/** Selects `word` in paragraph `n`. */
async function select(page, n, word) {
  await page.evaluate(([i, w]) => {
    const p = document.querySelectorAll('.docx-page > p')[i];
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    for (let t = walker.nextNode(); t; t = walker.nextNode()) {
      const at = t.data.indexOf(w);
      if (at < 0) continue;
      document.querySelector('.docx-page').focus();
      getSelection().setBaseAndExtent(t, at, t, at + w.length);
      return;
    }
  }, [n, word]);
}

test('the toolbar offers the document’s own fonts, sizes and character styles', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY, OPTS));
  const font = page.getByRole('combobox', { name: 'Font', exact: true });
  await expect(font.locator('optgroup[label="Theme fonts"] option')).toHaveText(['Georgia (Headings)', 'Verdana (Body)']);
  await expect(font.locator('optgroup[label="In this document"] option')).toHaveText(['Courier New']);
  // Then Word's usual fonts, without those already listed.
  const others = await font.locator('optgroup[label="Other fonts"] option').allTextContents();
  expect(others).toEqual(expect.arrayContaining(['Arial', 'Calibri', 'Cambria', 'Times New Roman']));
  expect(others).not.toContain('Courier New');
  expect(others).not.toContain('Georgia');
  const size = page.getByRole('combobox', { name: 'Font size' });
  await expect(size.locator('optgroup[label="In this document"] option')).toHaveText(['11', '14', '16']);
  await expect(size.locator('optgroup[label="Other sizes"] option').first()).toHaveText('8');
  const style = page.getByRole('combobox', { name: 'Character style' });
  await expect(style.locator('option')).toHaveText(['No character style', 'Strong']);
  // The caret's font and size show.
  await select(page, 1, 'Brand');
  await expect(font).toHaveValue('Courier New');
  await expect(size).toHaveValue('14');
});

test('a blank document offers Word’s usual fonts, written by name', async ({ page }) => {
  await page.goto('./');
  await page.getByText('Blank document').click();
  const font = page.getByRole('combobox', { name: 'Font', exact: true });
  await page.locator('.docx-page').click();
  await page.keyboard.type('Dear Ms Smith');
  await page.keyboard.press('Shift+Home');
  await font.selectOption('Times New Roman');
  await expect(font).toHaveValue('Times New Roman');
  const { xml } = await download(page);
  expect(xml).toContain('<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>');
});

test('a font, a theme font and a size are written as Word writes them, on the selected text only', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY, OPTS));
  await select(page, 0, 'words');
  await page.getByRole('combobox', { name: 'Font', exact: true }).selectOption('Courier New');
  await select(page, 0, 'style');
  await page.getByRole('combobox', { name: 'Font', exact: true }).selectOption('theme:major');
  await page.getByRole('combobox', { name: 'Font size' }).selectOption('16');
  expect(await para(page, 0).locator('span', { hasText: 'words' }).evaluate((el) => getComputedStyle(el).fontFamily)).toMatch(/^"?Courier New/);
  const { xml } = await download(page);
  const p = paragraphXml(xml, 0);
  expect(p).toContain('<w:r><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/></w:rPr><w:t xml:space="preserve">words</w:t></w:r>');
  expect(p).toContain('<w:r><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr><w:t xml:space="preserve">style</w:t></w:r>');
  expect(p).toMatch(/^<w:r><w:t xml:space="preserve">Plain <\/w:t><\/w:r>/);
});

test('a colour from the theme or the document, or Automatic; the button applies the last one', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY, OPTS));
  await select(page, 0, 'Plain');
  await page.getByTitle('Font colour: the document’s theme and its own colours').click();
  const palette = page.getByRole('dialog', { name: 'Font colour' });
  await expect(palette.locator('.swatches').first().locator('.swatch')).toHaveCount(10);
  await expect(palette.locator('.swatches').nth(1).locator('.swatch')).toHaveCount(3); // 7030A0 and FF0000 in the text, 808080 in a style
  await palette.getByRole('button', { name: 'Accent 1' }).click();
  await select(page, 0, 'here');
  await page.getByRole('button', { name: /^Font colour: Accent 1$/ }).click(); // the last colour chosen
  await select(page, 1, 'Brand');
  await page.getByTitle('Font colour: the document’s theme and its own colours').click();
  await palette.getByRole('button', { name: 'Automatic' }).click();
  expect(await para(page, 0).locator('span', { hasText: 'Plain' }).evaluate((el) => getComputedStyle(el).color)).toBe('rgb(14, 100, 103)');
  const { xml } = await download(page);
  expect(paragraphXml(xml, 0)).toContain('<w:r><w:rPr><w:color w:val="0E6467" w:themeColor="accent1"/></w:rPr><w:t xml:space="preserve">Plain</w:t></w:r>');
  expect(paragraphXml(xml, 0)).toContain('<w:r><w:rPr><w:color w:val="0E6467" w:themeColor="accent1"/></w:rPr><w:t xml:space="preserve">here</w:t></w:r>');
  expect(paragraphXml(xml, 1)).toContain('<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/><w:color w:val="auto"/><w:sz w:val="28"/></w:rPr><w:t xml:space="preserve">Brand</w:t>');
});

test('a character style is applied, and taken off again', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY, OPTS));
  await select(page, 0, 'words');
  const style = page.getByRole('combobox', { name: 'Character style' });
  await style.selectOption('Strong');
  await select(page, 1, 'strong');
  await expect(style).toHaveValue('Strong');
  await style.selectOption('');
  const { xml } = await download(page);
  expect(paragraphXml(xml, 0)).toContain('<w:r><w:rPr><w:rStyle w:val="Strong"/></w:rPr><w:t xml:space="preserve">words</w:t></w:r>');
  expect(paragraphXml(xml, 1)).not.toContain('Strong');
});

test('Clear formatting takes bold, italic, font, size and colour off; highlight stays; Undo puts them back', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY, OPTS));
  await select(page, 2, 'Loud');
  await page.keyboard.press('ControlOrMeta+Space');
  await expect(para(page, 2).locator('strong')).toHaveText('text'); // only the selection changed
  await expect(para(page, 2).locator('mark')).toHaveText(['Loud', ' text']);
  let { xml } = await download(page);
  expect(paragraphXml(xml, 2)).toMatch(/^<w:r><w:rPr><w:highlight w:val="yellow"\/><\/w:rPr><w:t xml:space="preserve">Loud<\/w:t><\/w:r>/);
  await select(page, 2, ' text');
  await page.getByRole('button', { name: /^Clear formatting/ }).click();
  ({ xml } = await download(page));
  expect(paragraphXml(xml, 2)).toBe('<w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t xml:space="preserve">Loud text</w:t></w:r>');
  await page.keyboard.press('ControlOrMeta+z');
  await page.keyboard.press('ControlOrMeta+z');
  await expect(para(page, 2).locator('strong')).toHaveText('Loud text');
});

test('a font, size or colour needs text selected; the status bar says so', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY, OPTS));
  await para(page, 0).click();
  await page.getByRole('combobox', { name: 'Font size' }).selectOption('16');
  await expect(status(page)).toHaveText('Select the text first');
});
