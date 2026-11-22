// New style… (the Paragraph style menu), as Word's Create New Style from Formatting: a paragraph or
// character style made from the selection's formatting (w:customStyle, in Word's gallery), given to
// the selection, whose own formatting that the style now gives is taken off it, so the style rules
// it (a later Modify style… changes it too). Ctrl+Z takes the style away again.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const MONO = '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/><w:sz w:val="20"/>';
const CODE = `<w:p><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/><w:spacing w:after="0"/><w:ind w:left="284"/></w:pPr><w:r><w:rPr>${MONO}</w:rPr><w:t>pumpctl --start</w:t></w:r></w:p>`;
const P = (t) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
const DOC = P('Start the pump:') + CODE + `<w:p><w:r><w:t xml:space="preserve">Then run </w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/><w:color w:val="C7254E"/></w:rPr><w:t>pumpctl</w:t></w:r><w:r><w:t xml:space="preserve"> again.</w:t></w:r></w:p>`;
const styleMenu = (page) => page.getByRole('combobox', { name: 'Paragraph style' });
const style = (xml, id) => new RegExp(`<w:style [^>]*w:styleId="${id}".*?</w:style>`, 's').exec(xml)?.[0];
const savedParagraph = (xml, text) => xml.match(/<w:p[ >].*?<\/w:p>/gs).find((p) => p.includes(text));

async function newStyle(page, name, type) {
  await styleMenu(page).selectOption({ label: 'New style…' });
  const dialog = page.getByRole('dialog', { name: 'Create a new style' });
  await dialog.getByLabel('Style name').fill(name);
  await dialog.getByLabel('Style type').selectOption(type);
  await dialog.getByRole('button', { name: 'OK' }).click();
}

test('a paragraph style from a code sample\'s formatting, given to it, and changed later with Modify style', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await page.locator('.docx-page p', { hasText: 'pumpctl --start' }).click();
  await newStyle(page, 'Code', 'paragraph');
  await expect(styleMenu(page)).toHaveValue('style:Code');

  let { zip, xml } = await download(page);
  const styles = await zip.file('word/styles.xml').async('string');
  const code = style(styles, 'Code');
  expect(code).toContain('w:type="paragraph" w:customStyle="1" w:styleId="Code"><w:name w:val="Code"/><w:basedOn w:val="Normal"/><w:next w:val="Code"/><w:qFormat/>');
  expect(code).toContain('<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>');
  expect(code).toContain('<w:ind w:left="284"/>');
  expect(code).toContain('<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/>');
  // The paragraph is in the style, without the formatting the style now gives it.
  expect(savedParagraph(xml, 'pumpctl --start')).toBe('<w:p><w:pPr><w:pStyle w:val="Code"/></w:pPr><w:r><w:t xml:space="preserve">pumpctl --start</w:t></w:r></w:p>');

  // The style is offered for other paragraphs.
  await page.locator('.docx-page p', { hasText: 'Start the pump' }).click();
  await styleMenu(page).selectOption({ label: 'Code' });
  ({ xml } = await download(page));
  expect(savedParagraph(xml, 'Start the pump')).toContain('<w:pStyle w:val="Code"/>');
  await expect(page.locator('.docx-page p', { hasText: 'Start the pump' })).toHaveCSS('font-family', /Consolas/);
});

test('a character style from the selected text, and Ctrl+Z takes it away', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await page.evaluate(() => {
    const p = [...document.querySelectorAll('.docx-page p')].find((e) => e.textContent.startsWith('Then run'));
    const node = [...p.querySelectorAll('*')].flatMap((e) => [...e.childNodes]).find((n) => n.nodeType === 3 && n.data === 'pumpctl');
    document.querySelector('.docx-page').focus();
    getSelection().setBaseAndExtent(node, 0, node, node.length);
  });
  await newStyle(page, 'Inline code', 'character');
  await expect(page.getByRole('combobox', { name: 'Character style' })).toHaveValue('Inlinecode');
  let { zip, xml } = await download(page);
  const styles = await zip.file('word/styles.xml').async('string');
  expect(style(styles, 'Inlinecode')).toContain('w:type="character" w:customStyle="1" w:styleId="Inlinecode"><w:name w:val="Inline code"/><w:qFormat/><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/><w:color w:val="C7254E"/></w:rPr>');
  expect(savedParagraph(xml, 'Then run')).toContain('<w:r><w:rPr><w:rStyle w:val="Inlinecode"/></w:rPr><w:t xml:space="preserve">pumpctl</w:t></w:r>');

  // A name already used is refused, saying why.
  await styleMenu(page).selectOption({ label: 'New style…' });
  const dialog = page.getByRole('dialog', { name: 'Create a new style' });
  await dialog.getByLabel('Style name').fill('normal');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(dialog.getByRole('alert')).toContainText('has a style called “normal” already');
  await dialog.getByRole('button', { name: 'Cancel' }).click();

  await page.locator('.docx-page p', { hasText: 'Then run' }).click();
  await page.keyboard.press('ControlOrMeta+z');
  ({ zip, xml } = await download(page));
  expect(await zip.file('word/styles.xml').async('string')).not.toContain('Inlinecode');
  expect(savedParagraph(xml, 'Then run')).toContain('<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/>');
});
