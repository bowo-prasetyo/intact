// Superscript, subscript and highlight, as in Word: toolbar buttons and shortcuts, superscript and
// subscript excluding each other, Word's highlight colours, and existing formatting kept or removed.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const DOC = '<w:p><w:r><w:t xml:space="preserve">Area 10 m2 of H2O water</w:t></w:r></w:p>'
  + '<w:p><w:r><w:rPr><w:highlight w:val="green"/></w:rPr><w:t>Green note</w:t></w:r><w:r><w:t xml:space="preserve"> and </w:t></w:r>'
  + '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>sup</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="baseline"/><w:highlight w:val="none"/></w:rPr><w:t xml:space="preserve"> explicit</w:t></w:r></w:p>';
const status = (page) => page.locator('.statusbar [role=status]');
const firstParagraph = (xml) => xml.match(/<w:body><w:p>(.*?)<\/w:p>/)[1];
const secondParagraph = (xml) => xml.match(/<w:body><w:p>.*?<\/w:p><w:p>(.*?)<\/w:p>/)[1];

/** Selects `length` characters of the text node containing `text`, from `offset` in `text`. */
async function select(page, text, offset = 0, length = text.length - offset) {
  await page.evaluate(([t, o, l]) => {
    const root = document.querySelector('.docx-page');
    const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    for (let n = w.nextNode(); n; n = w.nextNode()) {
      const i = n.data.indexOf(t);
      if (i < 0) continue;
      const range = document.createRange();
      range.setStart(n, i + o);
      range.setEnd(n, i + o + l);
      root.focus();
      getSelection().removeAllRanges();
      getSelection().addRange(range);
      return;
    }
    throw new Error(`no text "${t}"`);
  }, [text, offset, length]);
}

test('superscript and subscript from the toolbar, each turning the other off', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  const sup = page.getByTitle('Superscript (Ctrl+Shift+=)');
  const sub = page.getByTitle('Subscript (Ctrl+=)');
  await select(page, 'm2 of', 1, 1);
  await sup.click();
  await expect(sup).toHaveAttribute('aria-pressed', 'true');
  await select(page, 'H2O', 1, 1);
  await sub.click();
  await expect(sub).toHaveAttribute('aria-pressed', 'true');
  let { xml } = await download(page);
  expect(firstParagraph(xml)).toBe('<w:r><w:t xml:space="preserve">Area 10 m</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t xml:space="preserve">2</w:t></w:r>'
    + '<w:r><w:t xml:space="preserve"> of H</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="subscript"/></w:rPr><w:t xml:space="preserve">2</w:t></w:r><w:r><w:t xml:space="preserve">O water</w:t></w:r>');

  // Subscript on the superscript "2" replaces it; pressing it again takes it off.
  await page.locator('.docx-page p').first().locator('sup').evaluate((el) => {
    const range = document.createRange();
    range.selectNodeContents(el);
    getSelection().removeAllRanges();
    getSelection().addRange(range);
  });
  await sub.click();
  await expect(page.locator('.docx-page p').first().locator('sup')).toHaveCount(0);
  await expect(page.locator('.docx-page p').first().locator('sub')).toHaveCount(2);
  await page.keyboard.press('ControlOrMeta+Equal'); // Ctrl+= : subscript off again
  ({ xml } = await download(page));
  expect(firstParagraph(xml)).toBe('<w:r><w:t xml:space="preserve">Area 10 m2 of H</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="subscript"/></w:rPr><w:t xml:space="preserve">2</w:t></w:r><w:r><w:t xml:space="preserve">O water</w:t></w:r>');
});

test('Ctrl+Shift+= makes superscript', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await select(page, 'Area', 0, 4);
  await page.keyboard.press('ControlOrMeta+Shift+Equal');
  const { xml } = await download(page);
  expect(firstParagraph(xml)).toContain('<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t xml:space="preserve">Area</w:t></w:r>');
});

test('highlight in Word\'s colours: the button, the palette, No colour and Ctrl+Alt+H', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  const button = page.getByTitle(/^Highlight: /);
  await expect(button).toHaveAttribute('title', 'Highlight: Yellow (Ctrl+Alt+H)');
  await select(page, 'Area');
  await button.click();
  let { xml } = await download(page);
  expect(firstParagraph(xml)).toContain('<w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t xml:space="preserve">Area</w:t></w:r>');

  // Another colour from the palette becomes the button's colour.
  await select(page, 'water');
  await page.getByTitle('Highlight colour').click();
  await page.getByRole('dialog', { name: 'Highlight colour' }).getByRole('button', { name: 'Turquoise' }).click();
  await expect(button).toHaveAttribute('title', 'Highlight: Turquoise (Ctrl+Alt+H)');
  await expect(page.locator('.docx-page mark[data-color="cyan"]')).toHaveText('water');

  // Ctrl+Alt+H on text already in that colour takes it off; No colour takes any colour off.
  await select(page, 'water');
  await page.keyboard.press('ControlOrMeta+Alt+KeyH');
  await expect(page.locator('.docx-page mark[data-color="cyan"]')).toHaveCount(0);
  await select(page, 'Green note');
  await page.getByTitle('Highlight colour').click();
  await page.getByRole('dialog', { name: 'Highlight colour' }).getByRole('button', { name: 'No colour' }).click();
  ({ xml } = await download(page));
  expect(firstParagraph(xml)).not.toContain('cyan');
  expect(secondParagraph(xml)).toMatch(/^<w:r><w:t xml:space="preserve">Green note<\/w:t><\/w:r>/); // highlight gone

  // Nothing selected: highlighting needs a selection, as Word's highlighter would.
  await select(page, 'Area', 0, 0);
  await button.click();
  await expect(status(page)).toHaveText('Select the text to highlight first');
});

test('existing formatting shows and stays: superscript as superscript, an explicit "none" kept', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await expect(page.locator('.docx-page sup')).toHaveText('sup');
  await expect(page.locator('.docx-page mark[data-color="green"]')).toHaveText('Green note');
  await select(page, 'Area', 0, 1);
  await page.keyboard.type('a'); // an edit elsewhere
  const { xml } = await download(page);
  expect(secondParagraph(xml)).toContain('<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t xml:space="preserve">sup</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="baseline"/><w:highlight w:val="none"/></w:rPr><w:t xml:space="preserve"> explicit</w:t></w:r>');
});

test('Ctrl+Space takes superscript off but keeps highlight, as in Word', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await page.evaluate(() => {
    const p = document.querySelectorAll('.docx-page p')[1];
    const range = document.createRange();
    range.selectNodeContents(p);
    document.querySelector('.docx-page').focus();
    getSelection().removeAllRanges();
    getSelection().addRange(range);
  });
  await page.keyboard.press('ControlOrMeta+Space');
  await expect(page.locator('.docx-page sup')).toHaveCount(0);
  await expect(page.locator('.docx-page mark')).toHaveText('Green note');
});

test('pasted superscript and subscript are kept', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await select(page, 'Area', 0, 0);
  await page.evaluate(() => {
    const event = new ClipboardEvent('paste', { bubbles: true, cancelable: true });
    const dt = new DataTransfer();
    dt.setData('text/html', '<p>E=mc<sup>2</sup>, CO<span style="vertical-align: sub">2</span> </p>');
    dt.setData('text/plain', 'E=mc2, CO2 ');
    Object.defineProperty(event, 'clipboardData', { value: dt });
    document.querySelector('.docx-page p').dispatchEvent(event);
  });
  const { xml } = await download(page);
  expect(firstParagraph(xml)).toMatch(/^<w:r><w:t xml:space="preserve">E=mc<\/w:t><\/w:r><w:r><w:rPr><w:vertAlign w:val="superscript"\/><\/w:rPr><w:t xml:space="preserve">2<\/w:t><\/w:r><w:r><w:t xml:space="preserve">, CO<\/w:t><\/w:r><w:r><w:rPr><w:vertAlign w:val="subscript"\/><\/w:rPr><w:t xml:space="preserve">2<\/w:t><\/w:r>/);
});
