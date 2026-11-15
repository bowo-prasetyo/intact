// Word's keyboard shortcuts for what Intact can do: alignment, styles, lists, list levels, clearing
// character formatting, page breaks and non-breaking spaces.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const NUM = '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>';
const DOC = '<w:p><w:r><w:t xml:space="preserve">Alpha paragraph</w:t></w:r></w:p>'
  + '<w:p><w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>Bold italic</w:t></w:r></w:p>'
  + `<w:p>${NUM}<w:r><w:t>List item</w:t></w:r></w:p>`;
const status = (page) => page.locator('.statusbar [role=status]');

async function caretIn(page, text, offset = 0) {
  await page.evaluate(([t, o]) => {
    const root = document.querySelector('.docx-page');
    const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    for (let n = w.nextNode(); n; n = w.nextNode()) {
      if (!n.data.includes(t)) continue;
      const range = document.createRange();
      range.setStart(n, n.data.indexOf(t) + o);
      root.focus();
      getSelection().removeAllRanges();
      getSelection().addRange(range);
      return;
    }
    throw new Error(`no text "${t}"`);
  }, [text, offset]);
}

const firstParagraph = (xml) => xml.match(/<w:body><w:p>(.*?)<\/w:p>/)[1];

for (const [keys, jc] of [['ControlOrMeta+e', 'center'], ['ControlOrMeta+r', 'right'], ['ControlOrMeta+j', 'both']]) {
  test(`${keys.replace('ControlOrMeta', 'Ctrl')} aligns the paragraph (${jc})`, async ({ page }) => {
    await openInEditor(page, await buildDocx(DOC));
    await caretIn(page, 'Alpha');
    await page.keyboard.press(keys);
    let { xml } = await download(page);
    expect(firstParagraph(xml)).toContain(`<w:pPr><w:jc w:val="${jc}"/></w:pPr>`);
    await caretIn(page, 'Alpha');
    await page.keyboard.press('ControlOrMeta+l'); // back to left, as the Align left button does
    ({ xml } = await download(page));
    expect(firstParagraph(xml)).toContain('<w:pPr><w:jc w:val="left"/></w:pPr>');
  });
}

test('Ctrl+Alt+1, 2, 3 apply headings and Ctrl+Shift+N goes back to Normal', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  const style = page.getByLabel('Paragraph style');
  await caretIn(page, 'Alpha');
  for (const [n, value] of [['1', 'h1'], ['2', 'h2'], ['3', 'h3']]) {
    await page.keyboard.press(`ControlOrMeta+Alt+Digit${n}`);
    await expect(style).toHaveValue(value);
  }
  let { xml } = await download(page);
  expect(firstParagraph(xml)).toContain('<w:pStyle w:val="Heading3"/>');
  await caretIn(page, 'Alpha');
  await page.keyboard.press('ControlOrMeta+Shift+N');
  await expect(style).toHaveValue('normal');
  ({ xml } = await download(page));
  expect(firstParagraph(xml)).not.toContain('w:pStyle');
});

test('Ctrl+Shift+L makes a bulleted list; Ctrl+M and Ctrl+Shift+M change list levels', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await caretIn(page, 'Alpha');
  await page.keyboard.press('ControlOrMeta+Shift+L');
  await expect(page.getByTitle('Bullets (Ctrl+Shift+L)')).toHaveAttribute('aria-pressed', 'true');
  await caretIn(page, 'List item');
  await page.keyboard.press('ControlOrMeta+m');
  let { xml } = await download(page);
  expect(xml).toMatch(/<w:numPr><w:ilvl w:val="1"\/><w:numId w:val="1"\/><\/w:numPr><\/w:pPr><w:r><w:t xml:space="preserve">List item/);
  await caretIn(page, 'List item');
  await page.keyboard.press('ControlOrMeta+Shift+M');
  ({ xml } = await download(page));
  expect(xml).toMatch(/<w:numPr><w:ilvl w:val="0"\/><w:numId w:val="1"\/><\/w:numPr><\/w:pPr><w:r><w:t xml:space="preserve">List item/);

  await caretIn(page, 'Bold italic');
  await page.keyboard.press('ControlOrMeta+m');
  await expect(status(page)).toHaveText('Indenting works on list items here');
});

test('Ctrl+Space takes bold and italic off the selection', async ({ page }) => {
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
  const { xml } = await download(page);
  expect(xml).toContain('<w:p><w:r><w:t xml:space="preserve">Bold italic</w:t></w:r></w:p>');
});

test('Ctrl+Enter inserts a page break, tracked while tracking', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await caretIn(page, 'Alpha paragraph', 5);
  await page.keyboard.press('ControlOrMeta+Enter');
  await expect(page.locator('.docx-page .page-break')).toHaveCount(1);
  let { xml } = await download(page);
  expect(firstParagraph(xml)).toBe('<w:r><w:t xml:space="preserve">Alpha</w:t></w:r><w:r><w:br w:type="page"/></w:r><w:r><w:t xml:space="preserve"> paragraph</w:t></w:r>');

  await page.keyboard.press('ControlOrMeta+z');
  await page.getByRole('button', { name: 'Track changes' }).click();
  await page.getByRole('dialog', { name: 'Your name on tracked changes' }).getByRole('textbox').fill('Rita');
  await page.getByRole('button', { name: 'Start tracking' }).click();
  await caretIn(page, 'Alpha paragraph', 5);
  await page.keyboard.press('ControlOrMeta+Enter');
  ({ xml } = await download(page));
  expect(firstParagraph(xml)).toMatch(/<w:ins w:id="\d+" w:author="Rita" w:date="[^"]+"><w:r><w:br w:type="page"\/><\/w:r><\/w:ins>/);
});

test('Ctrl+Shift+Space types a non-breaking space', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await caretIn(page, 'Alpha paragraph', 5);
  await page.keyboard.press('ControlOrMeta+Shift+Space');
  const { xml } = await download(page);
  expect(firstParagraph(xml)).toContain('Alpha  paragraph');
});

test('toolbar tooltips name the shortcuts', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  for (const title of ['Center (Ctrl+E)', 'Align left (Ctrl+L)', 'Align right (Ctrl+R)', 'Justify (Ctrl+J)', 'Bullets (Ctrl+Shift+L)', 'Track Changes (Ctrl+Shift+E)']) {
    await expect(page.getByTitle(title, { exact: true })).toHaveCount(1);
  }
});
