// Line and paragraph spacing and indentation, as Word's Line and Paragraph Spacing menu, Paragraph
// dialog and shortcuts set them. Only what is changed is written, in the paragraph's own properties
// (<w:spacing>, <w:ind>), with everything else kept; tracked while tracking; undone with Ctrl+Z.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

// "Kept" carries properties Intact doesn't set (keep with next, a tab stop, Word's "lines" spacing).
const DOC = '<w:p><w:r><w:t>Alpha</w:t></w:r></w:p>'
  + '<w:p><w:pPr><w:keepNext/><w:tabs><w:tab w:val="left" w:pos="2000"/></w:tabs><w:spacing w:beforeLines="100" w:before="240" w:after="120"/><w:jc w:val="center"/></w:pPr><w:r><w:t>Kept</w:t></w:r></w:p>'
  + '<w:p><w:r><w:t>Gamma</w:t></w:r></w:p>';
const savedParagraph = (xml, text) => xml.match(/<w:p[ >].*?<\/w:p>/gs).find((p) => p.includes(`>${text}<`));
const run = (text) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

async function caretIn(page, text) {
  await page.locator('.docx-page p', { hasText: text }).click();
  await page.keyboard.press('End');
}

async function selectParagraphs(page, from, to) {
  await page.evaluate(([a, b]) => {
    const ps = [...document.querySelectorAll('.docx-page p')];
    const range = document.createRange();
    range.setStart(ps.find((p) => p.textContent === a).firstChild, 0);
    const last = ps.find((p) => p.textContent === b).firstChild;
    range.setEnd(last, last.length);
    document.querySelector('.docx-page').focus();
    getSelection().removeAllRanges();
    getSelection().addRange(range);
  }, [from, to]);
}

const menu = (page) => page.getByRole('menu', { name: 'Line and paragraph spacing' });
const openMenu = (page) => page.getByRole('button', { name: 'Line and paragraph spacing' }).click();

test('the menu sets line spacing on the selected paragraphs, checks the current one, and Ctrl+Z undoes it', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await selectParagraphs(page, 'Alpha', 'Kept');
  await openMenu(page);
  await expect(menu(page).getByRole('menuitemradio', { name: '1.0' })).toHaveAttribute('aria-checked', 'true');
  await menu(page).getByRole('menuitemradio', { name: '1.5' }).click();
  await expect(menu(page)).toBeHidden();
  await expect(page.locator('.docx-page p').first()).toHaveCSS('line-height', /px$/);
  let { xml } = await download(page);
  expect(savedParagraph(xml, 'Alpha')).toBe(`<w:p><w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr>${run('Alpha')}</w:p>`);
  // Everything else the paragraph had is kept, in Word's order.
  expect(savedParagraph(xml, 'Kept')).toBe('<w:p><w:pPr><w:keepNext/><w:tabs><w:tab w:val="left" w:pos="2000"/></w:tabs>'
    + `<w:spacing w:beforeLines="100" w:before="240" w:after="120" w:line="360" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>${run('Kept')}</w:p>`);
  expect(savedParagraph(xml, 'Gamma')).toBe(`<w:p>${run('Gamma')}</w:p>`);

  await caretIn(page, 'Alpha');
  await openMenu(page);
  await expect(menu(page).getByRole('menuitemradio', { name: '1.5' })).toHaveAttribute('aria-checked', 'true');
  await page.keyboard.press('Escape');
  await page.keyboard.press('ControlOrMeta+z');
  ({ xml } = await download(page));
  expect(savedParagraph(xml, 'Alpha')).toBe(`<w:p>${run('Alpha')}</w:p>`);
});

test('Ctrl+2, Ctrl+5 and Ctrl+1 set double, 1.5 and single spacing; Ctrl+0 adds and removes space before', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await caretIn(page, 'Alpha');
  await page.keyboard.press('ControlOrMeta+2');
  let { xml } = await download(page);
  expect(savedParagraph(xml, 'Alpha')).toContain('<w:spacing w:line="480" w:lineRule="auto"/>');
  await caretIn(page, 'Alpha');
  await page.keyboard.press('ControlOrMeta+5');
  await page.keyboard.press('ControlOrMeta+0');
  ({ xml } = await download(page));
  expect(savedParagraph(xml, 'Alpha')).toContain('<w:spacing w:line="360" w:lineRule="auto" w:before="240"/>');
  await caretIn(page, 'Alpha');
  await page.keyboard.press('ControlOrMeta+1');
  await page.keyboard.press('ControlOrMeta+0');
  ({ xml } = await download(page));
  expect(savedParagraph(xml, 'Alpha')).toContain('<w:spacing w:line="240" w:lineRule="auto" w:before="0"/>');
});

test('"Remove space after paragraph" and "Add space before paragraph" follow what the paragraph has', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await caretIn(page, 'Kept');
  await openMenu(page);
  await expect(menu(page).getByRole('menuitem', { name: 'Remove space before paragraph' })).toBeVisible();
  await menu(page).getByRole('menuitem', { name: 'Remove space after paragraph' }).click();
  await caretIn(page, 'Alpha');
  await openMenu(page);
  await menu(page).getByRole('menuitem', { name: 'Add space before paragraph' }).click();
  const { xml } = await download(page);
  expect(savedParagraph(xml, 'Kept')).toContain('<w:spacing w:beforeLines="100" w:before="240" w:after="0"/>');
  expect(savedParagraph(xml, 'Alpha')).toContain('<w:spacing w:before="240"/>');
});

test('the Paragraph dialog shows the paragraph\'s spacing and writes only what is changed in it', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await caretIn(page, 'Kept');
  await openMenu(page);
  await menu(page).getByRole('menuitem', { name: 'Paragraph…' }).click();
  const dialog = page.getByRole('dialog', { name: 'Paragraph' });
  await expect(dialog.getByLabel('Before (pt)')).toHaveValue('12');
  await expect(dialog.getByLabel('After (pt)')).toHaveValue('6');
  await expect(dialog.getByLabel('Line spacing')).toHaveValue('single');
  await dialog.getByLabel('Left (cm)').fill('2.54');
  await dialog.getByLabel('Special').selectOption('hanging');
  await dialog.getByLabel('By (cm)').fill('1.27');
  await dialog.getByLabel('After (pt)').fill('9');
  await dialog.getByLabel('Line spacing').selectOption('exact');
  await expect(dialog.getByLabel('At (pt)')).toHaveValue('12');
  await dialog.getByLabel('At (pt)').fill('18');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(dialog).toBeHidden();
  const kept = page.locator('.docx-page p', { hasText: 'Kept' });
  await expect(kept).toHaveCSS('padding-left', '96px');
  await expect(kept).toHaveCSS('text-indent', '-48px');
  await expect(kept).toHaveCSS('line-height', '24px');
  const { xml } = await download(page);
  // Before (12 pt) and Word's "lines" spacing weren't touched: kept as they were.
  expect(savedParagraph(xml, 'Kept')).toBe('<w:p><w:pPr><w:keepNext/><w:tabs><w:tab w:val="left" w:pos="2000"/></w:tabs>'
    + '<w:spacing w:beforeLines="100" w:before="240" w:after="180" w:line="360" w:lineRule="exact"/><w:ind w:left="1440" w:hanging="720"/><w:jc w:val="center"/></w:pPr>'
    + `${run('Kept')}</w:p>`);
});

test('the Paragraph dialog changes nothing when OK is pressed without changes, and Cancel changes nothing', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await caretIn(page, 'Gamma');
  await openMenu(page);
  await menu(page).getByRole('menuitem', { name: 'Paragraph…' }).click();
  const dialog = page.getByRole('dialog', { name: 'Paragraph' });
  await dialog.getByRole('button', { name: 'OK' }).click();
  await openMenu(page);
  await menu(page).getByRole('menuitem', { name: 'Paragraph…' }).click();
  await dialog.getByLabel('Left (cm)').fill('5');
  await dialog.getByRole('button', { name: 'Cancel' }).click();
  await expect(page.locator('.app-header .dirty')).toHaveCount(0);
});

test('a first-line indent from the dialog; Increase Indent keeps it and moves the left indent', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await caretIn(page, 'Gamma');
  await openMenu(page);
  await menu(page).getByRole('menuitem', { name: 'Paragraph…' }).click();
  const dialog = page.getByRole('dialog', { name: 'Paragraph' });
  await dialog.getByLabel('Special').selectOption('firstLine');
  await dialog.getByLabel('By (cm)').fill('1.27');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await page.getByTitle('Increase indent (Ctrl+M)').click();
  const { xml } = await download(page);
  expect(savedParagraph(xml, 'Gamma')).toBe(`<w:p><w:pPr><w:ind w:firstLine="720" w:left="720"/></w:pPr>${run('Gamma')}</w:p>`);
});

test('while tracking, spacing and indentation are tracked paragraph changes', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await page.getByRole('button', { name: 'Track changes' }).click();
  const dialog = page.getByRole('dialog', { name: 'Your name on tracked changes' });
  await dialog.getByRole('textbox').fill('Rita');
  await dialog.getByRole('button', { name: 'Start tracking' }).click();
  await caretIn(page, 'Alpha');
  await page.keyboard.press('ControlOrMeta+2');
  await page.keyboard.press('ControlOrMeta+m');
  const { xml } = await download(page);
  expect(savedParagraph(xml, 'Alpha')).toMatch(/^<w:p><w:pPr><w:spacing w:line="480" w:lineRule="auto"\/><w:ind w:left="720"\/><w:pPrChange w:id="\d+" w:author="Rita" w:date="[^"]+"><w:pPr\/><\/w:pPrChange><\/w:pPr>/);
});

test('a list item set in from the dialog keeps its indentation, also after its list changed', async ({ page }) => {
  const NUM = '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>';
  await openInEditor(page, await buildDocx(`<w:p>${NUM}<w:r><w:t>One</w:t></w:r></w:p><w:p>${NUM}<w:r><w:t>Two</w:t></w:r></w:p>`));
  await page.locator('.docx-page p', { hasText: 'Two' }).click();
  await page.keyboard.press('Home');
  await page.keyboard.press('Tab'); // its list level changes...
  await openMenu(page);
  await menu(page).getByRole('menuitem', { name: 'Paragraph…' }).click();
  const dialog = page.getByRole('dialog', { name: 'Paragraph' });
  await expect(dialog.getByLabel('Left (cm)')).toHaveValue('2.54'); // ...to level 2's indentation
  await dialog.getByLabel('Left (cm)').fill('3');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(page.locator('.docx-page p', { hasText: 'Two' })).toHaveCSS('padding-left', /^113\.[34]/);
  const { xml } = await download(page);
  expect(savedParagraph(xml, 'Two')).toBe(`<w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr><w:ind w:left="1701"/></w:pPr>${run('Two')}</w:p>`);
});
