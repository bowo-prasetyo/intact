// Numbered headings (Numbering ▾ > 1 Heading 1 → 1.1 Heading 2 → 1.1.1 Heading 3), as Word's heading
// numbering writes them: one multilevel list whose levels are linked to the heading styles
// (w:pStyle in each w:lvl), and each heading style numbered by it (w:numPr in the style). The
// paragraphs aren't changed: their styles number them. Ctrl+Z puts the styles back.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const h = (n, text) => `<w:p><w:pPr><w:pStyle w:val="Heading${n}"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
const para = (text) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const BODY = h(1, 'Introduction') + para('Text.') + h(2, 'Scope') + h(2, 'Audience') + h(3, 'Operators')
  + h(1, 'Installation') + h(2, 'Requirements');
const markers = (page) => page.locator('.docx-page .list').evaluateAll((els) => els.map((e) => `${e.dataset.marker.trim()} ${e.textContent}`));
const library = (page) => page.getByRole('menu', { name: 'Numbering library' });

async function saved(page) {
  const { zip } = await download(page);
  return Object.fromEntries(await Promise.all(['word/document.xml', 'word/styles.xml', 'word/numbering.xml'].map(async (p) => [p, await zip.file(p)?.async('string')])));
}

test('the headings are numbered by their styles, 1, 1.1, 1.1.1, as Word links them', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await page.locator('.docx-page p', { hasText: 'Text.' }).click();
  await page.getByRole('button', { name: 'Numbering library' }).click();
  await library(page).getByRole('menuitem', { name: /^1 Heading 1/ }).click();
  await expect.poll(() => markers(page)).toEqual(['1 Introduction', '1.1 Scope', '1.2 Audience', '1.2.1 Operators', '2 Installation', '2.1 Requirements']);

  const files = await saved(page);
  const doc = files['word/document.xml'];
  // The paragraphs are as they were: no numbering of their own.
  expect(doc).not.toContain('<w:numPr>');
  // Heading 1 numbered at the first level (no w:ilvl), Heading 2 at the second, Heading 3 (added) at the third.
  const styles = files['word/styles.xml'];
  const numId = /w:styleId="Heading1".*?<w:numPr><w:numId w:val="(\d+)"\/><\/w:numPr>/s.exec(styles)[1];
  expect(styles).toMatch(new RegExp(`w:styleId="Heading2".*?<w:numPr><w:ilvl w:val="1"/><w:numId w:val="${numId}"/></w:numPr>`, 's'));
  expect(styles).toMatch(new RegExp(`w:styleId="Heading3".*?<w:numPr><w:ilvl w:val="2"/><w:numId w:val="${numId}"/></w:numPr>`, 's'));
  // The list's levels are linked to the styles, in CT_Lvl's order.
  const numbering = files['word/numbering.xml'];
  const abstractId = new RegExp(`<w:num w:numId="${numId}"><w:abstractNumId w:val="(\\d+)"/>`).exec(numbering)[1];
  const abstract = new RegExp(`<w:abstractNum w:abstractNumId="${abstractId}">.*?</w:abstractNum>`, 's').exec(numbering)[0];
  expect(abstract).toContain('<w:multiLevelType w:val="multilevel"/>');
  expect(abstract).toContain('<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Heading1"/><w:lvlText w:val="%1"/>');
  expect(abstract).toContain('<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Heading2"/><w:lvlText w:val="%1.%2"/>');

  // A heading added later is numbered too, by its style.
  await page.locator('.docx-page :is(p,h2)', { hasText: 'Requirements' }).click();
  await page.keyboard.press('End');
  await page.keyboard.press('Enter');
  await page.keyboard.type('Next');
  await page.getByRole('combobox', { name: 'Paragraph style' }).selectOption({ label: 'Heading 2' });
  await expect.poll(() => markers(page)).toContain('2.2 Next');
});

test('Ctrl+Z takes the numbering off again, and Remove heading numbering does too', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await page.locator('.docx-page p', { hasText: 'Text.' }).click();
  await page.getByRole('button', { name: 'Numbering library' }).click();
  await library(page).getByRole('menuitem', { name: /^1 Heading 1/ }).click();
  await expect(page.locator('.docx-page .list')).toHaveCount(6);
  await page.keyboard.press('ControlOrMeta+z');
  await expect(page.locator('.docx-page .list')).toHaveCount(0);

  await page.getByRole('button', { name: 'Numbering library' }).click();
  await library(page).getByRole('menuitem', { name: /^1 Heading 1/ }).click();
  await expect(page.locator('.docx-page .list')).toHaveCount(6);
  await page.getByRole('button', { name: 'Numbering library' }).click();
  await library(page).getByRole('menuitem', { name: /^Remove heading numbering/ }).click();
  await expect(page.locator('.docx-page .list')).toHaveCount(0);
  const styles = (await saved(page))['word/styles.xml'];
  for (const n of [1, 2, 3]) expect(new RegExp(`w:styleId="Heading${n}".*?</w:style>`, 's').exec(styles)[0]).not.toContain('<w:numPr>');
  expect(styles).toMatch(/w:styleId="Heading1"><w:name w:val="heading 1"\/><w:basedOn w:val="Normal"\/><w:rPr>/);
});
