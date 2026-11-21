// A list of figures or of tables, as Word's References > Insert Table of Figures: Word's TOC field
// with \c and the caption label, its entries the captions of that label (each linked to its caption
// through a _Toc bookmark, in Word's Table of Figures style), under a heading in TOC Heading style;
// Update (F9) adds captions written later.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const P = (t) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
const CAPTION = (label, n, text) => `<w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr><w:r><w:t xml:space="preserve">${label} </w:t></w:r>`
  + `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SEQ ${label} \\* ARABIC </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>`
  + `<w:r><w:t>${n}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve">: ${text}</w:t></w:r></w:p>`;
const THESIS = P('Front matter') + P('Chapter text.') + CAPTION('Figure', 1, 'Apparatus') + CAPTION('Table', 1, 'Samples')
  + P('More text.') + CAPTION('Figure', 2, 'Results');
const status = (page) => page.locator('.statusbar [role=status]');

async function insertMenu(page, item) {
  await page.getByRole('button', { name: /Insert a check box or date/ }).click();
  await page.getByRole('menuitem', { name: item }).click();
}
async function caretAfter(page, text) {
  await page.locator('.docx-page > p', { hasText: text }).first().click();
  await page.keyboard.press('End');
}

test('List of figures: the Figure captions, linked, under "List of Figures", as Word\'s TOC \\c field', async ({ page }) => {
  await openInEditor(page, await buildDocx(THESIS));
  await caretAfter(page, 'Front matter');
  await page.keyboard.press('Enter');
  await insertMenu(page, /List of figures/);
  await expect(status(page)).toContainText('List of figures inserted: 2 captions');
  const blocks = page.locator('.docx-page > p');
  await expect(blocks.nth(1)).toHaveText('List of Figures');
  await expect(blocks.nth(2)).toContainText('Figure 1: Apparatus');
  await expect(blocks.nth(3)).toContainText('Figure 2: Results');
  await expect(page.locator('.docx-page > p', { hasText: /^Table 1: Samples/ })).toHaveCount(1); // not listed: only the caption itself

  const { zip, xml } = await download(page);
  expect(xml).toMatch(/<w:pStyle w:val="TOCHeading"\/><\/w:pPr><w:r><w:t xml:space="preserve">List of Figures<\/w:t>/);
  expect(xml).toContain(' TOC \\h \\z \\c "Figure" ');
  expect(xml).toMatch(/<w:pStyle w:val="TableofFigures"\/>.*?<w:hyperlink w:anchor="_Toc\d+" w:history="1"><w:r>.*?<w:t xml:space="preserve">Figure 1: Apparatus<\/w:t>/);
  expect(xml).toMatch(/<w:bookmarkStart w:id="\d+" w:name="_Toc\d+"\/><w:r><w:t xml:space="preserve">Figure <\/w:t>/);
  expect(xml).not.toMatch(/<w:hyperlink[^>]*>(?:(?!<\/w:hyperlink>).)*Table 1: Samples/);
  const styles = await zip.file('word/styles.xml').async('string');
  expect(styles).toContain('<w:style w:type="paragraph" w:styleId="TableofFigures"><w:name w:val="table of figures"/>');
  expect(await zip.file('word/settings.xml').async('string')).toContain('<w:updateFields w:val="true"/>');
});

test('List of tables lists the Table captions; Update (F9) adds one written later, in its place', async ({ page }) => {
  await openInEditor(page, await buildDocx(THESIS));
  await caretAfter(page, 'Front matter');
  await page.keyboard.press('Enter');
  await insertMenu(page, /List of tables/);
  await expect(status(page)).toContainText('List of tables inserted: 1 caption');
  await expect(page.locator('.docx-page > p').nth(1)).toHaveText('List of Tables');
  await expect(page.locator('.docx-page > p').nth(2)).toContainText('Table 1: Samples');

  // A new table caption at the end, then F9.
  await caretAfter(page, 'Figure 2: Results');
  await page.keyboard.press('Enter');
  await insertMenu(page, /Caption/);
  const cap = page.getByRole('dialog', { name: 'Caption' });
  await cap.getByLabel('Label').selectOption('Table');
  await cap.getByLabel('Caption text').fill('Costs');
  await cap.getByRole('button', { name: 'Insert' }).click();
  await page.keyboard.press('F9');
  await expect(page.locator('.docx-page > p').nth(3)).toContainText('Table 2: Costs');
  const { xml } = await download(page);
  expect(xml).toContain(' TOC \\h \\z \\c "Table" ');
  expect(xml).toMatch(/<w:hyperlink w:anchor="_Toc\d+" w:history="1">.*?Table 2: Costs/);
});

test('a list of figures with no captions yet says how to fill it', async ({ page }) => {
  await openInEditor(page, await buildDocx(P('Front matter') + P('Text.')));
  await caretAfter(page, 'Front matter');
  await page.keyboard.press('Enter');
  await insertMenu(page, /List of figures/);
  await expect(status(page)).toHaveText('List of figures inserted: no figure captions yet – add them with Caption…, then Update (F9)');
  await expect(page.locator('.docx-page > p').nth(2)).toHaveText('No table of figures entries found.');
});
