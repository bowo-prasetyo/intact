// An index, as Word's References tab makes one: Mark index entry… (Alt+Shift+X) puts Word's XE field
// after the selected text (Mark all: in each paragraph it is in), which shows nothing; Insert ›
// Index puts "Index" and Word's INDEX field, built at once from the entries: one line an entry
// (Index 1), its subentries below it (Index 2), a letter's heading before its entries (Index
// Heading), with the pages they are about on; Update (F9) rebuilds it.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const P = (t) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
const XE = (entry) => `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> XE "${entry}" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
const MANUAL = P('The pump moves water.') + P('Check the valve weekly.') + P('Prime the pump before use.') + `<w:p><w:r><w:t>Seals wear out.</w:t></w:r>${XE('Pump:seals')}</w:p>` + P('End.');

async function insertMenu(page, item) {
  await page.getByRole('button', { name: /Insert a check box or date/ }).click();
  await page.getByRole('menuitem', { name: item }).click();
}

async function select(page, text, words) {
  await page.evaluate(([t, w]) => {
    const p = [...document.querySelectorAll('.docx-page p')].find((e) => e.textContent.includes(t));
    const node = [...p.childNodes].map((n) => (n.nodeType === 3 ? n : n.firstChild)).find((n) => n?.nodeType === 3 && n.data.includes(w));
    const at = node.data.indexOf(w);
    document.querySelector('.docx-page').focus();
    getSelection().setBaseAndExtent(node, at, node, at + w.length);
  }, [text, words]);
}

const indexLines = (page) => page.locator('.docx-page > p').evaluateAll((ps) => ps.filter((p) => /^Index/.test(p.dataset.style || '')).map((p) => `${p.dataset.style}|${p.textContent}`));

test('entries marked with Alt+Shift+X and Mark all, then an index built from them', async ({ page }) => {
  await openInEditor(page, await buildDocx(MANUAL));
  await select(page, 'Check the valve', 'valve');
  await page.keyboard.press('Alt+Shift+X');
  const dialog = page.getByRole('dialog', { name: 'Mark index entry' });
  await expect(dialog.getByLabel('Main entry')).toHaveValue('valve');
  await expect(dialog).toContainText('The document has 1 entry marked.');
  await dialog.getByLabel('Main entry').fill('Valve');
  await dialog.getByRole('button', { name: 'Mark', exact: true }).click();
  await expect(dialog).toBeHidden();
  // Nothing shows in the text.
  await expect(page.locator('.docx-page p', { hasText: 'Check the valve' })).toHaveText('Check the valve weekly.');

  await select(page, 'The pump moves', 'pump');
  await insertMenu(page, /^Mark index entry…/);
  await dialog.getByLabel('Main entry').fill('Pump');
  await dialog.getByRole('button', { name: 'Mark all' }).click();
  await expect(page.getByRole('status')).toContainText('2 index entries “Pump” marked');

  await page.locator('.docx-page p', { hasText: 'End.' }).click();
  await page.keyboard.press('End');
  await insertMenu(page, /^Index/);
  await expect.poll(() => indexLines(page)).toEqual([
    'IndexHeading|P', 'Index1|Pump, 1', 'Index2|seals, 1',
    'IndexHeading|V', 'Index1|Valve, 1',
  ]);

  const { zip, xml } = await download(page);
  expect(xml).toContain('<w:t xml:space="preserve">Check the valve</w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> XE "Valve" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>');
  expect(xml.match(/ XE "Pump" /g)).toHaveLength(2);
  expect(xml).toContain('<w:instrText xml:space="preserve"> INDEX \\h "A" \\c "1" </w:instrText>');
  expect(xml).toMatch(/<w:pStyle w:val="Heading1"\/><\/w:pPr><w:r><w:t xml:space="preserve">Index<\/w:t>/);
  const styles = await zip.file('word/styles.xml').async('string');
  expect(styles).toContain('w:styleId="Index1"><w:name w:val="index 1"/>');
  expect(styles).toContain('w:styleId="IndexHeading"><w:name w:val="index heading"/>');
  // Word works the page numbers out again when it opens the file.
  expect(await zip.file('word/settings.xml')?.async('string') ?? '').toContain('<w:updateFields w:val="true"/>');
});

test('F9 rebuilds the index with entries marked later, and their pages', async ({ page }) => {
  const filler = Array.from({ length: 60 }, (_, i) => P(`Line ${i + 1}.`)).join('');
  await openInEditor(page, await buildDocx(`<w:p><w:r><w:t>Alpha</w:t></w:r>${XE('Alpha')}</w:p>${filler}<w:p><w:r><w:t>Beta</w:t></w:r>${XE('Beta')}</w:p>${P('Last.')}`));
  await page.locator('.docx-page p', { hasText: 'Last.' }).click();
  await page.keyboard.press('End');
  await insertMenu(page, /^Index/);
  await expect.poll(() => indexLines(page)).toEqual(['IndexHeading|A', 'Index1|Alpha, 1', 'IndexHeading|B', 'Index1|Beta, 2']);
  await select(page, 'Line 3.', 'Line');
  await page.keyboard.press('Alt+Shift+X');
  await page.getByRole('dialog', { name: 'Mark index entry' }).getByRole('button', { name: 'Mark', exact: true }).click();
  await expect.poll(() => indexLines(page)).toContain('Index1|Line, 1');
  await page.keyboard.press('F9');
  await expect(page.getByRole('status')).toContainText('the index (3 entries)');
});
