// Word's Line and Page Breaks, in the Paragraph dialog: Keep with next, Keep lines together and Page
// break before. Written as the paragraph's own <w:keepNext/>, <w:keepLines/> and <w:pageBreakBefore/>
// only where its style doesn't say the same (off as w:val="0" where the style has it on), and
// followed by the page estimate (editor/pages.js).
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, W_NS } from './helpers.js';

const para = (text, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r><w:t>${text}</w:t></w:r></w:p>`;
const lines = (n) => Array.from({ length: n }, (_, i) => para(`Line ${i + 1} of the manual.`)).join('');
const LONG = 'This step explains in some detail what the operator does before the machine is started. '.repeat(6);
const savedParagraph = (xml, text) => xml.match(/<w:p[ >].*?<\/w:p>/gs).find((p) => p.includes(text));

// A style that keeps its paragraphs with the next, as a step's lead-in might.
const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${W_NS}>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Lead"><w:name w:val="Lead"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/></w:pPr></w:style>
</w:styles>`;

async function openDialog(page, text) {
  await page.locator('.docx-page p', { hasText: text }).first().click();
  await page.getByRole('button', { name: 'Line and paragraph spacing' }).click();
  await page.getByRole('menu', { name: 'Line and paragraph spacing' }).getByRole('menuitem', { name: 'Paragraph…' }).click();
  return page.getByRole('dialog', { name: 'Paragraph' });
}

test('the dialog shows the breaks a paragraph has, from its style too, and writes only what changes', async ({ page }) => {
  const body = para('Plain') + para('Lead in', '<w:pStyle w:val="Lead"/>') + para('Own', '<w:keepLines/><w:jc w:val="center"/>');
  await openInEditor(page, await buildDocx(body, { parts: { 'word/styles.xml': STYLES } }));

  let dialog = await openDialog(page, 'Plain');
  await expect(dialog.getByLabel('Keep with next')).not.toBeChecked();
  await dialog.getByLabel('Keep with next').check();
  await dialog.getByLabel('Page break before').check();
  await dialog.getByRole('button', { name: 'OK' }).click();

  dialog = await openDialog(page, 'Lead in');
  await expect(dialog.getByLabel('Keep with next')).toBeChecked(); // from its style
  await dialog.getByLabel('Keep with next').uncheck();
  await dialog.getByRole('button', { name: 'OK' }).click();

  dialog = await openDialog(page, 'Own');
  await expect(dialog.getByLabel('Keep lines together')).toBeChecked();
  await dialog.getByLabel('Keep lines together').uncheck();
  await dialog.getByRole('button', { name: 'OK' }).click();

  const { xml } = await download(page);
  expect(savedParagraph(xml, '>Plain<')).toContain('<w:pPr><w:keepNext/><w:pageBreakBefore/></w:pPr>');
  expect(savedParagraph(xml, '>Lead in<')).toContain('<w:pPr><w:pStyle w:val="Lead"/><w:keepNext w:val="0"/></w:pPr>');
  expect(savedParagraph(xml, '>Own<')).toContain('<w:pPr><w:jc w:val="center"/></w:pPr>');

  // Undone in one step each.
  await page.keyboard.press('ControlOrMeta+z');
  await page.keyboard.press('ControlOrMeta+z');
  await page.keyboard.press('ControlOrMeta+z');
  const after = await download(page);
  expect(savedParagraph(after.xml, '>Plain<')).not.toContain('keepNext');
});

test('Page break before starts a new page in the estimate', async ({ page }) => {
  await openInEditor(page, await buildDocx(lines(3) + para('Appendix')));
  const count = page.locator('.statusbar .page-count');
  await expect(count).toHaveText('≈ 1 page');
  const dialog = await openDialog(page, 'Appendix');
  await dialog.getByLabel('Page break before').check();
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(count).toHaveText('≈ 2 pages');
});

test('a paragraph kept together moves to the next page whole, with the one kept with it', async ({ page }) => {
  const mark = page.locator('.page-breaks-layer .page-estimate').first();
  // How many lines the first page holds with this browser's fonts: those above the first break.
  await openInEditor(page, await buildDocx(lines(90)));
  await expect(mark).toBeVisible();
  const y = (await mark.boundingBox()).y;
  const capacity = await page.locator('.docx-page p').evaluateAll((ps, at) => ps.filter((p) => p.getBoundingClientRect().top < at).length, y);
  expect(capacity).toBeGreaterThan(20);
  // Lines up to three from the end of the page, a lead-in, then a step of several lines that
  // starts on the page's last two lines and goes on to the next.
  await openInEditor(page, await buildDocx(lines(capacity - 3) + para('Lead in') + para(LONG)));
  const topOf = (text) => page.locator('.docx-page p', { hasText: text }).first().evaluate((p) => p.getBoundingClientRect().top);
  await expect(mark).toBeVisible();
  // The marks are drawn again when the estimate changes, and for a moment there is none to measure:
  // then the measure is "not yet" (Infinity), and the poll waits for the mark, rather than failing.
  const markY = async () => (await mark.boundingBox())?.y ?? null;
  // At first the long step is split across the pages.
  await expect.poll(async () => { const at = await markY(); return at == null ? -Infinity : at - (await topOf('This step explains')); }).toBeGreaterThan(10);

  // Measured together, as the page may scroll (Firefox scrolls the caret's paragraph into view).
  const gap = async (text) => { const at = await markY(); return at == null ? Infinity : Math.abs(at - (await topOf(text))); };

  let dialog = await openDialog(page, 'This step explains');
  await dialog.getByLabel('Keep lines together').check();
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect.poll(() => gap('This step explains')).toBeLessThan(12);

  dialog = await openDialog(page, 'Lead in');
  await dialog.getByLabel('Keep with next').check();
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect.poll(() => gap('Lead in')).toBeLessThan(12);
});
