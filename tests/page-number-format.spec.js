// Format Page Numbers, as Word's (Layout > Page number ▾ > Format page numbers…): the number format
// (1, 2, 3; - 1 -; a, b, c; i, ii, iii…) and where numbering starts, for the section the caret is in,
// the whole document, or a new section from the caret on (a report's front matter in i, ii, iii and
// its chapters from 1). Written in the section's w:pgNumType, in the order Word's schema requires,
// shown on each section break, and printed so.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, W_NS } from './helpers.js';
import JSZip from 'jszip';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const para = (text) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const field = (instr, result) => `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${result}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
const FOOTER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ${W_NS}><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">Page </w:t></w:r>${field('PAGE', '1')}</w:p></w:ftr>`;
const LONG = Array.from({ length: 90 }, (_, i) => para(`Line ${i + 1}.`)).join('');
const status = (page) => page.locator('.statusbar [role=status]');
const dialog = (page) => page.getByRole('dialog', { name: 'Page number format' });
const sections = (xml) => xml.match(/<w:sectPr>.*?<\/w:sectPr>/gs);

/** A document with a footer holding the page number, and `body`; `finalExtra` goes in the final section after its margins. */
async function withFooter(body, finalExtra = '') {
  const buffer = await buildDocx(body, {
    parts: { 'word/footer1.xml': FOOTER },
    rels: `<Relationship Id="rIdF" Type="${REL}/footer" Target="footer1.xml"/>`,
    sectPr: '<w:footerReference w:type="default" r:id="rIdF"/>',
  });
  if (!finalExtra) return buffer;
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file('word/document.xml').async('string');
  zip.file('word/document.xml', xml.replace(/(<w:pgMar [^>]*\/>)(<\/w:sectPr>\s*<\/w:body>)/, `$1${finalExtra}$2`));
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function openFormat(page, paragraph = 'Line 1.') {
  await page.locator('.docx-page p', { hasText: new RegExp(`^${paragraph.replace(/\./g, '\\.')}$`) }).first().click();
  await page.getByRole('button', { name: 'Page number and header options' }).click();
  await page.getByRole('menuitem', { name: /^Format page numbers/ }).click();
  await expect(dialog(page)).toBeVisible();
}

test('i, ii, iii from 3: written in the section\'s w:pgNumType, after its margins, and printed so', async ({ page, browserName }) => {
  await openInEditor(page, await withFooter(LONG));
  await openFormat(page);
  await expect(dialog(page).getByRole('combobox', { name: 'Number format', exact: true })).toHaveValue('decimal');
  await expect(dialog(page).getByLabel('Continue from previous section')).toBeChecked();
  await expect(dialog(page).getByLabel('Start at page number')).toBeDisabled();
  await dialog(page).getByRole('combobox', { name: 'Number format', exact: true }).selectOption('lowerRoman');
  await dialog(page).getByLabel('Start at', { exact: true }).check();
  await dialog(page).getByLabel('Start at page number').fill('3');
  await dialog(page).getByRole('button', { name: 'OK' }).click();
  await expect(status(page)).toHaveText('Page numbers: i, ii, iii, starting at 3 for the whole document');
  // Printing numbers the pages iii, iv…
  await page.evaluate(() => window.dispatchEvent(new Event('beforeprint')));
  const css = await page.evaluate(() => [...document.head.querySelectorAll('style')].map((s) => s.textContent).find((t) => t.includes('@page')));
  if (browserName === 'chromium') {
    expect(css).toContain('content: "Page " counter(page, lower-roman);');
    expect(css).toContain('@page :first { counter-set: page 3; }');
  }
  await page.evaluate(() => window.dispatchEvent(new Event('afterprint')));
  const { xml } = await download(page);
  expect(xml).toContain('<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:pgNumType w:fmt="lowerRoman" w:start="3"/></w:sectPr>');
});

test('the numbering a file has is shown, kept when untouched, and taken back to Arabic numerals that continue', async ({ page }) => {
  const buffer = await withFooter(LONG, '<w:pgNumType w:fmt="upperRoman" w:start="5" w:chapStyle="1"/>');
  await openInEditor(page, buffer);
  // Untouched, it goes back as it was.
  await page.locator('.docx-page p').first().click();
  await page.keyboard.press('End');
  await page.keyboard.type(' edited');
  let { xml } = await download(page);
  expect(xml).toContain('<w:pgNumType w:fmt="upperRoman" w:start="5" w:chapStyle="1"/>');
  await openFormat(page, 'Line 2.');
  await expect(dialog(page).getByRole('combobox', { name: 'Number format', exact: true })).toHaveValue('upperRoman');
  await expect(dialog(page).getByLabel('Start at page number')).toHaveValue('5');
  await dialog(page).getByRole('combobox', { name: 'Number format', exact: true }).selectOption('decimal');
  await dialog(page).getByLabel('Continue from previous section').check();
  await dialog(page).getByRole('button', { name: 'OK' }).click();
  ({ xml } = await download(page));
  // Word's default (Arabic numerals, continuing) is no format and no start; the chapter style stays.
  expect(xml).toContain('<w:pgNumType w:chapStyle="1"/>');
});

test('Ctrl+Z puts the numbering back; OK with nothing changed changes nothing', async ({ page }) => {
  await openInEditor(page, await withFooter(LONG));
  await openFormat(page);
  await dialog(page).getByRole('button', { name: 'OK' }).click();
  await expect(page.locator('.app-header .dirty')).toHaveCount(0);
  await openFormat(page);
  await dialog(page).getByRole('combobox', { name: 'Number format', exact: true }).selectOption('numberInDash');
  await dialog(page).getByRole('button', { name: 'OK' }).click();
  await expect(status(page)).toHaveText('Page numbers: - 1 -, - 2 -, - 3 - for the whole document');
  await page.locator('.docx-page p').first().click();
  await page.keyboard.press('ControlOrMeta+z');
  await page.keyboard.press('End');
  await page.keyboard.type(' edited');
  const { xml } = await download(page);
  expect(xml).not.toContain('pgNumType');
});

test('This point forward: front matter in i, ii, iii, the chapters from 1 in a new section', async ({ page }) => {
  const body = para('Preface') + para('Contents') + para('Chapter one') + LONG;
  await openInEditor(page, await withFooter(body));
  // The front matter: lower Roman numerals.
  await openFormat(page, 'Preface');
  await dialog(page).getByRole('combobox', { name: 'Number format', exact: true }).selectOption('lowerRoman');
  await dialog(page).getByRole('button', { name: 'OK' }).click();
  // From Chapter one: a new section, numbered 1, 2, 3 from 1.
  await openFormat(page, 'Chapter one');
  await expect(dialog(page).getByRole('combobox', { name: 'Number format', exact: true })).toHaveValue('lowerRoman');
  await dialog(page).getByRole('combobox', { name: 'Apply page number format to' }).selectOption('forward');
  await expect(dialog(page)).toContainText('A section break (next page) goes before the paragraph with the caret');
  await dialog(page).getByRole('combobox', { name: 'Number format', exact: true }).selectOption('decimal');
  await dialog(page).getByLabel('Start at', { exact: true }).check();
  await dialog(page).getByLabel('Start at page number').fill('1');
  await dialog(page).getByRole('button', { name: 'OK' }).click();
  await expect(status(page)).toHaveText('Page numbers: 1, 2, 3, starting at 1 from this point forward (1 section break added)');
  // The section break says how its section is numbered.
  await expect(page.locator('[data-section-note]')).toHaveAttribute('data-section-note', 'Section 1: A4 portrait · pages i, ii, iii · section break (next page)');
  // Now two sections: the dialog is for the caret's.
  await openFormat(page, 'Preface');
  await expect(dialog(page).getByRole('combobox', { name: 'Apply page number format to' })).toHaveValue('section');
  await expect(dialog(page).getByRole('combobox', { name: 'Number format', exact: true })).toHaveValue('lowerRoman');
  await dialog(page).getByRole('button', { name: 'Cancel' }).click();
  const { xml } = await download(page);
  const [front, chapters] = sections(xml);
  expect(front).toContain('<w:footerReference w:type="default" r:id="rIdF"/>');
  expect(front).toContain('<w:type w:val="nextPage"/>');
  expect(front).toContain('<w:pgNumType w:fmt="lowerRoman"/>');
  expect(chapters).toContain('<w:pgNumType w:start="1"/>');
  // The break is in the paragraph before Chapter one.
  expect(xml).toMatch(/<w:p><w:pPr><w:sectPr>(?:(?!<\/w:p>).)*<\/w:sectPr><\/w:pPr><w:r><w:t[^>]*>Contents<\/w:t><\/w:r><\/w:p><w:p><w:r><w:t[^>]*>Chapter one/s);
});
