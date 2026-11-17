// Update fields (F9, or the toolbar), as Word's for the whole text: caption numbers (SEQ), note
// numbers in cross-references (NOTEREF), cross-references (REF) and tables of contents (TOC),
// rebuilt from the headings in the table's own formatting. Page numbers are Word's: an entry keeps
// the number it had (a new one takes the number above it) and the file asks Word to update them.
import fs from 'node:fs';
import JSZip from 'jszip';
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, downloadFile, download, W_NS } from './helpers.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const t = (text) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const fld = (type) => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`;
const instr = (text) => `<w:r><w:instrText xml:space="preserve"> ${text} </w:instrText></w:r>`;
const field = (code, result) => `${fld('begin')}${instr(code)}${fld('separate')}${t(result)}${fld('end')}`;
const bm = (id, name, inner) => `<w:bookmarkStart w:id="${id}" w:name="${name}"/>${inner}<w:bookmarkEnd w:id="${id}"/>`;
const entry = (style, anchor, text, page) => `<w:hyperlink w:anchor="${anchor}" w:history="1">${t(text)}<w:r><w:tab/></w:r>${field(`PAGEREF ${anchor} \\h`, page)}</w:hyperlink>`;
const TABS = '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9016"/></w:tabs>';
// The table as Word last saved it: "Background" was called "Old name" then, and "Results" wasn't there.
const TOC_PARAS = `<w:p><w:pPr><w:pStyle w:val="TOC1"/>${TABS}</w:pPr>${fld('begin')}${instr('TOC \\o "1-3" \\h \\z \\u')}${fld('separate')}${entry('TOC1', '_Toc1', 'Introduction', '1')}</w:p>`
  + `<w:p><w:pPr><w:pStyle w:val="TOC2"/>${TABS}<w:ind w:left="220"/></w:pPr>${entry('TOC2', '_Toc2', 'Old name', '2')}</w:p>`
  + `<w:p>${fld('end')}</w:p>`;
const SDT = (paras) => `<w:sdt><w:sdtPr><w:id w:val="-1"/><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>`
  + `<w:p><w:r><w:t>Contents</w:t></w:r></w:p>${paras}</w:sdtContent></w:sdt>`;
const H1 = '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>';
const H2 = '<w:pPr><w:outlineLvl w:val="1"/></w:pPr>';
const NOTE_REF = '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>';
const BODY = (toc) => `${toc}`
  + `<w:p>${H1}${bm(10, '_Toc1', t('Introduction'))}</w:p>`
  + `<w:p>${t('As ')}${field('REF _Ref1 \\h', 'Figure 9')}${t(' shows, and note ')}${field('NOTEREF _Ref2 \\h', '7')}${t(' says.')}</w:p>`
  + `<w:p>${H2}${bm(11, '_Toc2', t('Background'))}</w:p>`
  + `<w:p>${t('A claim.')}${bm(12, '_Ref2', NOTE_REF)}</w:p>`
  + `<w:p>${bm(13, '_Ref1', `${t('Figure ')}${field('SEQ Figure \\* ARABIC', '3')}`)}${t(': the first chart')}</w:p>`
  + `<w:p>${t('Figure ')}${field('SEQ Figure \\* ARABIC', '7')}${t(': the second chart')}</w:p>`
  + `<w:p>${H1}${t('Results')}</w:p>`;
const FOOTNOTES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes ${W_NS}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>`
  + `<w:footnote w:id="1"><w:p><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>${t(' The note.')}</w:p></w:footnote></w:footnotes>`;
const SETTINGS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ${W_NS}><w:zoom w:percent="100"/></w:settings>`;
const OPTS = {
  parts: { 'word/footnotes.xml': FOOTNOTES, 'word/settings.xml': SETTINGS },
  rels: `<Relationship Id="rIdF" Type="${REL}/footnotes" Target="footnotes.xml"/><Relationship Id="rIdS" Type="${REL}/settings" Target="settings.xml"/>`,
};
const status = (page) => page.locator('.statusbar [role=status]');
const results = (page) => page.locator('.docx-page > p .field').evaluateAll((els) => els.map((e) => e.textContent));

async function update(page) {
  await page.locator('.docx-page > p', { hasText: 'A claim.' }).click();
  await page.keyboard.press('F9');
}

test('F9 renumbers captions, updates cross-references and note numbers, and rebuilds the table of contents', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY(SDT(TOC_PARAS)), OPTS));
  await expect(page.locator('.sdt-locked')).toContainText('Old name');
  await update(page);
  await expect(status(page)).toHaveText('Updated the table of contents (3 entries), 1 cross-reference, 2 caption numbers and 1 note number; page numbers are Word’s: the file asks Word to update them when it opens it');
  await expect.poll(() => results(page)).toEqual(['Figure 1', '1', '1', '2']);
  await expect(page.locator('.sdt-locked p')).toHaveText(['Contents', 'Introduction\t1', 'Background\t2', 'Results\t2', '']);

  const zip = await JSZip.loadAsync(fs.readFileSync(await downloadFile(page)));
  const xml = await zip.file('word/document.xml').async('string');
  // The table, as Word writes one, in its own formatting; the page number each entry had.
  expect(xml).toContain(`<w:p><w:pPr><w:pStyle w:val="TOC2"/>${TABS}<w:ind w:left="220"/></w:pPr><w:hyperlink w:anchor="_Toc2" w:history="1"><w:r><w:t xml:space="preserve">Background</w:t></w:r>`);
  const added = xml.match(/<w:bookmarkStart w:id="\d+" w:name="(_Toc\d+)"\/><w:r><w:t xml:space="preserve">Results/);
  expect(added).not.toBeNull(); // the new heading got Word's bookmark...
  expect(xml).toContain(`<w:hyperlink w:anchor="${added[1]}" w:history="1"><w:r><w:t xml:space="preserve">Results</w:t></w:r>`); // ...and the entry links to it
  expect(xml).toContain(`<w:instrText xml:space="preserve"> PAGEREF ${added[1]} \\h </w:instrText>`);
  expect(xml).toContain(`<w:p>${fld('end')}</w:p></w:sdtContent></w:sdt>`);
  expect(xml).toContain('<w:t xml:space="preserve">Figure </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SEQ Figure \\* ARABIC </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t xml:space="preserve">2</w:t></w:r>');
  expect(await zip.file('word/settings.xml').async('string')).toContain('<w:updateFields w:val="true"/>');
});

test('Ctrl+Z after F9 puts the fields back as they were', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY(SDT(TOC_PARAS)), OPTS));
  await update(page);
  await expect.poll(() => results(page)).toEqual(['Figure 1', '1', '1', '2']);
  await page.keyboard.press('ControlOrMeta+z');
  await expect.poll(() => results(page)).toEqual(['Figure 9', '7', '3', '7']);
  await expect(page.locator('.sdt-locked')).toContainText('Old name');
});

test('a table of contents in the text (no content control) is rebuilt the same way', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY(TOC_PARAS), OPTS));
  await page.getByRole('button', { name: 'Update fields' }).click();
  await expect(page.locator('.docx-page > p').first()).toHaveText('Introduction\t1');
  const { xml } = await download(page);
  expect(xml).toMatch(/^.*<w:body><w:p><w:pPr><w:pStyle w:val="TOC1"\/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9016"\/><\/w:tabs><\/w:pPr><w:r><w:fldChar w:fldCharType="begin"\/><\/w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u <\/w:instrText><\/w:r><w:r><w:fldChar w:fldCharType="separate"\/><\/w:r><w:hyperlink w:anchor="_Toc1" w:history="1">/s);
  expect(xml).toContain('Background</w:t></w:r>');
  expect(xml).not.toContain('Old name');
});

test('a cross-reference to a numbered paragraph, with "above" or "below"', async ({ page }) => {
  const NUM = '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>';
  const body = `<w:p>${t('See clause ')}${field('REF _Ref5 \\r \\h', '9')}${t(', ')}${field('REF _Ref5 \\p \\h', 'here')}${t('.')}</w:p>`
    + `<w:p>${NUM}${t('First clause')}</w:p><w:p>${NUM}${bm(20, '_Ref5', t('Second clause'))}</w:p>`;
  await openInEditor(page, await buildDocx(body));
  await page.locator('.docx-page > p').first().click();
  await page.keyboard.press('F9');
  await expect.poll(() => results(page)).toEqual(['2', 'below']);
  await expect(status(page)).toHaveText('Updated 2 cross-references');
});

test('a cross-reference to text no longer there is left as it was, and said so', async ({ page }) => {
  await openInEditor(page, await buildDocx(`<w:p>${t('See ')}${field('REF _Gone \\h', 'Table 4')}</w:p>`));
  await page.locator('.docx-page > p').first().click();
  await page.keyboard.press('F9');
  await expect(status(page)).toHaveText('The fields are up to date; 1 cross-reference to text no longer in the document was left as it was');
  await expect.poll(() => results(page)).toEqual(['Table 4']);
});

test('the review says the fields were updated here, and that Word updates the page numbers', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY(SDT(TOC_PARAS)), OPTS));
  await page.locator('.docx-page > p', { hasText: 'A claim.' }).click();
  await page.keyboard.press('End');
  await page.keyboard.type(' More.');
  await page.keyboard.press('F9');
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  const review = page.getByRole('dialog', { name: 'Review your changes' });
  await expect(review).toContainText('were updated here (Update fields, F9). Page numbers are Word’s: the file asks Word to update them when it is opened.');
});
