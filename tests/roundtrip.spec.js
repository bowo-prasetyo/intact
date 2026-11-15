// The product promise: "nothing you didn't touch changes". These tests open documents, save them
// (with and without edits) and check that everything the user did not edit means the same thing.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, canonicalize, documentXmlOf, FIELDS } from './helpers.js';

const NUM = '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>';

// One paragraph per feature, so a failure points at the feature.
const KITCHEN_SINK = `
<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Heading</w:t></w:r></w:p>
<w:p><w:pPr><w:jc w:val="both"/><w:spacing w:after="120"/></w:pPr>
  <w:r><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:b w:val="1"/><w:color w:val="C00000"/><w:sz w:val="28"/></w:rPr><w:t>Bold red Georgia</w:t></w:r>
  <w:r><w:t xml:space="preserve"> plain </w:t></w:r>
  <w:r><w:rPr><w:i/><w:u w:val="double"/></w:rPr><w:t>italic double-underline</w:t></w:r>
  <w:r><w:rPr><w:rStyle w:val="Strong"/><w:b w:val="0"/></w:rPr><w:t xml:space="preserve"> style with bold off</w:t></w:r>
  <w:r><w:rPr><w:dstrike/></w:rPr><w:t xml:space="preserve"> double strike</w:t></w:r>
  <w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>2</w:t></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">Tab:</w:t></w:r><w:r><w:tab/><w:t>after tab, line</w:t><w:br/><w:t>break, §\u00a05 (non-breaking space) and a ZWSP\u200bhere</w:t></w:r></w:p>
<w:p>${NUM}<w:r><w:t>First item</w:t></w:r></w:p>
<w:p>${NUM}<w:r><w:t>Second item</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="ListBullet"/></w:pPr><w:r><w:t>Bullet from a list style</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="ListBullet"/><w:numPr><w:ilvl w:val="0"/></w:numPr><w:ind w:left="500"/></w:pPr><w:r><w:t>Style bullet with its own level and indent</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="ListBullet"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr></w:pPr><w:r><w:t>List style with numbering switched off</w:t></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">Link: </w:t></w:r><w:hyperlink r:id="rId3" w:history="1"><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>example</w:t></w:r></w:hyperlink><w:bookmarkStart w:id="0" w:name="mark"/><w:r><w:t xml:space="preserve"> bookmarked</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>
<w:p><w:r><w:t xml:space="preserve">Kept </w:t></w:r><w:ins w:id="1" w:author="Ann" w:date="2026-01-01T00:00:00Z"><w:r><w:t>inserted by Ann</w:t></w:r></w:ins><w:del w:id="2" w:author="Ann" w:date="2026-01-01T00:00:00Z"><w:r><w:delText xml:space="preserve"> deleted by Ann</w:delText></w:r></w:del></w:p>
<w:p><w:r><w:t xml:space="preserve">Page: </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple><w:r><w:t xml:space="preserve"> of </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>3</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">Image: </w:t></w:r><w:r><w:drawing><wp:inline><wp:extent cx="95250" cy="95250"/><wp:docPr id="1" name="Picture 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rId4"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
<w:p><w:r><w:t>Before break</w:t></w:r><w:r><w:br w:type="page"/></w:r><w:r><w:t>After break</w:t></w:r></w:p>
<w:sdt><w:sdtPr><w:alias w:val="Clause"/><w:tag w:val="clause"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Clause in a content control</w:t></w:r></w:p></w:sdtContent></w:sdt>
<w:sdt><w:sdtPr><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Locked clause</w:t></w:r></w:p></w:sdtContent></w:sdt>
<w:p><w:r><w:t xml:space="preserve">Fields: </w:t></w:r>${FIELDS}</w:p>
<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>
  <w:tr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/><w:gridSpan w:val="2"/><w:shd w:val="clear" w:fill="D9E2F3"/></w:tcPr><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Merged header</w:t></w:r></w:p></w:tc></w:tr>
  <w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>Tall</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc></w:tr>
  <w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc><w:tc><w:p><w:r><w:t>B3</w:t></w:r></w:p></w:tc></w:tr>
</w:tbl>
<w:p><w:r><w:t>Last paragraph</w:t></w:r></w:p>`;

test('opening and downloading without edits keeps the document unchanged', async ({ page }) => {
  const docx = await buildDocx(KITCHEN_SINK);
  await openInEditor(page, docx);
  const { xml } = await download(page);
  const original = await canonicalize(page, await documentXmlOf(docx));
  expect(await canonicalize(page, xml)).toEqual(original);
});

test('a second save of a saved file changes nothing', async ({ page }) => {
  await openInEditor(page, await buildDocx(KITCHEN_SINK));
  const first = await download(page);
  await page.getByRole('button', { name: 'Close' }).click();
  await openInEditor(page, await first.zip.generateAsync({ type: 'nodebuffer' }));
  const second = await download(page);
  expect(await canonicalize(page, second.xml)).toEqual(await canonicalize(page, first.xml));
});

test('editing one paragraph leaves every other paragraph unchanged', async ({ page }) => {
  const docx = await buildDocx(KITCHEN_SINK);
  await openInEditor(page, docx);
  const last = page.locator('.docx-page > p').last();
  await last.click();
  await page.keyboard.press('End');
  await page.keyboard.type(' edited');
  const { xml } = await download(page);

  const before = await canonicalize(page, await documentXmlOf(docx));
  const after = await canonicalize(page, xml);
  expect(after.length).toBe(before.length);
  const edited = before.findIndex((b) => b.p?.inlines.some((i) => i.text === 'Last paragraph'));
  after.forEach((block, i) => {
    if (i !== edited) expect(block, `block ${i}`).toEqual(before[i]);
  });
  expect(after[edited].p.inlines).toEqual([{ rPr: '', text: 'Last paragraph edited' }]);
});

test('tracked changes keep their author and date, and are shown read-only', async ({ page }) => {
  await openInEditor(page, await buildDocx(`<w:p><w:r><w:t xml:space="preserve">Hello </w:t></w:r><w:ins w:id="1" w:author="Ann" w:date="2026-01-01T00:00:00Z"><w:r><w:t>world</w:t></w:r></w:ins></w:p>`));
  const ins = page.locator('.docx-page .raw-inline.inserted');
  await expect(ins).toHaveText('world');
  await expect(ins).toHaveAttribute('contenteditable', 'false');
  await expect(ins).toHaveAttribute('title', /Inserted by Ann/);

  await page.locator('.docx-page p').first().click();
  await page.keyboard.press('Home');
  await page.keyboard.type('Oh, ');
  const { xml } = await download(page);
  expect(xml).toContain('w:author="Ann"');
  expect(xml).toMatch(/<w:ins [^>]*><w:r><w:t>world<\/w:t><\/w:r><\/w:ins>/);
  expect(xml).toContain('Oh, Hello ');
});
