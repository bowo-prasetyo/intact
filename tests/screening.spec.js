// A CV's check for employers' screening software (editor/screening.js), in the review: what such
// software is known to miss or mix up. Open, and counted among the things to check, only for a
// document that reads as a CV; there, closed, for any other.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, W_NS } from './helpers.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const para = (text) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const heading = (text) => `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
const CV = heading('Experience') + para('Marketing Manager, Globex, 2021 to now') + heading('Education') + para('B.Sc. Business, 2017') + heading('Skills') + para('SQL, HubSpot');
const headerWith = (text) => ({
  parts: { 'word/header1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${W_NS}>${para(text)}</w:hdr>` },
  rels: `<Relationship Id="rIdH1" Type="${REL}/header" Target="header1.xml"/>`,
  sectPr: '<w:headerReference w:type="default" r:id="rIdH1"/>',
});

const review = (page) => page.getByRole('dialog', { name: 'Review your changes' });
/** How many things the review's first line says there are to check. */
async function toCheck(page) {
  const text = await review(page).locator('.overview').innerText();
  return Number(text.match(/^(\d+) things? to check/)?.[1] || 0);
}
const section = (page) => review(page).locator('section.screening');
async function openReview(page) {
  await page.getByRole('button', { name: 'Review', exact: true }).click();
  await expect(review(page)).toBeVisible();
}

test('contact details only in the header of a CV are flagged, counted and shown', async ({ page }) => {
  // The same file as a letter, then as a CV: one more thing to check.
  const docx = await buildDocx(para('Alex Morgan') + para('Marketing Manager, Globex'), headerWith('alex@example.com · +62 812 0000 0000'));
  await openInEditor(page, docx, 'Alex_Morgan_letter.docx');
  await openReview(page);
  await expect(section(page).locator('details')).not.toHaveAttribute('open', '');
  const asLetter = await toCheck(page);
  await review(page).getByRole('button', { name: 'Keep editing' }).click();
  await page.getByRole('button', { name: 'Close' }).click();
  await page.setInputFiles('input[type=file]', { name: 'Alex_Morgan_CV.docx', mimeType: 'application/octet-stream', buffer: docx });
  await openReview(page);
  await expect(section(page).locator('details')).toHaveAttribute('open', '');
  await expect(section(page)).toContainText('CV screening software');
  await expect(section(page)).toContainText('Your e-mail address and phone number are only in the header');
  expect(await toCheck(page)).toBe(asLetter + 1);
  await section(page).getByRole('button', { name: 'Show' }).click();
  await expect(review(page)).toBeHidden();
  await expect(page.locator('.statusbar [role=status]')).toContainText('Contact details in the header or footer');
});

test('contact details in the text too are fine', async ({ page }) => {
  await openInEditor(page, await buildDocx(para('Alex Morgan · alex@example.com · +62 812 0000 0000') + CV, headerWith('alex@example.com')), 'Alex_Morgan_CV.docx');
  await openReview(page);
  await expect(section(page)).toContainText('Nothing found that such software is known to miss');
  await expect(section(page).locator('.count')).toHaveText('none found');
});

test('a CV without an e-mail address, laid out in a table and columns, called "CV"', async ({ page }) => {
  const table = '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Manager, Globex</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>2021 to now</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
  const columns = '<w:p><w:pPr><w:sectPr><w:cols w:num="2" w:space="720"/><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr><w:r><w:t>Two columns of skills</w:t></w:r></w:p>';
  await openInEditor(page, await buildDocx(CV + table + para('Left column') + columns + para('After')), 'CV.docx');
  await openReview(page);
  await expect(section(page)).toContainText('No e-mail address');
  await expect(section(page)).toContainText('Text in columns');
  await expect(section(page)).toContainText('Text laid out in a table');
  await expect(section(page)).toContainText('The file is called “CV”');
  await expect(section(page).locator('.count')).toHaveText('2 to check, 2 tips');
});

test('pictures are a tip: no text is read in them', async ({ page }) => {
  const picture = '<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="952500" cy="952500"/><wp:docPr id="1" name="Photo"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="1" name="Photo"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId4"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
  await openInEditor(page, await buildDocx(para('alex@example.com') + picture + CV), 'Alex_Morgan_CV.docx');
  await openReview(page);
  await expect(section(page)).toContainText('Tip: Pictures.');
  await expect(section(page).locator('.count')).toHaveText('1 tip');
});

test('for a document that isn\'t a CV the check is there, closed, and not counted', async ({ page }) => {
  await openInEditor(page, await buildDocx(para('Dear Sir,') + para('Please find the contract attached.'), headerWith('legal@example.com')), 'Contract.docx');
  await openReview(page);
  await expect(section(page).locator('details')).not.toHaveAttribute('open', '');
  // Its warning (contact details in the header) isn't among the things to check.
  await expect(section(page).locator('li.warn')).toHaveCount(1);
  const count = await toCheck(page);
  const others = await review(page).locator('li.warn:not(section.screening li)').count();
  expect(count).toBeLessThanOrEqual(others);
});

test('Help opens the review at the check', async ({ page }) => {
  await openInEditor(page, await buildDocx(para('Text')));
  await page.getByRole('button', { name: 'Help' }).click();
  await page.getByRole('button', { name: /Check a CV for screening software/ }).click();
  await expect(section(page).locator('details')).toHaveAttribute('open', '');
});
