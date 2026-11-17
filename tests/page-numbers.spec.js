// Page numbers, as Word's Insert Page Number: Word's PAGE field at the caret in a header or footer;
// from the text, into an empty footer, or into a new footer for all pages when the document has
// none (its part, relationship, content type and the section's reference written only while it is
// in the page). A footer with text keeps its design: the caret goes there to say where.
import fs from 'node:fs';
import JSZip from 'jszip';
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, downloadFile, W_NS } from './helpers.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const part = (root, inner) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${root} ${W_NS}>${inner}</w:${root}>`;
const FIELD = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE   \\* MERGEFORMAT </w:instrText></w:r>'
  + '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t xml:space="preserve">1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';
const BODY = '<w:p><w:r><w:t>Body text.</w:t></w:r></w:p>';
const withFooter = (inner) => ({
  parts: { 'word/footer1.xml': part('ftr', inner) },
  rels: `<Relationship Id="rIdF1" Type="${REL}/footer" Target="footer1.xml"/>`,
  sectPr: '<w:footerReference w:type="default" r:id="rIdF1"/>',
});
const status = (page) => page.locator('.statusbar [role=status]');
const footer = (page) => page.locator('.docx-hf-band.footer .docx-hf-part');

async function saved(page) {
  const zip = await JSZip.loadAsync(fs.readFileSync(await downloadFile(page)));
  const text = async (name) => zip.file(name)?.async('string') ?? null;
  return { zip, text };
}

test('a document without a footer gets one with the page number, centred, for all pages', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await expect(footer(page)).toHaveCount(0);
  await page.locator('.docx-page p').first().click();
  await page.getByRole('button', { name: 'Page number' }).click();
  await expect(footer(page)).toHaveText('1');
  await expect(status(page)).toHaveText('Page number added to the footer: Word numbers the pages');
  const { text } = await saved(page);
  expect(await text('word/footer1.xml')).toContain(`<w:p><w:pPr><w:pStyle w:val="Footer"/><w:jc w:val="center"/></w:pPr>${FIELD}</w:p>`);
  expect(await text('word/document.xml')).toContain('<w:sectPr><w:footerReference w:type="default" r:id="rIdFooter1"/><w:pgSz');
  expect(await text('word/_rels/document.xml.rels')).toContain(`<Relationship Id="rIdFooter1" Type="${REL}/footer" Target="footer1.xml"/>`);
  expect(await text('[Content_Types].xml')).toContain('<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>');
  expect(await text('word/styles.xml')).toContain('w:styleId="Footer"');
});

test('Undo takes the new footer away, and it is not saved', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await page.locator('.docx-page p').first().click();
  await page.getByRole('button', { name: 'Page number' }).click();
  await expect(footer(page)).toHaveText('1');
  let { text } = await saved(page); // saved once with the footer...
  expect(await text('word/footer1.xml')).not.toBeNull();
  await page.locator('.docx-page p').first().click();
  await page.keyboard.press('ControlOrMeta+z');
  await expect(footer(page)).toHaveCount(0);
  await page.locator('.docx-page p').first().click();
  await page.keyboard.press('End');
  await page.keyboard.type(' More.'); // ...then saved without it
  ({ text } = await saved(page));
  expect(await text('word/footer1.xml')).toBeNull();
  expect(await text('word/document.xml')).not.toContain('footerReference');
  expect(await text('word/_rels/document.xml.rels')).not.toContain('footer');
  expect(await text('[Content_Types].xml')).not.toContain('footer');
});

test('with the caret in a footer, the page number goes at the caret', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY, withFooter('<w:p><w:r><w:t xml:space="preserve">Acme Ltd · Page </w:t></w:r></w:p>')));
  await footer(page).locator('p').click();
  await page.keyboard.press('End');
  await page.getByRole('button', { name: 'Page number' }).click();
  await expect(footer(page)).toHaveText('Acme Ltd · Page 1');
  await expect(footer(page).locator('.field')).toHaveAttribute('title', 'Page number – Word updates it');
  const { text } = await saved(page);
  expect(await text('word/footer1.xml')).toContain(`<w:p><w:r><w:t xml:space="preserve">Acme Ltd · Page </w:t></w:r>${FIELD}</w:p>`);
  expect(await text('word/document.xml')).toContain('<w:footerReference w:type="default" r:id="rIdF1"/><w:pgSz'); // no other footer
});

test('from the text, a footer with text is kept as it is: the caret goes there to say where', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY, withFooter('<w:p><w:r><w:t>Confidential</w:t></w:r></w:p>')));
  await page.locator('.docx-page > p').first().click();
  await page.getByRole('button', { name: 'Page number' }).click();
  await expect(status(page)).toHaveText('Click where the number goes in the footer (or header), then Page number again');
  await expect(footer(page)).toHaveText('Confidential');
  await page.getByRole('button', { name: 'Page number' }).click(); // the caret is now in the footer
  await expect(footer(page)).toHaveText('Confidential1');
});

test('from the text, an empty footer gets the number, centred', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY, withFooter('<w:p/>')));
  await page.locator('.docx-page > p').first().click();
  await page.getByRole('button', { name: 'Page number' }).click();
  await expect(footer(page)).toHaveText('1');
  const { text } = await saved(page);
  expect(await text('word/footer1.xml')).toContain(`<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${FIELD}</w:p>`);
});

test('while tracking, the page number is a tracked insertion', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY, withFooter('<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r></w:p>')));
  await page.getByRole('button', { name: 'Track changes' }).click();
  const dialog = page.getByRole('dialog', { name: 'Your name on tracked changes' });
  await dialog.getByRole('textbox').fill('Rita');
  await dialog.getByRole('button', { name: 'Start tracking' }).click();
  await footer(page).locator('p').click();
  await page.keyboard.press('End');
  await page.getByRole('button', { name: 'Page number' }).click();
  const { text } = await saved(page);
  expect(await text('word/footer1.xml')).toMatch(new RegExp(`<w:ins w:id="\\d+" w:author="Rita" w:date="[^"]+">${FIELD.replace(/[\\*]/g, (c) => `\\${c}`)}</w:ins>`));
});
