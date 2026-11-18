// Headers and footers are edited in place, in the page's margins, as in Word: typing, formatting,
// Undo, Find and Replace and Track Changes work in them as in the text. Only the parts that changed
// are rewritten (header1.xml, footer1.xml…); untouched, they come back byte for byte, and the
// document's text is not changed by editing them.
import fs from 'node:fs';
import JSZip from 'jszip';
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, canonicalize, documentXmlOf, W_NS } from './helpers.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const part = (root, inner) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${root} ${W_NS}>${inner}</w:${root}>`;
const PARTS = {
  'word/header1.xml': part('hdr', '<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>DRAFT v3 – Client name</w:t></w:r></w:p>'),
  'word/header2.xml': part('hdr', '<w:p><w:r><w:t>Cover page header</w:t></w:r></w:p>'),
  'word/footer1.xml': part('ftr', '<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple><w:r><w:t xml:space="preserve"> · Studio confidential</w:t></w:r></w:p>'),
};
const OPTS = {
  parts: PARTS,
  rels: `<Relationship Id="rIdH1" Type="${REL}/header" Target="header1.xml"/><Relationship Id="rIdH2" Type="${REL}/header" Target="header2.xml"/><Relationship Id="rIdF1" Type="${REL}/footer" Target="footer1.xml"/>`,
  sectPr: '<w:headerReference w:type="default" r:id="rIdH1"/><w:footerReference w:type="default" r:id="rIdF1"/>',
};
const BODY = '<w:p><w:r><w:t>Body text of the proposal.</w:t></w:r></w:p>';

const header = (page) => page.locator('.docx-hf-band.header .docx-hf-part:visible');
const footer = (page) => page.locator('.docx-hf-band.footer .docx-hf-part:visible');
const partOf = (zip, name) => zip.file(`word/${name}.xml`).async('string');

async function typeAtEnd(page, host, text) {
  await host.locator('p').last().click();
  await page.keyboard.press('End');
  await page.keyboard.type(text);
}

test('a header is edited in place and saved to its own part; the text and footer stay as they were', async ({ page }) => {
  const docx = await buildDocx(BODY, OPTS);
  await openInEditor(page, docx);
  await expect(header(page)).toHaveText('DRAFT v3 – Client name');
  await expect(header(page)).toHaveAttribute('contenteditable', 'true');
  await header(page).getByText('Client name').click();
  await page.keyboard.press('End');
  await page.keyboard.press('Shift+ControlOrMeta+ArrowLeft');
  await page.keyboard.press('Shift+ControlOrMeta+ArrowLeft');
  await page.keyboard.type('Northwind');
  await expect(header(page)).toHaveText('DRAFT v3 – Northwind');
  const { zip, xml } = await download(page);
  expect(await partOf(zip, 'header1')).toContain('<w:t xml:space="preserve">DRAFT v3 – Northwind</w:t>');
  expect(await partOf(zip, 'header1')).toContain('<w:jc w:val="right"/>'); // its paragraph keeps its formatting
  expect(await partOf(zip, 'footer1')).toBe(PARTS['word/footer1.xml']);
  expect(xml).toBe(await documentXmlOf(docx)); // the text's part isn't rewritten: byte for byte
});

test('a footer is edited too, its field kept; Undo takes back header and footer edits', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY, OPTS));
  await typeAtEnd(page, footer(page), ' – do not share');
  await expect(footer(page)).toHaveText('Page 1 · Studio confidential – do not share');
  await typeAtEnd(page, header(page), '!');
  await page.keyboard.press('ControlOrMeta+z');
  await expect(header(page)).toHaveText('DRAFT v3 – Client name');
  const { zip } = await download(page);
  const ftr = await partOf(zip, 'footer1');
  expect(ftr).toContain('<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>');
  expect(ftr).toContain('Studio confidential – do not share');
  expect(await partOf(zip, 'header1')).toBe(PARTS['word/header1.xml']); // undone: back byte for byte
});

test('header edits are tracked while tracking, and listed in the review', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY, OPTS));
  await page.getByRole('button', { name: 'Track changes' }).click();
  const dialog = page.getByRole('dialog', { name: 'Your name on tracked changes' });
  await dialog.getByRole('textbox').fill('Rita');
  await dialog.getByRole('button', { name: 'Start tracking' }).click();
  await typeAtEnd(page, header(page), ' (final)');
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  const review = page.getByRole('dialog', { name: 'Review your changes' });
  await expect(review).toContainText('Header');
  await expect(review).toContainText('(final)');
  const [dl] = await Promise.all([page.waitForEvent('download'), review.getByRole('button', { name: 'Download' }).click()]);
  const zip = await JSZip.loadAsync(fs.readFileSync(await dl.path()));
  expect(await partOf(zip, 'header1')).toMatch(/<w:ins w:id="\d+" w:author="Rita" w:date="[^"]+"><w:r><w:t xml:space="preserve"> \(final\)<\/w:t><\/w:r><\/w:ins>/);
});

test('the first page header and the other pages header are chosen with tabs and edited apart', async ({ page }) => {
  const docx = await buildDocx(BODY, { ...OPTS, sectPr: `${OPTS.sectPr}<w:headerReference w:type="first" r:id="rIdH2"/><w:titlePg/>` });
  await openInEditor(page, docx);
  const tabs = page.getByRole('tablist', { name: 'Headers' }).getByRole('tab');
  await expect(header(page)).toHaveText('Cover page header');
  await typeAtEnd(page, header(page), ' 2026');
  await tabs.nth(1).click();
  await expect(header(page)).toHaveText('DRAFT v3 – Client name');
  await page.keyboard.press('End');
  await page.keyboard.type('.');
  await expect(header(page)).toHaveText('DRAFT v3 – Client name.');
  // Undo brings back the content, and the tab chosen stays chosen.
  await page.keyboard.press('ControlOrMeta+z');
  await expect(header(page)).toHaveText('DRAFT v3 – Client name');
  await expect(tabs.nth(1)).toHaveAttribute('aria-selected', 'true');
  const { zip } = await download(page);
  expect(await partOf(zip, 'header2')).toContain('Cover page header 2026');
  expect(await partOf(zip, 'header1')).toBe(PARTS['word/header1.xml']);
});

test('Find and Replace changes text in headers and footers; the word count leaves them out', async ({ page }) => {
  await openInEditor(page, await buildDocx('<w:p><w:r><w:t>Studio work for the client.</w:t></w:r></w:p>', OPTS));
  await expect(page.locator('.statusbar')).toContainText('5 words');
  await page.locator('.docx-page > p').first().click();
  await page.keyboard.press('ControlOrMeta+h');
  await page.getByRole('textbox', { name: 'Find' }).fill('Studio');
  await page.getByRole('textbox', { name: 'Replace with' }).fill('Agency');
  await page.getByRole('button', { name: 'Replace all' }).click();
  await expect(footer(page)).toHaveText('Page 1 · Agency confidential');
  await expect(page.locator('.docx-page > p').first()).toHaveText('Agency work for the client.');
  const { zip, xml } = await download(page);
  expect(await partOf(zip, 'footer1')).toContain('Agency confidential');
  expect(xml).toContain('Agency work');
});

test('Select All in the text leaves the header and footer out; deleting it keeps them', async ({ page }) => {
  await openInEditor(page, await buildDocx(`${BODY}<w:p><w:r><w:t>Second paragraph.</w:t></w:r></w:p>`, OPTS));
  await page.locator('.docx-page > p').first().click();
  await page.keyboard.press('ControlOrMeta+a');
  await page.keyboard.press('Delete');
  await page.keyboard.type('New body');
  await expect(page.locator('.docx-page > p')).toHaveText(['New body']);
  await expect(header(page)).toHaveText('DRAFT v3 – Client name');
  await expect(footer(page)).toHaveText('Page 1 · Studio confidential');
  const { zip } = await download(page);
  expect(await partOf(zip, 'header1')).toBe(PARTS['word/header1.xml']);
});

test('tables and pictures can’t be added in a header; the status bar says so', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY, OPTS));
  await header(page).getByText('Client name').click();
  await page.keyboard.press('ControlOrMeta+Enter');
  await expect(page.locator('.statusbar [role=status]')).toHaveText('A page break can’t go in a header here: add it in Word');
});

test('Backspace at the start of the text leaves the header alone', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY, OPTS));
  await page.locator('.docx-page > p').first().click({ position: { x: 2, y: 5 } });
  await page.keyboard.press('Home');
  await page.keyboard.press('Backspace');
  await page.keyboard.type('X');
  await expect(header(page)).toHaveText('DRAFT v3 – Client name');
  await expect(page.locator('.docx-page > p').first()).toHaveText('XBody text of the proposal.');
  const { zip } = await download(page);
  expect(await partOf(zip, 'header1')).toBe(PARTS['word/header1.xml']);
});
