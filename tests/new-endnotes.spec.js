// Inserting endnotes, as Word's Insert Endnote (Ctrl+Alt+D): as footnotes are inserted, numbered
// i, ii, iii in the text's order, in Word's Endnote Text and Endnote Reference styles, with an
// endnotes part made for a document without one.
import fs from 'node:fs';
import JSZip from 'jszip';
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, downloadFile } from './helpers.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const BODY = '<w:p><w:r><w:t xml:space="preserve">First claim.</w:t></w:r></w:p><w:p><w:r><w:t xml:space="preserve">Second claim.</w:t></w:r></w:p>';
const notes = (page) => page.locator('.docx-notes[data-notes="endnote"] .docx-note:not([hidden])').evaluateAll((els) => els.map((e) => `${e.dataset.n}:${e.textContent}`));
async function caretAfter(page, text) {
  await page.locator('.docx-page > p', { hasText: text }).click();
  await page.keyboard.press('End');
}

test('Ctrl+Alt+D adds an endnote, numbered i, ii in the text\'s order, saved with its part', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await caretAfter(page, 'Second claim.');
  await page.keyboard.press('ControlOrMeta+Alt+d');
  await page.keyboard.type('Later source.');
  await caretAfter(page, 'First claim.');
  await page.getByRole('button', { name: 'Insert endnote' }).click();
  await page.keyboard.type('Earlier source.');
  await expect.poll(() => notes(page)).toEqual(['i: Earlier source.', 'ii: Later source.']);
  // A footnote alongside goes in its own section, numbered on its own.
  await caretAfter(page, 'Second claim.');
  await page.keyboard.press('ControlOrMeta+Alt+f');
  await page.keyboard.type('A footnote.');
  await expect(page.locator('.docx-notes[data-notes="footnote"] .docx-note')).toHaveText(['A footnote.']);
  await expect(page.locator('.docx-page > .docx-notes').evaluateAll((els) => els.map((e) => e.dataset.notes))).resolves.toEqual(['footnote', 'endnote']);

  const zip = await JSZip.loadAsync(fs.readFileSync(await downloadFile(page)));
  const text = (name) => zip.file(name)?.async('string');
  const doc = await text('word/document.xml');
  expect(doc).toContain('First claim.</w:t></w:r><w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:id="2"/></w:r>');
  const endnotes = await text('word/endnotes.xml');
  expect(endnotes).toContain('<w:endnote w:type="separator" w:id="-1">');
  expect(endnotes).toContain('<w:endnote w:id="2"><w:p><w:pPr><w:pStyle w:val="EndnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteRef/></w:r><w:r><w:t xml:space="preserve"> Earlier source.</w:t></w:r></w:p></w:endnote>');
  expect(await text('word/_rels/document.xml.rels')).toContain(`Type="${REL}/endnotes" Target="endnotes.xml"`);
  expect(await text('[Content_Types].xml')).toContain('<Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>');
  const styles = await text('word/styles.xml');
  expect(styles).toContain('w:styleId="EndnoteText"');
  expect(styles).toContain('<w:style w:type="character" w:styleId="EndnoteReference">');
});

test('Undo takes a new endnote away; nothing is saved for it', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await caretAfter(page, 'First claim.');
  await page.keyboard.press('ControlOrMeta+Alt+d');
  await page.keyboard.press('ControlOrMeta+z');
  await expect(page.locator('.docx-page .note-ref')).toHaveCount(0);
  await caretAfter(page, 'First claim.');
  await page.keyboard.type(' Edited');
  const zip = await JSZip.loadAsync(fs.readFileSync(await downloadFile(page)));
  expect(zip.file('word/endnotes.xml')).toBeNull();
  expect(await zip.file('word/document.xml').async('string')).not.toContain('endnoteReference');
});
