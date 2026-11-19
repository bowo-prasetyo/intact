// Inserting footnotes, as Word's Insert Footnote (Ctrl+Alt+F): a reference at the caret and its note
// below the text, numbered in the text's order, in Word's Footnote Text and Footnote Reference
// styles. A document without footnotes gets the part for them; a note is saved only while its
// reference is in the text, so Undo or deleting the reference takes it away.
import fs from 'node:fs';
import JSZip from 'jszip';
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, downloadFile, W_NS } from './helpers.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const mark = (local) => `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:${local}/></w:r>`;
const FOOTNOTES = `${XML}<w:footnotes ${W_NS}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>`
  + `<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>`
  + `<w:footnote w:id="1"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>${mark('footnoteRef')}<w:r><w:t xml:space="preserve"> Existing note.</w:t></w:r></w:p></w:footnote></w:footnotes>`;
const WITH_NOTES = {
  parts: { 'word/footnotes.xml': FOOTNOTES },
  rels: `<Relationship Id="rIdF" Type="${REL}/footnotes" Target="footnotes.xml"/>`,
};
const BODY = '<w:p><w:r><w:t xml:space="preserve">First claim.</w:t></w:r></w:p><w:p><w:r><w:t xml:space="preserve">Second claim.</w:t></w:r></w:p>';
const notes = (page) => page.locator('.docx-notes[data-notes="footnote"] .docx-note:not([hidden])').evaluateAll((els) => els.map((e) => `${e.dataset.n}:${e.textContent}`));

async function saved(page) {
  const zip = await JSZip.loadAsync(fs.readFileSync(await downloadFile(page)));
  return (name) => zip.file(name)?.async('string') ?? null;
}
async function caretAfter(page, text) {
  await page.locator('.docx-page > p', { hasText: text }).click();
  await page.keyboard.press('End');
}

test('Ctrl+Alt+F adds a footnote: a reference at the caret, its note below with the caret in it', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await caretAfter(page, 'Second claim.');
  await page.keyboard.press('ControlOrMeta+Alt+f');
  await page.keyboard.type('Second source.');
  await caretAfter(page, 'First claim.');
  await page.getByRole('button', { name: 'Insert footnote' }).click();
  await page.keyboard.type('First source.');
  // Numbered, and listed, in the text's order.
  await expect.poll(() => notes(page)).toEqual(['1: First source.', '2: Second source.']);
  await expect(page.locator('.docx-page > p .note-ref')).toHaveText(['', '']);
  expect(await page.locator('.docx-page > p .note-ref').evaluateAll((els) => els.map((e) => e.dataset.n))).toEqual(['1', '2']);

  const text = await saved(page);
  const doc = await text('word/document.xml');
  const id = (claim) => doc.match(new RegExp(`${claim}</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="(\\d+)"/></w:r>`))[1];
  const footnotes = await text('word/footnotes.xml');
  expect(footnotes).toContain('<w:footnote w:type="separator" w:id="-1">');
  expect(footnotes).toContain(`<w:footnote w:id="${id('First claim.')}"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>${mark('footnoteRef')}<w:r><w:t xml:space="preserve"> First source.</w:t></w:r></w:p></w:footnote>`);
  expect(footnotes).toContain(`<w:footnote w:id="${id('Second claim.')}">`);
  expect(await text('word/_rels/document.xml.rels')).toContain(`Type="${REL}/footnotes" Target="footnotes.xml"`);
  expect(await text('[Content_Types].xml')).toContain('<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>');
  const styles = await text('word/styles.xml');
  expect(styles).toContain('w:styleId="FootnoteText"');
  expect(styles).toContain('<w:style w:type="character" w:styleId="FootnoteReference">');
});

test('a new footnote joins the document\'s own, numbered among them', async ({ page }) => {
  const body = `<w:p><w:r><w:t xml:space="preserve">Known.</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r></w:p><w:p><w:r><w:t xml:space="preserve">Before it.</w:t></w:r></w:p>`;
  await openInEditor(page, await buildDocx(`<w:p><w:r><w:t xml:space="preserve">Opening.</w:t></w:r></w:p>${body}`, WITH_NOTES));
  await caretAfter(page, 'Opening.');
  await page.keyboard.press('ControlOrMeta+Alt+f');
  await page.keyboard.type('New one.');
  await expect.poll(() => notes(page)).toEqual(['1: New one.', '2: Existing note.']);
  const text = await saved(page);
  const footnotes = await text('word/footnotes.xml');
  expect(footnotes).toContain(`<w:footnote w:id="1"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>${mark('footnoteRef')}<w:r><w:t xml:space="preserve"> Existing note.</w:t></w:r></w:p></w:footnote>`);
  expect(footnotes).toContain('<w:footnote w:id="2"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>');
  expect(footnotes).toContain(' New one.</w:t>');
  expect(await text('word/document.xml')).toContain('Opening.</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r>');
});

test('Undo takes a new footnote away; with its reference deleted, its note isn\'t saved', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await caretAfter(page, 'First claim.');
  await page.keyboard.press('ControlOrMeta+Alt+f');
  await page.keyboard.type('Gone.');
  await page.keyboard.press('ControlOrMeta+z'); // the typing
  await page.keyboard.press('ControlOrMeta+z'); // the footnote
  await expect(page.locator('.docx-page .note-ref')).toHaveCount(0);
  await caretAfter(page, 'First claim.');
  await page.keyboard.press('ControlOrMeta+Alt+f');
  await page.keyboard.type('Also gone.');
  await page.locator('.docx-note .note-mark').click(); // goes beside its reference in the text
  await page.keyboard.press('Backspace'); // deletes the reference
  await expect(page.locator('.docx-page > p .note-ref')).toHaveCount(0);
  await expect.poll(() => notes(page)).toEqual([]);
  await page.keyboard.type(' Edited');
  const text = await saved(page);
  expect(await text('word/footnotes.xml')).toBeNull();
  expect(await text('word/document.xml')).not.toContain('footnoteReference');
  expect(await text('word/_rels/document.xml.rels')).not.toContain('footnotes');
});

test('footnotes go in the text: not in a note, header or footer', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY, WITH_NOTES));
  await caretAfter(page, 'First claim.');
  await page.keyboard.press('ControlOrMeta+Alt+f');
  await page.keyboard.type('In a note');
  await page.keyboard.press('ControlOrMeta+Alt+f');
  await expect(page.locator('.statusbar [role=status]')).toHaveText('A footnote can’t go in a footnote here: add it in Word');
});

test('while tracking, the reference is a tracked insertion', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await page.getByRole('button', { name: 'Track changes' }).click();
  const dialog = page.getByRole('dialog', { name: 'Your name on tracked changes' });
  await dialog.getByRole('textbox').fill('Rita');
  await dialog.getByRole('button', { name: 'Start tracking' }).click();
  await caretAfter(page, 'First claim.');
  await page.keyboard.press('ControlOrMeta+Alt+f');
  await page.keyboard.type('Tracked note.');
  const text = await saved(page);
  expect(await text('word/document.xml')).toMatch(/<w:ins w:id="\d+" w:author="Rita" w:date="[^"]+"><w:r><w:rPr><w:rStyle w:val="FootnoteReference"\/><\/w:rPr><w:footnoteReference w:id="1"\/><\/w:r><\/w:ins>/);
  expect(await text('word/footnotes.xml')).toMatch(/<w:ins w:id="\d+" w:author="Rita" w:date="[^"]+"><w:r><w:t xml:space="preserve">Tracked note\.<\/w:t><\/w:r><\/w:ins>/);
});

test('End goes past a footnote reference at the end of the line, so typing carries on after it', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await caretAfter(page, 'First claim.');
  await page.keyboard.press('ControlOrMeta+Alt+f');
  await page.keyboard.type('A source.');
  const line = page.locator('.docx-page > p', { hasText: 'First claim.' });
  await line.click({ position: { x: 2, y: 4 } });
  await page.keyboard.press('End');
  await page.keyboard.type(' More.');
  // As in Word: the reference stays on "claim.", the new text follows it.
  expect(await line.evaluate((p) => [...p.childNodes].map((n) => (n.nodeType === 3 ? n.data : n.classList.contains('note-ref') ? '¹' : n.textContent)).join(''))).toBe('First claim.¹ More.');
  await page.keyboard.press('Home');
  await page.keyboard.press('Shift+End');
  await expect.poll(() => page.evaluate(() => getSelection().toString().trim())).toBe('First claim. More.');
  const text = await saved(page);
  expect(await text('word/document.xml')).toMatch(/First claim\.<\/w:t><\/w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"\/><\/w:rPr><w:footnoteReference w:id="\d+"\/><\/w:r><w:r><w:t xml:space="preserve"> More\.<\/w:t><\/w:r><\/w:p>/);
});
