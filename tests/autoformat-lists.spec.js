// Automatic lists, as Word's AutoFormat As You Type: "1. ", "a) ", "(i) ", "* " or "- " typed at the
// start of a paragraph makes it a list item of that kind; Ctrl+Z gives the text back as typed. The
// lists are written as Word writes them (their own numbering definition, starting where typed), and
// the AutoFormat switch in the status bar turns this off with the smart quotes.
import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import fs from 'node:fs';
import { buildDocx, openInEditor, download, downloadFile } from './helpers.js';

const DOC = '<w:p><w:r><w:t>Intro</w:t></w:r></w:p>';
const markers = (page) => page.locator('.docx-page .list').evaluateAll((els) => els.map((e) => `${e.dataset.marker} ${e.textContent}`));
const savedParagraph = (xml, text) => xml.match(/<w:p[ >].*?<\/w:p>/gs).find((p) => p.includes(`>${text}<`));

/** A new paragraph after "Intro", with the caret in it. */
async function newParagraph(page) {
  await page.locator('.docx-page p', { hasText: 'Intro' }).click();
  await page.keyboard.press('End');
  await page.keyboard.press('Enter');
}

test('"1. " starts a numbered list; Enter continues it and Enter on an empty item ends it', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await newParagraph(page);
  await page.keyboard.type('1. First');
  await page.keyboard.press('Enter');
  await page.keyboard.type('Second');
  await expect.poll(() => markers(page)).toEqual(['1. First', '2. Second']);
  await page.keyboard.press('Enter');
  await page.keyboard.press('Enter');
  await page.keyboard.type('After');
  await expect.poll(() => markers(page)).toEqual(['1. First', '2. Second']);

  const path = await downloadFile(page);
  const zip = await JSZip.loadAsync(fs.readFileSync(path));
  const xml = await zip.file('word/document.xml').async('string');
  const numbering = await zip.file('word/numbering.xml').async('string');
  const numId = savedParagraph(xml, 'First').match(/<w:numId w:val="(\d+)"\/>/)[1];
  expect(savedParagraph(xml, 'First')).toBe(`<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t xml:space="preserve">First</w:t></w:r></w:p>`);
  expect(savedParagraph(xml, 'Second')).toContain(`<w:numId w:val="${numId}"/>`);
  expect(savedParagraph(xml, 'After')).toBe('<w:p><w:r><w:t xml:space="preserve">After</w:t></w:r></w:p>');
  const num = numbering.match(new RegExp(`<w:num w:numId="${numId}">.*?</w:num>`, 's'))[0];
  const abstractId = num.match(/<w:abstractNumId w:val="(\d+)"\/>/)[1];
  expect(num).toContain('<w:startOverride w:val="1"/>');
  const level0 = numbering.match(new RegExp(`<w:abstractNum w:abstractNumId="${abstractId}">.*?<w:lvl w:ilvl="0">(.*?)</w:lvl>`, 's'))[1];
  expect(level0).toContain('<w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>');
});

test('letters, Roman numerals and brackets give lists in that style, from the number typed', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await newParagraph(page);
  await page.keyboard.type('c) Third letter');
  await page.keyboard.press('Enter');
  await page.keyboard.press('Enter');
  await page.keyboard.type('(i) Roman');
  await page.keyboard.press('Enter');
  await page.keyboard.type('Next');
  await page.keyboard.press('Enter');
  await page.keyboard.press('Enter');
  await page.keyboard.type('A. Capital');
  await page.keyboard.press('Enter');
  await page.keyboard.press('Enter');
  await page.keyboard.type('7) Seven');
  await expect.poll(() => markers(page)).toEqual(['c) Third letter', '(i) Roman', '(ii) Next', 'A. Capital', '7) Seven']);

  const path = await downloadFile(page);
  const numbering = await (await JSZip.loadAsync(fs.readFileSync(path))).file('word/numbering.xml').async('string');
  expect(numbering).toContain('<w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1)"/>');
  expect(numbering).toContain('<w:numFmt w:val="lowerRoman"/><w:lvlText w:val="(%1)"/>');
  expect(numbering).toContain('<w:numFmt w:val="upperLetter"/><w:lvlText w:val="%1."/>');
  expect(numbering).toContain('<w:numFmt w:val="decimal"/><w:lvlText w:val="%1)"/>');
  expect(numbering).toContain('<w:startOverride w:val="3"/>');
  expect(numbering).toContain('<w:startOverride w:val="7"/>');
});

test('"* " and "- " start bulleted lists: Word\'s bullet, and a dash', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await newParagraph(page);
  await page.keyboard.type('* Star');
  await page.keyboard.press('Enter');
  await page.keyboard.press('Enter');
  await page.keyboard.type('- Dash');
  await expect.poll(() => markers(page)).toEqual(['• Star', '- Dash']);
  const path = await downloadFile(page);
  const numbering = await (await JSZip.loadAsync(fs.readFileSync(path))).file('word/numbering.xml').async('string');
  expect(numbering).toContain('<w:numFmt w:val="bullet"/><w:lvlText w:val="-"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>');
});

test('"2. " under a numbered list continues it', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await newParagraph(page);
  await page.keyboard.type('1. One');
  await page.keyboard.press('Enter');
  await page.keyboard.press('Backspace'); // out of the list, keeping the paragraph
  await page.keyboard.type('2. Two');
  await expect.poll(() => markers(page)).toEqual(['1. One', '2. Two']);
  const { xml } = await download(page);
  const id = (text) => savedParagraph(xml, text).match(/<w:numId w:val="(\d+)"\/>/)[1];
  expect(id('Two')).toBe(id('One'));
});

test('Ctrl+Z after the list is made gives back the text as typed', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await newParagraph(page);
  await page.keyboard.type('1. ');
  await expect.poll(() => markers(page)).toEqual(['1. ']);
  await page.keyboard.press('ControlOrMeta+z');
  await expect.poll(() => markers(page)).toEqual([]);
  await expect(page.locator('.docx-page p').nth(1)).toHaveText('1. ');
  await page.keyboard.type('stays as typed');
  const { xml } = await download(page);
  expect(xml).toContain('<w:p><w:r><w:t xml:space="preserve">1. stays as typed</w:t></w:r></w:p>');
});

test('text that only looks like a list number, or typed mid-paragraph, stays as it is', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await newParagraph(page);
  await page.keyboard.type('Version 1. is out; 2026. was good; ab) no; (1. no');
  await page.keyboard.press('Enter');
  await page.keyboard.type('1.5 is a number');
  await expect.poll(() => markers(page)).toEqual([]);
});

test('with AutoFormat off, "1. " is typed as it is', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await page.getByRole('button', { name: /^AutoFormat: on$/ }).click();
  await newParagraph(page);
  await page.keyboard.type('1. Plain');
  await expect.poll(() => markers(page)).toEqual([]);
  await expect(page.locator('.docx-page p').nth(1)).toHaveText('1. Plain');
});

test('while tracking, the list is a tracked paragraph change and the typed number is not kept', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await page.getByRole('button', { name: 'Track changes' }).click();
  const dialog = page.getByRole('dialog', { name: 'Your name on tracked changes' });
  await dialog.getByRole('textbox').fill('Rita');
  await dialog.getByRole('button', { name: 'Start tracking' }).click();
  await newParagraph(page);
  await page.keyboard.type('1. Tracked item');
  await expect.poll(() => markers(page)).toEqual(['1. Tracked item']);
  const { xml } = await download(page);
  const p = savedParagraph(xml, 'Tracked item');
  expect(p).toContain('<w:numPr><w:ilvl w:val="0"/>');
  expect(p).toMatch(/<w:pPrChange w:id="\d+" w:author="Rita" w:date="[^"]+"><w:pPr\/><\/w:pPrChange>/);
  expect(p).toMatch(/<w:ins w:id="\d+" w:author="Rita" w:date="[^"]+"><w:r><w:t xml:space="preserve">Tracked item<\/w:t><\/w:r><\/w:ins>/);
  expect(xml).not.toContain('>1. <');
  expect(xml).not.toContain('<w:delText');
});
