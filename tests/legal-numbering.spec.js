// The Numbering library (Numbering ▾), as Word's Numbering Library and Multilevel List: Indonesia's
// legal numbering (Pasal 1 → (1) → a. → 1.), chapters with Pasal numbered on through them (BAB I →
// Pasal 1, w:lvlRestart 0), and 1. → 1.1. → 1.1.1. Each is a multilevel list written as Word writes
// one; Tab and Shift+Tab move a paragraph between its levels; a Pasal is centred on its own line.
import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import fs from 'node:fs';
import { buildDocx, openInEditor, downloadFile } from './helpers.js';

const para = (text) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const markers = (page) => page.locator('.docx-page .list').evaluateAll((els) => els.map((e) => `${e.dataset.marker}|${e.textContent}`));

async function fromLibrary(page, name) {
  await page.getByRole('button', { name: 'Numbering library' }).click();
  await page.getByRole('menu', { name: 'Numbering library' }).getByRole('menuitem', { name: new RegExp(`^${name.replace(/[.()]/g, '\\$&')}`) }).click();
}

async function selectAll(page) {
  await page.locator('.docx-page p').first().click();
  await page.keyboard.press('ControlOrMeta+a');
}

/** Puts the caret at the start of the paragraph with `text` and presses `key` (Tab, Shift+Tab). */
async function levelOf(page, text, key) {
  await page.locator('.docx-page p', { hasText: text }).click();
  await page.keyboard.press('Home');
  await page.keyboard.press(key);
}

test('Pasal 1 → (1) → a.: levels by Tab, a Pasal centred on its own line, ayat restarting in each Pasal', async ({ page }) => {
  const body = ['', 'Dalam Peraturan ini yang dimaksud dengan:', 'Daerah adalah Kabupaten Contoh.', 'Bupati adalah Bupati Contoh.', '', 'Peraturan ini mulai berlaku pada tanggal diundangkan.'].map((t, i) => para(t || `Pasal-${i}`)).join('');
  await openInEditor(page, await buildDocx(body));
  await selectAll(page);
  await fromLibrary(page, 'Pasal 1');
  await expect(page.locator('.statusbar [role=status]')).toHaveText('Numbered Pasal 1 → (1) → a. → 1.: Tab and Shift+Tab change the level');
  for (const t of ['Dalam Peraturan', 'Peraturan ini mulai']) await levelOf(page, t, 'Tab');
  for (const t of ['Daerah adalah', 'Bupati adalah']) {
    await levelOf(page, t, 'Tab');
    await levelOf(page, t, 'Tab');
  }
  await expect.poll(() => markers(page)).toEqual([
    'Pasal 1|Pasal-0', '(1)|Dalam Peraturan ini yang dimaksud dengan:', 'a.|Daerah adalah Kabupaten Contoh.', 'b.|Bupati adalah Bupati Contoh.',
    'Pasal 2|Pasal-4', '(1)|Peraturan ini mulai berlaku pada tanggal diundangkan.',
  ]);
  // A Pasal is centred, its number in the line.
  const pasal = page.locator('.docx-page p', { hasText: 'Pasal-0' });
  await expect(pasal).toHaveCSS('text-align', 'center');
  await expect(pasal).toHaveAttribute('data-marker-inline', '');
  const zip = await JSZip.loadAsync(fs.readFileSync(await downloadFile(page)));
  const numbering = await zip.file('word/numbering.xml').async('string');
  expect(numbering).toContain('<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:suff w:val="space"/><w:lvlText w:val="Pasal %1"/><w:lvlJc w:val="left"/><w:pPr><w:keepNext/><w:spacing w:before="240"/><w:ind w:left="0" w:firstLine="0"/><w:jc w:val="center"/></w:pPr></w:lvl>');
  expect(numbering).toContain('<w:lvlText w:val="(%2)"/>');
  expect(numbering).toContain('<w:numFmt w:val="lowerLetter"/><w:suff w:val="tab"/><w:lvlText w:val="%3."/>');
  const doc = await zip.file('word/document.xml').async('string');
  expect(doc).toMatch(/<w:numPr><w:ilvl w:val="2"\/><w:numId w:val="\d+"\/><\/w:numPr>.*?Daerah adalah/s);
});

test('BAB I → Pasal 1: Pasal numbers run on through the chapters (w:lvlRestart 0), ayat restart', async ({ page }) => {
  const body = ['Ketentuan Umum', 'P1', 'A1', 'Ketentuan Penutup', 'P2', 'A2'].map(para).join('');
  await openInEditor(page, await buildDocx(body));
  await selectAll(page);
  await fromLibrary(page, 'BAB I');
  for (const t of ['P1', 'P2']) await levelOf(page, t, 'Tab');
  for (const t of ['A1', 'A2']) {
    await levelOf(page, t, 'Tab');
    await levelOf(page, t, 'Tab');
  }
  await expect.poll(() => markers(page)).toEqual(['BAB I|Ketentuan Umum', 'Pasal 1|P1', '(1)|A1', 'BAB II|Ketentuan Penutup', 'Pasal 2|P2', '(1)|A2']);
  const zip = await JSZip.loadAsync(fs.readFileSync(await downloadFile(page)));
  const numbering = await zip.file('word/numbering.xml').async('string');
  expect(numbering).toContain('<w:numFmt w:val="upperRoman"/><w:suff w:val="space"/><w:lvlText w:val="BAB %1"/>');
  expect(numbering).toContain('<w:numFmt w:val="decimal"/><w:lvlRestart w:val="0"/><w:suff w:val="space"/><w:lvlText w:val="Pasal %2"/>');
});

test('1. → 1.1. → 1.1.1., and Word\'s own numbering from the same menu', async ({ page }) => {
  await openInEditor(page, await buildDocx(['One', 'One one', 'One two', 'Two'].map(para).join('')));
  await selectAll(page);
  await fromLibrary(page, '1. → 1.1.');
  for (const t of ['One one', 'One two']) await levelOf(page, t, 'Tab');
  await expect.poll(() => markers(page)).toEqual(['1.|One', '1.1.|One one', '1.2.|One two', '2.|Two']);
  // The same paragraphs as Word's numbering: their levels kept.
  await selectAll(page);
  await fromLibrary(page, '1. → a. → i.');
  await expect.poll(() => markers(page)).toEqual(['1.|One', 'a.|One one', 'b.|One two', '2.|Two']);
});
