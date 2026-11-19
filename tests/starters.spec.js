// Starter documents on the start page: a letter, a memo, minutes, a quote, an invoice, a CV and a
// report. Each opens as a new document with Fill in showing its placeholders, is written as Word
// writes a document (its properties in the order Word's schema requires, its parts related and
// typed), is on the paper used where the browser is, and comes back unchanged when opened again.
import fs from 'node:fs';
import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import { downloadFile } from './helpers.js';
import { RPR_ORDER, PPR_ORDER } from '../src/docx/xml.js';

const STARTERS = ['Letter', 'Memo', 'Meeting minutes', 'Quote', 'Invoice', 'CV', 'Report'];
const FILES = { Letter: 'Letter', Memo: 'Memo', 'Meeting minutes': 'Minutes', Quote: 'Quote', Invoice: 'Invoice', CV: 'CV', Report: 'Report' };
const TCPR_ORDER = ['cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd', 'noWrap', 'tcMar', 'textDirection', 'tcFitText', 'vAlign', 'hideMark'];
const TBLPR_ORDER = ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize', 'tblStyleColBandSize', 'tblW', 'jc', 'tblCellSpacing', 'tblInd', 'tblBorders', 'shd', 'tblLayout', 'tblCellMar', 'tblLook'];

async function start(page, name) {
  await page.goto('./');
  await page.getByRole('button', { name: new RegExp(`^${name}:`) }).click();
  await expect(page.locator('.docx-page')).toBeVisible();
}

async function saved(page) {
  const zip = await JSZip.loadAsync(fs.readFileSync(await downloadFile(page)));
  return { zip, read: (p) => zip.file(p)?.async('string') };
}

/** The local names of the w: children of each `<w:tag>…</w:tag>` in `xml` that are out of `order`. */
function outOfOrder(xml, tag, order) {
  const bad = [];
  for (const [, inner] of xml.matchAll(new RegExp(`<w:${tag}>(.*?)</w:${tag}>`, 'gs'))) {
    // Top-level children only: drop nested elements' content first.
    let flat = inner;
    for (let i = 0; i < 4; i++) flat = flat.replace(/<w:(\w+)\b[^>]*>(?:(?!<w:\1\b).)*?<\/w:\1>/gs, '<w:$1/>');
    const names = [...flat.matchAll(/<w:(\w+)\b[^>]*\/>/g)].map((m) => m[1]);
    const at = names.map((n) => order.indexOf(n));
    if (at.some((i) => i < 0) || at.some((i, k) => k && i < at[k - 1])) bad.push(names.join(' '));
  }
  return bad;
}

test('the start page offers the starters, on the paper used where the browser is', async ({ page }) => {
  await page.goto('./');
  const list = page.getByRole('region', { name: 'Or start from a template' });
  await expect(list.getByRole('button')).toHaveCount(STARTERS.length);
  for (const name of STARTERS) await expect(list.getByRole('button', { name: new RegExp(`^${name}:`) })).toBeVisible();
});

for (const name of STARTERS) {
  test(`${name}: opens with Fill in, and is written as Word writes a document`, async ({ page }) => {
    await start(page, name);
    await expect(page.getByRole('dialog', { name: /fill in/i })).toBeVisible();
    await expect(page.getByRole('textbox', { name: 'File name' })).toHaveValue(FILES[name]);
    const { zip, read } = await saved(page);
    const doc = await read('word/document.xml');
    expect(outOfOrder(doc, 'pPr', PPR_ORDER)).toEqual([]);
    expect(outOfOrder(doc, 'rPr', RPR_ORDER)).toEqual([]);
    expect(outOfOrder(doc, 'tcPr', TCPR_ORDER)).toEqual([]);
    expect(outOfOrder(doc, 'tblPr', TBLPR_ORDER)).toEqual([]);
    // Every style the text uses is defined; every part is related and typed.
    const styles = await read('word/styles.xml');
    for (const [, id] of doc.matchAll(/<w:(?:pStyle|tblStyle|rStyle) w:val="([^"]+)"\/>/g)) expect(styles).toContain(`w:styleId="${id}"`);
    const rels = await read('word/_rels/document.xml.rels');
    const types = await read('[Content_Types].xml');
    for (const [, target] of rels.matchAll(/Target="([^"]+\.xml)"/g)) {
      expect(zip.file(`word/${target}`), target).toBeTruthy();
      expect(types).toContain(`PartName="/word/${target}"`);
    }
    if (doc.includes('<w:numPr>')) expect(await read('word/numbering.xml')).toContain('<w:numFmt w:val="bullet"/>');
  });
}

test.describe('in the United States', () => {
  test.use({ locale: 'en-US' });
  test('a starter is on Letter paper, and its tables span it', async ({ page }) => {
    await start(page, 'Quote');
    const { read } = await saved(page);
    const doc = await read('word/document.xml');
    expect(doc).toContain('<w:pgSz w:w="12240" w:h="15840"/>');
    const widths = [...doc.match(/<w:tblGrid>.*?<\/w:tblGrid>/s)[0].matchAll(/w:w="(\d+)"/g)].reduce((a, m) => a + Number(m[1]), 0);
    expect(widths).toBeGreaterThan(9340);
    expect(widths).toBeLessThanOrEqual(9360);
  });
});

test.describe('in the United Kingdom', () => {
  test.use({ locale: 'en-GB' });
  test('a starter is on A4', async ({ page }) => {
    await page.goto('./');
    await expect(page.getByRole('region', { name: 'Or start from a template' })).toContainText('A4 paper, as used where you are');
    await start(page, 'Letter');
    const { read } = await saved(page);
    expect(await read('word/document.xml')).toContain('<w:pgSz w:w="11906" w:h="16838"/>');
  });
});

test('filling in the letter puts each answer everywhere it appears', async ({ page }) => {
  await start(page, 'Letter');
  const dialog = page.getByRole('dialog', { name: /fill in/i });
  await dialog.getByRole('textbox', { name: 'Your name', exact: true }).fill('Ana Putri');
  await dialog.getByRole('textbox', { name: 'Recipient name', exact: true }).fill('Mr Budi');
  await dialog.getByRole('button', { name: 'Fill in' }).click();
  await expect(page.locator('.docx-page')).not.toContainText('[Your name]');
  await expect(page.locator('.docx-page p', { hasText: 'Ana Putri' })).toHaveCount(2);
  await expect(page.locator('.docx-page p', { hasText: 'Dear Mr Budi,' })).toHaveCount(1);
  const { read } = await saved(page);
  const doc = await read('word/document.xml');
  expect(doc.match(/Ana Putri/g)).toHaveLength(2);
  // Made here: its edits aren't tracked, and no "edited by" is needed for a new document.
  expect(doc).not.toContain('<w:ins ');
});

test('the CV’s lists are Word bullets, and its contact details are in the text, not a header', async ({ page }) => {
  await start(page, 'CV');
  await expect(page.locator('.docx-page .list').first()).toHaveAttribute('data-marker', /./);
  await expect(page.locator('.docx-hf-band')).toHaveCount(0);
  // Just below the name (the title).
  await expect(page.locator('.docx-page > :is(p, h1)').nth(1)).toContainText('[Your email]');
});

test('the report numbers its pages in the footer', async ({ page }) => {
  await start(page, 'Report');
  await expect(page.locator('.docx-hf-band.footer .field.page-field')).toHaveCount(1);
  const { read } = await saved(page);
  expect(await read('word/footer1.xml')).toContain('<w:instrText xml:space="preserve"> PAGE   \\* MERGEFORMAT </w:instrText>');
  expect(await read('word/document.xml')).toMatch(/<w:sectPr><w:footerReference w:type="default" r:id="rId4"\/>/);
});

test('a starter downloaded and opened again comes back with the same text', async ({ page }) => {
  await start(page, 'Invoice');
  const text = await page.locator('.docx-page').innerText();
  const path = await downloadFile(page);
  await page.goto('./');
  await page.setInputFiles('input[type=file]', { name: 'Invoice.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: fs.readFileSync(path) });
  await expect(page.locator('.docx-page')).toBeVisible();
  expect(await page.locator('.docx-page').innerText()).toBe(text);
});
