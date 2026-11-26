// The healthcare administration starters: a group of their own, last on the start page. Each is a
// controlled document (organisation, ID and version in the header, numbered pages, a document-control
// table, check boxes and a sign-off table), built as the other starters are: on the paper used where
// the browser is, its tables across the text width, written in the order Word's schema requires,
// and with placeholders that each name one thing, so Fill in never puts one answer in two places
// that need two. Nothing in the document names this editor.
import fs from 'node:fs';
import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import { downloadFile } from './helpers.js';
import { RPR_ORDER, PPR_ORDER } from '../src/docx/xml.js';
import { STARTERS as ALL } from '../src/docx/starters.js';
import { PATTERN } from '../src/editor/placeholders.js';

const TEMPLATES = [
  ['Healthcare policy', 'Policy statement'],
  ['Standard operating procedure', 'Procedure'],
  ['Patient administration form', 'Registration'],
  ['Discharge administration checklist', 'Checklist'],
  ['Incident report', 'Factual account'],
  ['Patient complaint record', 'Acknowledgement and handling'],
  ['Healthcare quality audit', 'Findings'],
  ['Corrective action record (CAPA)', 'Cause analysis'],
];
const HEALTH = ALL.filter((s) => s.group === 'health');
const escape = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

async function start(page, title) {
  await page.goto('./');
  const group = page.getByRole('region', { name: 'Healthcare administration' });
  await group.getByRole('button', { name: new RegExp(`^${escape(title)}:`) }).click();
  await expect(page.locator('.docx-page').first()).toBeVisible();
}

async function saved(page) {
  const zip = await JSZip.loadAsync(fs.readFileSync(await downloadFile(page)));
  return { zip, read: (p) => zip.file(p)?.async('string') };
}

test('the start page groups the healthcare administration templates last, with how to check them before sharing', async ({ page }) => {
  await page.goto('./');
  const group = page.getByRole('region', { name: 'Healthcare administration' });
  await expect(group.getByRole('button')).toHaveCount(TEMPLATES.length);
  await expect(group).toContainText('not clinical forms or a compliance certification');
  await expect(group).toContainText('File > Inspect document');
  // After the general templates and the teachers', not before them.
  const titles = await page.locator('.starters h2').allTextContents();
  expect(titles.at(-1)).toBe('Healthcare administration');
  expect(titles[0]).toBe('Or start from a template');
});

test('each placeholder names one thing: no two places that need two answers share one', () => {
  // A placeholder used more than once in a document must mean the same thing each time: the
  // organisation, document ID and version (body and header) are the only ones.
  const SHARED = new Set(['[Organisation]', '[Document ID]', '[Version]']);
  for (const s of HEALTH) {
    const files = s.build('A4');
    const text = `${files.header}${files.body}`.replace(/<[^>]+>/g, '\n');
    const counts = {};
    for (const m of text.matchAll(PATTERN)) counts[m[0]] = (counts[m[0]] || 0) + 1;
    const repeated = Object.entries(counts).filter(([k, n]) => n > 1 && !SHARED.has(k)).map(([k]) => k);
    expect(repeated, s.title).toEqual([]);
  }
});

test('the documents say nothing of this editor, so nothing of it reaches whoever receives them', () => {
  for (const s of HEALTH) {
    const text = s.build('A4').body.replace(/<[^>]+>/g, ' ');
    expect(text, s.title).not.toMatch(/Intact|Inspect document|Before you release|Review →|TEMPLATE/);
  }
});

for (const [title, section] of TEMPLATES) {
  test(`${title}: opens with Fill in, as a controlled document written as Word writes one`, async ({ page }) => {
    await start(page, title);
    await expect(page.getByRole('dialog', { name: /fill in/i })).toBeVisible();
    await expect(page.getByRole('textbox', { name: 'File name' })).toHaveValue(title);
    const sheet = page.locator('.docx-page');
    await expect(sheet).toContainText(section);
    await expect(sheet).toContainText('Before approval or sharing');
    await expect(sheet).toContainText('Sign-off');
    const { zip, read } = await saved(page);
    const doc = await read('word/document.xml');
    // Its organisation, ID and version in the header, its pages numbered in the footer.
    expect(await read('word/header1.xml')).toContain('[Document ID]');
    expect(await read('word/footer1.xml')).toContain(' PAGE ');
    // Check boxes are content controls, as Word's Developer tab inserts them, in Word 2010's namespace.
    expect(doc).toContain('<w14:checkbox>');
    expect(doc).toContain('xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"');
    expect(doc).toContain('mc:Ignorable="w14"');
    const ids = [...doc.matchAll(/<w:sdtPr><w:id w:val="(-?\d+)"\/>/g)].map((m) => m[1]);
    expect(new Set(ids).size).toBe(ids.length);
    // Properties in the order Word's schema requires; every style defined; every part related and typed.
    for (const [tag, order] of [['pPr', PPR_ORDER], ['rPr', RPR_ORDER]]) {
      for (const [, inner] of doc.matchAll(new RegExp(`<w:${tag}>(.*?)</w:${tag}>`, 'gs'))) {
        const names = [...inner.replace(/<w:(\w+)\b[^>]*>(?:(?!<w:\1\b).)*?<\/w:\1>/gs, '<w:$1/>').matchAll(/<w:(\w+)\b[^>]*\/>/g)].map((m) => m[1]);
        const at = names.map((n) => order.indexOf(n));
        expect(at.every((i, k) => i >= 0 && (!k || i >= at[k - 1])), `${tag}: ${names.join(' ')}`).toBe(true);
      }
    }
    const styles = await read('word/styles.xml');
    for (const [, id] of doc.matchAll(/<w:(?:pStyle|tblStyle|rStyle) w:val="([^"]+)"\/>/g)) expect(styles).toContain(`w:styleId="${id}"`);
    const rels = await read('word/_rels/document.xml.rels');
    const types = await read('[Content_Types].xml');
    for (const [, target] of rels.matchAll(/Target="([^"]+\.xml)"/g)) {
      expect(zip.file(`word/${target}`), target).toBeTruthy();
      expect(types).toContain(`PartName="/word/${target}"`);
    }
    // No table is wider than the text: the page's width less its margins.
    const pageWidth = Number(/<w:pgSz w:w="(\d+)"/.exec(doc)[1]);
    const [, right, left] = /<w:pgMar [^>]*w:right="(\d+)"[^>]*w:left="(\d+)"/.exec(doc).map(Number);
    for (const [grid] of doc.matchAll(/<w:tblGrid>.*?<\/w:tblGrid>/gs)) {
      const width = [...grid.matchAll(/w:w="(\d+)"/g)].reduce((a, m) => a + Number(m[1]), 0);
      expect(width).toBeLessThanOrEqual(pageWidth - left - right);
    }
  });
}

test('filling in the dates and names of a CAPA keeps each where it belongs', async ({ page }) => {
  await start(page, 'Corrective action record (CAPA)');
  const dialog = page.getByRole('dialog', { name: /fill in/i });
  await dialog.getByRole('textbox', { name: 'Effective date', exact: true }).fill('2026-12-01');
  await dialog.getByRole('textbox', { name: 'Next review date', exact: true }).fill('2027-12-01');
  await dialog.getByRole('textbox', { name: 'Prepared by', exact: true }).fill('Sari Dewi, Quality lead');
  await dialog.getByRole('textbox', { name: 'Reviewed by', exact: true }).fill('Tomas Reyes, Risk manager');
  await dialog.getByRole('textbox', { name: 'Organisation', exact: true }).fill('Harbour Clinic');
  await dialog.getByRole('button', { name: 'Fill in' }).click();
  const { read } = await saved(page);
  const doc = await read('word/document.xml');
  expect(doc.match(/2026-12-01/g)).toHaveLength(1);
  expect(doc.match(/2027-12-01/g)).toHaveLength(1);
  expect(doc.match(/Sari Dewi/g)).toHaveLength(1);
  expect(doc.match(/Tomas Reyes/g)).toHaveLength(1);
  // The organisation is the same thing in the body and the header, so it is filled in both.
  expect(doc).toContain('Harbour Clinic');
  expect(await read('word/header1.xml')).toContain('Harbour Clinic');
});

test('a check box is ticked with a click, and saved ticked', async ({ page }) => {
  await start(page, 'Incident report');
  await page.getByRole('dialog', { name: /fill in/i }).getByRole('button', { name: 'Cancel' }).click();
  const box = page.locator('.docx-page .sdt-checkbox').first();
  await expect(box).toHaveText('☐');
  await box.click();
  await expect(box).toHaveText('☒');
  const { read } = await saved(page);
  expect(await read('word/document.xml')).toContain('<w14:checked w14:val="1"/>');
});

test.describe('in the United States', () => {
  test.use({ locale: 'en-US' });
  test('a healthcare template is on Letter paper, its tables across it', async ({ page }) => {
    await start(page, 'Standard operating procedure');
    const { read } = await saved(page);
    const doc = await read('word/document.xml');
    expect(doc).toContain('<w:pgSz w:w="12240" w:h="15840"/>');
    const widths = [...doc.matchAll(/<w:tblGrid>.*?<\/w:tblGrid>/gs)].map(([g]) => [...g.matchAll(/w:w="(\d+)"/g)].reduce((a, m) => a + Number(m[1]), 0));
    for (const w of widths) {
      expect(w).toBeGreaterThan(9340);
      expect(w).toBeLessThanOrEqual(9360);
    }
  });
});
