// Templates for teachers, on the start page: in English (a test, a worksheet, a rubric, a lesson plan
// and a letter to parents) and, when the interface is in Indonesian, in Indonesian (lembar soal,
// kisi-kisi, modul ajar, rubrik penilaian, surat untuk orang tua). A test's multiple choice is in two
// columns between continuous section breaks, its questions and choices numbered by Word (1., A.);
// answers are written on ruled lines; rubrics and the kisi-kisi are on landscape pages; the letters
// to parents are ready for mail merge from a class list. Each is written as Word writes a document.
import fs from 'node:fs';
import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import { RPR_ORDER, PPR_ORDER } from '../src/docx/xml.js';

const TCPR_ORDER = ['cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd', 'noWrap', 'tcMar', 'textDirection', 'tcFitText', 'vAlign', 'hideMark'];
const TBLPR_ORDER = ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize', 'tblStyleColBandSize', 'tblW', 'jc', 'tblCellSpacing', 'tblInd', 'tblBorders', 'shd', 'tblLayout', 'tblCellMar', 'tblLook'];
const SECTPR_ORDER = ['headerReference', 'footerReference', 'footnotePr', 'endnotePr', 'type', 'pgSz', 'pgMar', 'paperSrc', 'pgBorders', 'lnNumType', 'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg'];
const sections = (xml) => xml.match(/<w:sectPr\b.*?<\/w:sectPr>/gs);

/** The local names of the w: children of each `<w:tag>…</w:tag>` in `xml` that are out of `order`. */
function outOfOrder(xml, tag, order) {
  const bad = [];
  for (const [, inner] of xml.matchAll(new RegExp(`<w:${tag}>(.*?)</w:${tag}>`, 'gs'))) {
    let flat = inner;
    for (let i = 0; i < 4; i++) flat = flat.replace(/<w:(\w+)\b[^>]*>(?:(?!<w:\1\b).)*?<\/w:\1>/gs, '<w:$1/>');
    const names = [...flat.matchAll(/<w:(\w+)\b[^>]*\/>/g)].map((m) => m[1]);
    const at = names.map((n) => order.indexOf(n));
    if (at.some((i) => i < 0) || at.some((i, k) => k && i < at[k - 1])) bad.push(names.join(' '));
  }
  return bad;
}

async function start(page, name) {
  await page.goto('./');
  await page.getByRole('button', { name: new RegExp(`^${name}:`) }).click();
  await expect(page.locator('.docx-page')).toBeVisible();
}

/** Downloads it (the header's first button, then the review's Download) and reads the saved package. */
async function saved(page) {
  await page.locator('.app-header').getByRole('button', { name: /^(Download|Unduh)$/ }).click();
  const review = page.getByRole('dialog', { name: 'Review your changes' });
  const [dl] = await Promise.all([page.waitForEvent('download'), review.getByRole('button', { name: 'Download' }).click()]);
  const zip = await JSZip.loadAsync(fs.readFileSync(await dl.path()));
  return { zip, read: (p) => zip.file(p)?.async('string') };
}

/** What every starter must be: its parts in Word's order, related and typed, its styles defined. */
async function expectWordDocument({ zip, read }) {
  const doc = await read('word/document.xml');
  expect(outOfOrder(doc, 'pPr', PPR_ORDER)).toEqual([]);
  expect(outOfOrder(doc, 'rPr', RPR_ORDER)).toEqual([]);
  expect(outOfOrder(doc, 'tcPr', TCPR_ORDER)).toEqual([]);
  expect(outOfOrder(doc, 'tblPr', TBLPR_ORDER)).toEqual([]);
  expect(outOfOrder(doc, 'sectPr', SECTPR_ORDER)).toEqual([]);
  const styles = await read('word/styles.xml');
  for (const [, id] of doc.matchAll(/<w:(?:pStyle|tblStyle|rStyle) w:val="([^"]+)"\/>/g)) expect(styles).toContain(`w:styleId="${id}"`);
  const rels = await read('word/_rels/document.xml.rels');
  const types = await read('[Content_Types].xml');
  for (const [, target] of rels.matchAll(/Target="([^"]+\.xml)"/g)) {
    expect(zip.file(`word/${target}`), target).toBeTruthy();
    expect(types).toContain(`PartName="/word/${target}"`);
  }
  // Every list the text uses is defined.
  const numbering = await read('word/numbering.xml');
  for (const [, id] of doc.matchAll(/<w:numId w:val="(\d+)"\/>/g)) expect(numbering).toContain(`<w:num w:numId="${id}">`);
  return doc;
}

const ENGLISH = { Test: 'Test', Worksheet: 'Worksheet', Rubric: 'Rubric', 'Lesson plan': 'Lesson plan', 'Letter to parents': 'Letter to parents' };

test.describe('in English', () => {
  test.use({ locale: 'en-GB' });

  test('the start page offers the teachers’ templates as a group of their own', async ({ page }) => {
    await page.goto('./');
    const group = page.getByRole('region', { name: 'For teachers' });
    await expect(group.getByRole('button')).toHaveCount(Object.keys(ENGLISH).length);
    for (const name of Object.keys(ENGLISH)) await expect(group.getByRole('button', { name: new RegExp(`^${name}:`) })).toBeVisible();
    // The everyday templates are as they were, without them.
    await expect(page.getByRole('region', { name: 'Or start from a template' }).getByRole('button')).toHaveCount(8);
  });

  for (const [name, file] of Object.entries(ENGLISH)) {
    test(`${name}: opens with Fill in, and is written as Word writes a document`, async ({ page }) => {
      await start(page, name);
      await expect(page.getByRole('dialog', { name: /fill in/i })).toBeVisible();
      await expect(page.getByRole('textbox', { name: 'File name' })).toHaveValue(file);
      await page.getByRole('dialog', { name: /fill in/i }).getByRole('button', { name: 'Cancel' }).click();
      await expectWordDocument(await saved(page));
    });
  }

  test('the test: multiple choice in two columns on the same page, numbered 1. and A. by Word, answers on lines', async ({ page }) => {
    await start(page, 'Test');
    await page.getByRole('dialog', { name: /fill in/i }).getByRole('button', { name: 'Cancel' }).click();
    const columns = page.locator('.docx-page > .docx-columns');
    await expect(columns).toHaveCount(1);
    await expect(columns).toHaveAttribute('data-cols', '2');
    await expect(columns).toHaveCSS('column-rule-style', 'solid');
    const question = columns.locator('p', { hasText: '[Question 1]' });
    await expect(question).toHaveAttribute('data-marker', '1.');
    await expect(columns.locator('p', { hasText: 'Answer C' }).first()).toHaveAttribute('data-marker', 'C.');
    await expect(columns.locator('p', { hasText: '[Question 4]' })).toHaveAttribute('data-marker', '4.');
    // Part B goes on numbering after Part A, out of the columns.
    await expect(page.locator('.docx-page > p', { hasText: '[Question 5]' })).toHaveAttribute('data-marker', '5.');
    // No page break between the parts: the breaks are continuous.
    await expect(page.locator('[data-section-end]')).toHaveCount(2);
    for (const el of await page.locator('[data-section-end]').all()) await expect(el).toHaveAttribute('data-section-end', 'continuous');
    // The answer lines: a line under each row, no other borders.
    const line = page.locator('.docx-page > p:has-text("[Question 5]") + table td').first();
    await expect(line).toHaveCSS('border-bottom-style', 'solid');
    await expect(line).toHaveCSS('border-left-style', 'none');
    // A choice added with Enter is the next letter.
    await columns.locator('p', { hasText: 'Answer D' }).first().click();
    await page.keyboard.press('End');
    await page.keyboard.press('Enter');
    await page.keyboard.type('Answer E');
    await expect(columns.locator('p', { hasText: 'Answer E' })).toHaveAttribute('data-marker', 'E.');

    const { read } = await saved(page);
    const doc = await read('word/document.xml');
    const [first, choices, last] = sections(doc);
    expect(first).not.toContain('<w:type');
    expect(choices).toContain('<w:type w:val="continuous"/>');
    expect(choices).toContain('<w:cols w:num="2" w:space="708" w:sep="1"/>');
    expect(last).toContain('<w:type w:val="continuous"/>');
    expect(last).toContain('<w:cols w:space="708"/>');
    const numbering = await read('word/numbering.xml');
    expect(numbering).toContain('<w:numFmt w:val="upperLetter"/><w:lvlText w:val="%2."/>');
  });

  test('the rubric is on a landscape page, with room to write in its cells', async ({ page }) => {
    await start(page, 'Rubric');
    await expect(page.locator('.docx-sheet')).toHaveCSS('width', /^1122\.5/); // A4 landscape (297 mm)
    const row = page.locator('.docx-page table tr', { hasText: '[First criterion]' });
    await expect.poll(async () => (await row.boundingBox()).height).toBeGreaterThan(70); // 2 cm
    const { read } = await saved(page);
    expect(await read('word/document.xml')).toContain('<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>');
  });

  test('the letter to parents is merged from a class list: a letter for each student', async ({ page }) => {
    await start(page, 'Letter to parents');
    const fill = page.getByRole('dialog', { name: /fill in/i });
    for (const name of ['Parent name', 'Student name', 'Class']) await expect(fill.getByRole('textbox', { name, exact: true })).toBeVisible();
    await fill.getByRole('button', { name: 'From a list…' }).click();
    const merge = page.getByRole('dialog', { name: 'Mail merge' });
    await merge.locator('input[type=file]').setInputFiles({ name: 'class.csv', mimeType: 'text/csv', buffer: Buffer.from('Student name,Parent name,Class\r\nAna Putri,Mrs Putri,5B\r\nBudi Santoso,Mr Santoso,5B\r\n') });
    await expect(merge.getByRole('combobox', { name: 'Column for [Parent name]' })).toHaveValue('1');
    await expect(merge.getByRole('combobox', { name: 'Column for [Student name]' })).toHaveValue('0');
    await expect(merge.getByRole('combobox', { name: 'Column for [Class]' })).toHaveValue('2');
    const [dl] = await Promise.all([page.waitForEvent('download'), merge.getByRole('button', { name: 'Merge 2 rows' }).click()]);
    const xml = await (await JSZip.loadAsync(fs.readFileSync(await dl.path()))).file('word/document.xml').async('string');
    expect(xml).toContain('Dear Mrs Putri,');
    expect(xml).toContain('Dear Mr Santoso,');
    expect(xml).not.toContain('[Student name]');
  });
});

const INDONESIAN = { 'Lembar soal': 'Lembar soal', 'Kisi-kisi soal': 'Kisi-kisi soal', 'Modul ajar': 'Modul ajar', 'Rubrik penilaian': 'Rubrik penilaian', 'Surat untuk orang tua': 'Surat untuk orang tua' };

test.describe('in Indonesian', () => {
  test.use({ locale: 'id-ID' });

  test('the start page offers them after the naskah dinas, in Indonesian', async ({ page }) => {
    await page.goto('./');
    const regions = page.getByRole('region');
    await expect(regions.nth(1)).toHaveAccessibleName('Untuk guru (Kurikulum Merdeka)');
    const group = page.getByRole('region', { name: 'Untuk guru (Kurikulum Merdeka)' });
    await expect(group.getByRole('button')).toHaveCount(Object.keys(INDONESIAN).length);
    for (const name of Object.keys(INDONESIAN)) await expect(group.getByRole('button', { name: new RegExp(`^${name}:`) })).toBeVisible();
    await expect(page.getByRole('region', { name: 'Naskah dinas (Tata Naskah Dinas)' }).getByRole('button')).toHaveCount(11);
    await expect(page.getByRole('region', { name: 'Templat lain (dalam bahasa Inggris)' }).getByRole('button')).toHaveCount(8);
  });

  for (const [name, file] of Object.entries(INDONESIAN)) {
    test(`${name}: opens with Fill in, in Indonesian on A4, written as Word writes a document`, async ({ page }) => {
      await start(page, name);
      await expect(page.getByRole('dialog', { name: 'Isi placeholder' })).toBeVisible();
      await expect(page.getByRole('textbox', { name: 'Nama berkas' })).toHaveValue(file);
      await page.getByRole('dialog', { name: 'Isi placeholder' }).getByRole('button', { name: 'Batal' }).click();
      const pkg = await saved(page);
      const doc = await expectWordDocument(pkg);
      expect(doc).toMatch(/<w:pgSz w:w="(11906|16838)" w:h="(16838|11906)"/);
      expect(await pkg.read('word/styles.xml')).toContain('<w:lang w:val="id-ID"/>');
    });
  }

  test('the lembar soal: the letterhead, multiple choice in two columns, each part numbered from 1', async ({ page }) => {
    await start(page, 'Lembar soal');
    await page.getByRole('dialog', { name: 'Isi placeholder' }).getByRole('button', { name: 'Batal' }).click();
    await expect(page.locator('.docx-page > p').first()).toHaveText('[NAMA SEKOLAH]');
    await expect(page.locator('.docx-page > p').first()).toHaveCSS('font-family', /Times New Roman/);
    const columns = page.locator('.docx-page > .docx-columns');
    await expect(columns).toHaveAttribute('data-cols', '2');
    await expect(columns.locator('p', { hasText: '[Soal pilihan ganda 4]' })).toHaveAttribute('data-marker', '4.');
    await expect(columns.locator('p', { hasText: 'Jawaban B' }).first()).toHaveAttribute('data-marker', 'B.');
    await expect(page.locator('.docx-page > p', { hasText: '[Soal isian 1]' })).toHaveAttribute('data-marker', '1.');
    await expect(page.locator('.docx-page > p', { hasText: '[Soal uraian 2]' })).toHaveAttribute('data-marker', '2.');
    await expect(page.locator('p', { hasText: 'Berilah tanda silang' })).toHaveAttribute('data-section-note', 'Bagian 1: A4 potret · pemisah bagian (berkelanjutan)');
    const { read } = await saved(page);
    const doc = await read('word/document.xml');
    expect(sections(doc)).toHaveLength(3);
    expect(doc).toContain('<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"');
    const numbering = await read('word/numbering.xml');
    // Isian and uraian start again at 1.
    expect(numbering.match(/<w:startOverride w:val="1"\/>/g)).toHaveLength(2);
  });

  test('the kisi-kisi is a landscape table, signed by the head of school and the teacher', async ({ page }) => {
    await start(page, 'Kisi-kisi soal');
    await expect(page.locator('.docx-sheet')).toHaveCSS('width', /^1122\.5/);
    await expect(page.locator('.docx-page table').first().locator('tr').first()).toContainText('Indikator soal');
    await expect(page.locator('.docx-page table').last()).toContainText('Kepala [Nama sekolah]');
  });
});
