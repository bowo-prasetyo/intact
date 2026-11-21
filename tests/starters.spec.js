// Starter documents on the start page: a letter, a memo, minutes, a quote, an invoice, a CV, a
// report and a novel's manuscript. Each opens as a new document with Fill in showing its placeholders, is written as Word
// writes a document (its properties in the order Word's schema requires, its parts related and
// typed), is on the paper used where the browser is, and comes back unchanged when opened again.
import fs from 'node:fs';
import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import { downloadFile } from './helpers.js';
import { RPR_ORDER, PPR_ORDER } from '../src/docx/xml.js';

const STARTERS = ['Letter', 'Memo', 'Meeting minutes', 'Quote', 'Invoice', 'CV', 'Report', 'Manuscript'];
const FILES = { Letter: 'Letter', Memo: 'Memo', 'Meeting minutes': 'Minutes', Quote: 'Quote', Invoice: 'Invoice', CV: 'CV', Report: 'Report', Manuscript: 'Manuscript' };
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

test('the manuscript is in standard manuscript format: double-spaced, indented, chapters on new pages, a header after the title page', async ({ page }) => {
  await start(page, 'Manuscript');
  const first = page.locator('.docx-page > p', { hasText: 'The first paragraph of your story' });
  await expect(first).toHaveCSS('font-family', /Times New Roman/);
  await expect.poll(() => first.evaluate((el) => parseFloat(getComputedStyle(el).textIndent))).toBeCloseTo(48, 0); // half an inch
  await expect.poll(() => first.evaluate((el) => parseFloat(getComputedStyle(el).fontSize))).toBeCloseTo(16, 0); // 12 pt
  // The title page's lines are single-spaced and not indented; the scene break is centred.
  await expect(page.locator('.docx-page > p', { hasText: '[Street address]' })).toHaveCSS('text-indent', '0px');
  await expect(page.locator('.docx-page > p', { hasText: /^#$/ })).toHaveCSS('text-align', 'center');
  // The chapters are in the Headings pane.
  await page.getByRole('button', { name: /Headings/ }).click();
  const pane = page.getByRole('navigation', { name: /Headings/ });
  await expect(pane.getByRole('button', { name: 'Chapter One' })).toBeVisible();
  await expect(pane.getByRole('button', { name: 'Chapter Two' })).toBeVisible();
  const dialog = page.getByRole('dialog', { name: /fill in/i });
  for (const name of ['Your legal name', 'Novel title', 'Pen name', 'Surname', 'Short title', 'Word count, rounded']) await expect(dialog.getByRole('textbox', { name, exact: true })).toBeVisible();

  const { read } = await saved(page);
  const styles = await read('word/styles.xml');
  expect(styles).toContain('<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="0" w:line="480" w:lineRule="auto"/><w:ind w:firstLine="720"/></w:pPr></w:style>');
  expect(styles).toContain('<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/>');
  expect(styles.match(/<w:style [^>]*w:styleId="Heading1".*?<\/w:style>/s)[0]).toContain('<w:pageBreakBefore/>');
  const doc = await read('word/document.xml');
  expect(doc).toMatch(/<w:sectPr><w:headerReference w:type="default" r:id="rId5"\/>.*<w:titlePg\/><\/w:sectPr>/s);
  const header = await read('word/header1.xml');
  expect(header).toContain('[Surname] / [Short title] / ');
  expect(header).toContain(' PAGE ');
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

// Indonesian government documents (naskah dinas), offered first when the interface is in Indonesian:
// in Indonesian, on A4 in Arial 12 pt with 3 cm on the left and 2 cm elsewhere, with the letterhead
// (kop) and the signatory's name and NIP, written as Word writes a document.
test.describe('in Indonesian', () => {
  test.use({ locale: 'id-ID' });
  const NASKAH = {
    'Surat dinas': 'Surat dinas', 'Nota dinas': 'Nota dinas', 'Surat undangan': 'Surat undangan', 'Berita acara': 'Berita acara',
    Keputusan: 'Keputusan', 'Surat tugas': 'Surat tugas', 'Surat edaran': 'Surat edaran', 'Notula rapat': 'Notula rapat',
    'Laporan kegiatan': 'Laporan kegiatan', 'Telaahan staf': 'Telaahan staf', 'Surat kuasa': 'Surat kuasa',
  };
  // Those that start with the office's letterhead (kop); a decree, minutes, a report and a staff
  // study start with their title.
  const KOP = ['Surat dinas', 'Nota dinas', 'Surat undangan', 'Berita acara', 'Surat tugas', 'Surat edaran', 'Surat kuasa'];
  // Those whose pages are numbered from the second, at the top in the middle (- 2 -).
  const PAGED = ['Keputusan', 'Surat tugas', 'Surat edaran', 'Notula rapat', 'Laporan kegiatan', 'Telaahan staf'];
  const SECTPR_ORDER = ['headerReference', 'footerReference', 'footnotePr', 'endnotePr', 'type', 'pgSz', 'pgMar', 'paperSrc', 'pgBorders', 'lnNumType', 'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings', 'sectPrChange'];

  test('the start page offers the naskah dinas first, then the others (in English)', async ({ page }) => {
    await page.goto('./');
    const official = page.getByRole('region', { name: 'Naskah dinas (Tata Naskah Dinas)' });
    await expect(official.getByRole('button')).toHaveCount(Object.keys(NASKAH).length);
    for (const name of Object.keys(NASKAH)) await expect(official.getByRole('button', { name: new RegExp(`^${name}:`) })).toBeVisible();
    await expect(official).toContainText('Kertas A4, sesuai ketentuan tata naskah dinas; F4 dan Folio ada di Pengaturan halaman.');
    const others = page.getByRole('region', { name: 'Templat lain (dalam bahasa Inggris)' });
    await expect(others.getByRole('button')).toHaveCount(STARTERS.length);
  });

  /** Downloads it (Unduh, then the review's Download) and reads the saved package. */
  async function savedHere(page) {
    await page.locator('.app-header').getByRole('button', { name: 'Unduh' }).click();
    const review = page.getByRole('dialog', { name: 'Review your changes' });
    const [dl] = await Promise.all([page.waitForEvent('download'), review.getByRole('button', { name: 'Download' }).click()]);
    const zip = await JSZip.loadAsync(fs.readFileSync(await dl.path()));
    return { zip, read: (p) => zip.file(p)?.async('string') };
  }

  for (const [name, file] of Object.entries(NASKAH)) {
    test(`${name}: opens with Fill in, on A4 in Arial with the official margins, in Indonesian, as Word writes a document`, async ({ page }) => {
      await start(page, name);
      await expect(page.getByRole('dialog', { name: 'Isi placeholder' })).toBeVisible();
      await expect(page.getByRole('textbox', { name: 'Nama berkas' })).toHaveValue(file);
      if (KOP.includes(name)) await expect(page.locator('.docx-page > p').first()).toHaveText('[NAMA INSTANSI]');
      const { zip, read } = await savedHere(page);
      const doc = await read('word/document.xml');
      expect(doc).toContain('<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1701"');
      if (KOP.includes(name)) expect(doc).toContain('<w:pBdr><w:bottom w:val="thickThinSmallGap" w:sz="24" w:space="1" w:color="auto"/></w:pBdr>');
      expect(doc).toContain('NIP [NIP');
      if (PAGED.includes(name)) {
        // A page number at the top in the middle, between dashes, from the second page.
        expect(doc).toMatch(/<w:pgNumType w:fmt="numberInDash"[^>]*\/><w:cols [^>]*\/><w:titlePg\/><\/w:sectPr>\s*<\/w:body>/);
        expect(await read('word/header1.xml')).toContain('<w:jc w:val="center"/></w:pPr><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE ');
      }
      expect(outOfOrder(doc, 'sectPr', SECTPR_ORDER)).toEqual([]);
      expect(outOfOrder(doc, 'pPr', PPR_ORDER)).toEqual([]);
      expect(outOfOrder(doc, 'rPr', RPR_ORDER)).toEqual([]);
      expect(outOfOrder(doc, 'tcPr', TCPR_ORDER)).toEqual([]);
      expect(outOfOrder(doc, 'tblPr', TBLPR_ORDER)).toEqual([]);
      const styles = await read('word/styles.xml');
      expect(styles).toContain('<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/><w:sz w:val="24"/>');
      expect(styles).toContain('<w:lang w:val="id-ID"/>');
      for (const [, id] of doc.matchAll(/<w:(?:pStyle|tblStyle|rStyle) w:val="([^"]+)"\/>/g)) expect(styles).toContain(`w:styleId="${id}"`);
      const rels = await read('word/_rels/document.xml.rels');
      const types = await read('[Content_Types].xml');
      for (const [, target] of rels.matchAll(/Target="([^"]+\.xml)"/g)) {
        expect(zip.file(`word/${target}`), target).toBeTruthy();
        expect(types).toContain(`PartName="/word/${target}"`);
      }
    });
  }

  test('filling in the surat dinas: the office\'s name in its letterhead, the labels lined up at a tab stop', async ({ page }) => {
    await start(page, 'Surat dinas');
    const dialog = page.getByRole('dialog', { name: 'Isi placeholder' });
    await expect(dialog.getByText('1 tempat').first()).toBeVisible();
    await dialog.getByRole('textbox', { name: 'NAMA INSTANSI', exact: true }).fill('PEMERINTAH KABUPATEN CONTOH');
    await dialog.getByRole('textbox', { name: 'Nomor surat', exact: true }).fill('800/123/2026');
    await dialog.getByRole('button', { name: 'Isi' }).click();
    await expect(page.locator('.docx-page > p').first()).toHaveText('PEMERINTAH KABUPATEN CONTOH');
    const nomor = page.locator('.docx-page > p', { hasText: '800/123/2026' });
    await expect(nomor).toHaveText(/^Nomor\s*: 800\/123\/2026$/);
    const { read } = await savedHere(page);
    expect(await read('word/document.xml')).toMatch(/<w:tabs><w:tab w:val="left" w:pos="1701"\/><\/w:tabs>.*?Nomor<\/w:t>(?:<\/w:r><w:r>)?<w:tab\/>(?:<\/w:r><w:r>)?<w:t xml:space="preserve">: 800\/123\/2026/s);
  });

  test('the berita acara\'s signatures sit side by side in a table without borders', async ({ page }) => {
    await start(page, 'Berita acara');
    const table = page.locator('.docx-page table');
    await expect(table).toHaveCount(1);
    await expect(table.locator('td').first()).toContainText('PIHAK KEDUA,');
    await expect(table.locator('td').nth(1)).toContainText('PIHAK PERTAMA,');
    await expect(table.locator('td').first()).toHaveCSS('border-top-style', 'none');
  });

  test('the keputusan: its considerations numbered a., b. and 1., 2. as Word numbers them, the official and title typed once and shown in capitals', async ({ page }) => {
    await start(page, 'Keputusan');
    const dialog = page.getByRole('dialog', { name: 'Isi placeholder' });
    // [Jabatan penetap]: the title, the official, Menimbang, Menetapkan, the signature, and the attachment's heading and signature.
    await expect(dialog.getByText('7 tempat')).toBeVisible();
    await dialog.getByRole('textbox', { name: 'Jabatan penetap', exact: true }).fill('Kepala Badan Contoh');
    await dialog.getByRole('textbox', { name: 'Judul keputusan', exact: true }).fill('Pembentukan Tim Penyusun Anggaran');
    await dialog.getByRole('button', { name: 'Isi' }).click();
    // w:caps: typed once in ordinary case, shown in capitals (as the page draws it: innerText).
    await expect.poll(() => page.locator('.docx-page > p').first().innerText()).toBe('KEPUTUSAN KEPALA BADAN CONTOH');
    await expect(page.locator('.docx-page > p').first().locator('strong')).toHaveText(['KEPUTUSAN ', 'Kepala Badan Contoh']); // all of it bold
    const considerations = page.locator('.docx-page table').first();
    await expect(considerations.locator('td').first()).toHaveText('Menimbang');
    await expect(considerations.locator('td').first()).toHaveCSS('border-top-style', 'none');
    const points = considerations.locator('p.list');
    await expect(points).toHaveCount(4);
    expect(await points.evaluateAll((ps) => ps.map((el) => el.dataset.marker))).toEqual(['a.', 'b.', '1.', '2.']);
    await expect(points.nth(1)).toContainText('Keputusan Kepala Badan Contoh tentang Pembentukan Tim Penyusun Anggaran;');
    // The attachment in a section of its own, on a new page, numbered from 1 again.
    await expect(page.locator('.docx-page [data-section-end="nextPage"]')).toHaveCount(1);
    await expect(page.locator('.docx-page p', { hasText: /^LAMPIRAN$/ })).toBeVisible();
    const { read } = await savedHere(page);
    const doc = await read('word/document.xml');
    expect(doc).toContain('<w:sectPr><w:headerReference w:type="default" r:id="rId5"/><w:pgSz w:w="11906" w:h="16838"/>');
    expect(doc).toMatch(/<w:pgNumType w:fmt="numberInDash" w:start="1"\/><w:cols [^>]*\/><w:titlePg\/><\/w:sectPr>\s*<\/w:body>/);
    expect(doc).toContain('<w:caps/>');
    const numbering = await read('word/numbering.xml');
    expect(numbering).toContain('<w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1."/>');
  });

  test('the keputusan prints its attachment\'s pages numbered from 1, neither first page numbered', async ({ page, browserName }) => {
    test.skip(browserName !== 'chromium', 'Page-margin boxes: Chromium');
    await start(page, 'Keputusan');
    await page.keyboard.press('Escape');
    await page.evaluate(() => window.dispatchEvent(new Event('beforeprint')));
    const css = await page.evaluate(() => [...document.head.querySelectorAll('style')].map((s) => s.textContent).find((t) => t.includes('@page')));
    // The decree's first page (the document's) has no number; its others: - 2 -, - 3 -.
    expect(css).toMatch(/@page intact-s0 \{[^\n]*@top-center \{ content: "- " counter\(intact-run0\) " -";/);
    expect(css).toMatch(/@page intact-s0:first \{ counter-increment: intact-run0 1 intact-run1 0; @top-left \{ content: none; \} @top-center \{ content: none; \}/);
    // The attachment's: its first page (its number 1) has none, its others the dashes and number.
    expect(css).toMatch(/@page intact-s1 \{[^\n]*counter-increment: intact-run1;[^\n]*@top-center \{ content: counter\(intact-run1, intact-c\d+\) counter\(intact-run1, intact-c\d+\) counter\(intact-run1, intact-c\d+\);/);
    expect(css).toContain('range: infinite 0, 2 infinite;');
  });

  test('the surat tugas: the employee\'s details lined up, what to do numbered', async ({ page }) => {
    await start(page, 'Surat tugas');
    const table = page.locator('.docx-page table');
    await expect(table.locator('td', { hasText: /^Kepada$/ })).toBeVisible();
    await expect(table.locator('p', { hasText: '[NIP pegawai]' })).toHaveText(/^NIP\s*: \[NIP pegawai\]$/);
    expect(await table.locator('p.list').evaluateAll((ps) => ps.map((el) => el.dataset.marker))).toEqual(['1.', '2.', '1.', '2.']);
  });

  test('the laporan kegiatan: its parts lettered A., B., the introduction\'s numbered under it', async ({ page }) => {
    await start(page, 'Laporan kegiatan');
    const markers = await page.locator('.docx-page p.list').evaluateAll((ps) => ps.map((el) => `${el.dataset.marker} ${el.textContent}`));
    expect(markers).toEqual(['A. Pendahuluan', '1. Umum', '2. Maksud dan Tujuan', '3. Ruang Lingkup', '4. Dasar', 'B. Kegiatan yang Dilaksanakan', 'C. Hasil yang Dicapai', 'D. Kesimpulan dan Saran', 'E. Penutup']);
  });
});
