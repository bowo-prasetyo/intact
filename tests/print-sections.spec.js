// Printing (and so saving as PDF) a document of several sections, as Word prints it: each section
// on its own paper (a landscape attachment after a portrait letter), with its own header, footer
// and page numbers (i, ii for the first part, 1, 2 from the attachment on), and a later section's
// different first page (editor/printBands.js). And a page number written without a result, as some
// software writes one: shown as #, kept as it is, numbered when printed.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, W_NS } from './helpers.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const part = (root, inner) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${root} ${W_NS}>${inner}</w:${root}>`;
const para = (text) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const lines = (label, n) => Array.from({ length: n }, (_, i) => para(`${label} line ${i + 1}.`)).join('');
const field = (instr, result) => `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${result}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
const centred = (inner) => `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${inner}</w:p>`;
const A4 = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>';
const A4_LANDSCAPE = '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>';

/**
 * A decree (keputusan) and its attachment (lampiran): a portrait first section and a landscape
 * second, each with a header and a footer of its own. `first`: the sections' different first pages
 * (their first-page parts, with nothing on them); `restart`: the attachment numbers from 1.
 */
function decree({ first = false, restart = true, picture = false, bare = false } = {}) {
  const parts = {};
  let rels = '';
  const ref = (kind, type, id, inner) => {
    parts[`word/${id}.xml`] = part(kind === 'header' ? 'hdr' : 'ftr', inner);
    rels += `<Relationship Id="rId_${id}" Type="${REL}/${kind}" Target="${id}.xml"/>`;
    return `<w:${kind}Reference w:type="${type}" r:id="rId_${id}"/>`;
  };
  const number = bare
    // The whole field in one run, without a result, as some software writes a page number.
    ? '<w:r><w:fldChar w:fldCharType="begin"/><w:instrText xml:space="preserve"> PAGE </w:instrText><w:fldChar w:fldCharType="end"/></w:r>'
    : field('PAGE', '1');
  // The office's logo beside its name, as a letterhead has it (the header's own relationship to it).
  const logo = '<w:r><w:drawing><wp:inline><wp:extent cx="381000" cy="381000"/><wp:docPr id="9" name="Logo"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdLogo"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
  if (picture) parts['word/_rels/header1.xml.rels'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdLogo" Type="${REL}/image" Target="media/image1.png"/></Relationships>`;
  let one = ref('header', 'default', 'header1', picture ? `<w:p>${logo}<w:r><w:t>BADAN CONTOH</w:t></w:r></w:p>` : para('BADAN CONTOH'))
    + ref('footer', 'default', 'footer1', centred(`<w:r><w:t xml:space="preserve">Halaman </w:t></w:r>${number}`));
  let two = ref('header', 'default', 'header2', para('LAMPIRAN KEPUTUSAN NOMOR 12 TAHUN 2026'))
    + ref('footer', 'default', 'footer2', centred(`<w:r><w:t xml:space="preserve">Lampiran hal. </w:t></w:r>${field('PAGE', '1')}`));
  if (first) {
    one += ref('header', 'first', 'header3', '<w:p/>') + ref('footer', 'first', 'footer3', '<w:p/>');
    two += ref('header', 'first', 'header4', '<w:p/>') + ref('footer', 'first', 'footer4', '<w:p/>');
  }
  const titlePg = first ? '<w:titlePg/>' : '';
  const body = lines('Keputusan', 60)
    + `<w:p><w:pPr><w:sectPr>${one}${A4}<w:pgNumType w:fmt="lowerRoman" w:start="1"/>${titlePg}</w:sectPr></w:pPr><w:r><w:t>Ditetapkan di Jakarta.</w:t></w:r></w:p>`
    + lines('Lampiran', 40);
  // The final section's paper is set by buildDocx (portrait A4): the attachment's is landscape.
  const finalSect = `${two}<w:type w:val="nextPage"/>${restart ? '<w:pgNumType w:start="1"/>' : ''}${titlePg}`;
  return buildDocx(body, { parts, rels, sectPr: finalSect }).then((buf) => landscape(buf));
}

/** The final section of a document built by buildDocx, made landscape A4. */
async function landscape(buffer) {
  const JSZip = (await import('jszip')).default;
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file('word/document.xml').async('string');
  zip.file('word/document.xml', xml.replace(/<w:pgSz w:w="11906" w:h="16838"\/><w:pgMar [^>]*\/>(\s*<\/w:sectPr>\s*<\/w:body>)/, `${A4_LANDSCAPE}$1`));
  return zip.generateAsync({ type: 'nodebuffer' });
}

/** The print stylesheet as set just before printing. */
async function beforePrint(page) {
  await page.evaluate(() => window.dispatchEvent(new Event('beforeprint')));
  return page.evaluate(() => [...document.head.querySelectorAll('style')].map((s) => s.textContent).find((t) => t.includes('@page')));
}

/** The PDF's pages' sizes in points, in order: [[width, height]…]. */
const pageSizes = (pdf) => [...pdf.toString('latin1').matchAll(/\/MediaBox\s*\[\s*0 0 ([\d.]+) ([\d.]+)\s*\]/g)].map((m) => [Math.round(Number(m[1])), Math.round(Number(m[2]))]);

async function pdfDialog(page) {
  await page.getByRole('button', { name: 'More ways to save' }).click();
  await page.getByRole('menuitem', { name: /PDF/ }).click();
  return page.getByRole('dialog', { name: 'Save as PDF' });
}

test('each section prints on its own paper: a landscape attachment after a portrait decree', async ({ page, browserName }) => {
  test.skip(browserName !== 'chromium', 'Only Chromium makes PDFs');
  await openInEditor(page, await decree());
  const css = await beforePrint(page);
  // A named page for each section, with its paper; the page's blocks in their section's.
  expect(css).toContain('@page intact-s0 { size: 793.7333333333333px 1122.5333333333333px;');
  expect(css).toContain('@page intact-s1 { size: 1122.5333333333333px 793.7333333333333px;');
  expect(css).toMatch(/:nth-child\(n\+1\):nth-child\(-n\+\d+\) \{ page: intact-s0; \}/);
  expect(css).toMatch(/:nth-child\(n\+\d+\):nth-child\(-n\+\d+\) \{ page: intact-s1; \}/);
  const sizes = pageSizes(await page.pdf({ preferCSSPageSize: true }));
  // The decree's portrait pages, then the attachment's landscape ones, and nothing between.
  const turn = sizes.findIndex(([w, h]) => w > h);
  expect(turn).toBeGreaterThan(0);
  expect(sizes.slice(0, turn).every(([w, h]) => w === 595 && h === 842)).toBe(true);
  expect(sizes.slice(turn).every(([w, h]) => w === 842 && h === 595)).toBe(true);
});

test('each section has its own header, footer and page numbers on every page', async ({ page, browserName }) => {
  test.skip(browserName !== 'chromium', 'Page-margin boxes: Chromium');
  await openInEditor(page, await decree());
  const css = await beforePrint(page);
  // The decree's pages: its header, and i, ii, iii; the attachment's: its own, and 1, 2, 3.
  expect(css).toMatch(/@page intact-s0 \{[^\n]*@top-left \{ content: "BADAN CONTOH";/);
  expect(css).toMatch(/@page intact-s0 \{[^\n]*counter-increment: intact-run0;[^\n]*@bottom-center \{ content: "Halaman " counter\(intact-run0, lower-roman\);/);
  expect(css).toMatch(/@page intact-s1 \{[^\n]*@top-left \{ content: "LAMPIRAN KEPUTUSAN NOMOR 12 TAHUN 2026";/);
  expect(css).toMatch(/@page intact-s1 \{[^\n]*counter-increment: intact-run1;[^\n]*@bottom-center \{ content: "Lampiran hal. " counter\(intact-run1\);/);
  // Each run's start is added on the document's first page, where every counter is.
  expect(css).toContain('@page intact-s0:first { counter-increment: intact-run0 1 intact-run1 0;');
  // The bands drawn on the page aren't printed as well.
  await page.emulateMedia({ media: 'print' });
  await expect(page.locator('.docx-hf-band.header')).toBeHidden();
  await expect(page.locator('.docx-hf-band.footer')).toBeHidden();
  await page.emulateMedia({ media: 'screen' });
  await expect(await pdfDialog(page)).toContainText('with the header and footer on each');
});

test('a section that numbers on from the one before shares its counter', async ({ page, browserName }) => {
  test.skip(browserName !== 'chromium', 'Page-margin boxes: Chromium');
  await openInEditor(page, await decree({ restart: false }));
  const css = await beforePrint(page);
  // One run of numbering: the attachment's pages step the decree's counter on, in its format (1, 2).
  expect(css).toMatch(/@page intact-s1 \{[^\n]*counter-increment: intact-run0;[^\n]*counter\(intact-run0\);/);
  expect(css).toContain('@page intact-s0:first { counter-increment: intact-run0 1;');
});

test('a later section\'s different first page: nothing on it, its number left off, the others numbered', async ({ page, browserName }) => {
  test.skip(browserName !== 'chromium', 'Page-margin boxes: Chromium');
  await openInEditor(page, await decree({ first: true }));
  const css = await beforePrint(page);
  // The decree's first page is the document's: :first, with nothing on it.
  expect(css).toMatch(/@page intact-s0:first \{ counter-increment: intact-run0 1 intact-run1 0; @top-left \{ content: none; \}.*@bottom-center \{ content: none; \}/);
  // The attachment's first page is its page 1: its header and number are shown on its other pages only.
  const s1 = css.split('\n').find((l) => l.startsWith('@page intact-s1 '));
  expect(s1).toMatch(/@top-left \{ content: counter\(intact-run1, (intact-c\d+)\);/);
  const style = (name) => css.split('\n').find((l) => l.startsWith(`@counter-style ${name} `));
  const [, header] = s1.match(/@top-left \{ content: counter\(intact-run1, (intact-c\d+)\);/);
  expect(style(header)).toBe(`@counter-style ${header} { system: cyclic; symbols: "LAMPIRAN KEPUTUSAN NOMOR 12 TAHUN 2026"; range: infinite 0, 2 infinite; fallback: intact-blank; }`);
  const [, number] = s1.match(/@bottom-center \{ content: counter\(intact-run1, intact-c\d+\) counter\(intact-run1, (intact-c\d+)\);/);
  expect(style(number)).toBe(`@counter-style ${number} { system: extends decimal; range: infinite 0, 2 infinite; fallback: intact-blank; }`);
  // The page count is as before, the first page still a page.
  expect(pageSizes(await page.pdf({ preferCSSPageSize: true })).length).toBeGreaterThan(2);
});

test('a later section\'s different first page, numbered on from the one before, prints once and says why', async ({ page, browserName }) => {
  test.skip(browserName !== 'chromium', 'Page-margin boxes: Chromium');
  await openInEditor(page, await decree({ first: true, restart: false }));
  const css = await beforePrint(page);
  // Each section still on its own paper; the header and footer as the page shows them, once.
  expect(css).toContain('@page intact-s1 { size: 1122.5333333333333px 793.7333333333333px;');
  expect(css).not.toContain('@top-left');
  await page.emulateMedia({ media: 'print' });
  await expect(page.locator('.docx-hf-band.header')).toBeVisible();
  await page.emulateMedia({ media: 'screen' });
  const dialog = await pdfDialog(page);
  await expect(dialog).toContainText('The header and footer print once, at the start and the end');
  await expect(dialog).toContainText('Start its page numbers at a number');
});

test('a section\'s header with a picture prints once, each section still on its own paper', async ({ page }) => {
  await openInEditor(page, await decree({ picture: true }));
  const css = await beforePrint(page);
  expect(css).toContain('@page intact-s1 { size: 1122.5333333333333px 793.7333333333333px;');
  expect(css).not.toMatch(/@top-(left|center|right) \{ content: "/);
  const dialog = await pdfDialog(page);
  await expect(dialog).toContainText('The header prints once, at the start');
  await expect(dialog).toContainText('pictures or a table');
});

test('a page number written without a result shows #, prints numbered and is saved as it was', async ({ page, browserName }) => {
  await openInEditor(page, await decree({ bare: true }));
  const bare = page.locator('.docx-hf-band.footer .page-field-bare').first();
  await expect(bare).toHaveAttribute('data-instr', 'PAGE');
  expect(await bare.evaluate((el) => getComputedStyle(el, '::after').content)).toBe('"#"');
  await expect(page.locator('.docx-hf-band.footer')).not.toContainText('⧉');
  if (browserName === 'chromium') {
    const css = await beforePrint(page);
    expect(css).toMatch(/@bottom-center \{ content: "Halaman " counter\(intact-run0, lower-roman\);/);
    await page.evaluate(() => window.dispatchEvent(new Event('afterprint')));
  }
  // An edit elsewhere: the footer's run is written back as it was.
  await page.locator('.docx-page p', { hasText: 'Keputusan line 1.' }).click();
  await page.keyboard.press('End');
  await page.keyboard.type(' Edited.');
  const { zip } = await download(page);
  const footer = await zip.file('word/footer1.xml').async('string');
  expect(footer).toContain('<w:r><w:fldChar w:fldCharType="begin"/><w:instrText xml:space="preserve"> PAGE </w:instrText><w:fldChar w:fldCharType="end"/></w:r>');
});

test('the outline of the table cell being edited isn\'t printed', async ({ page }) => {
  const table = '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>';
  await openInEditor(page, await buildDocx(table));
  const cell = page.locator('td', { hasText: 'Cell' });
  await cell.click();
  await expect(cell).not.toHaveCSS('box-shadow', 'none');
  await page.emulateMedia({ media: 'print' });
  await expect(cell).toHaveCSS('box-shadow', 'none');
});
