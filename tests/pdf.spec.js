// Saving as PDF, and printing with headers and footers on every page (editor/printBands.js):
// Download's menu offers a PDF, made by the browser from its print window; a short dialog says how
// and what the PDF will show. Just before printing, a header or footer that is the same on every
// page is set to repeat (a fixed box), and page numbers and different first or even pages go in the
// page's margin boxes, numbered per page.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, W_NS } from './helpers.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const part = (root, inner) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${root} ${W_NS}>${inner}</w:${root}>`;
const para = (text) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const LONG = Array.from({ length: 70 }, (_, i) => para(`Line ${i + 1} of the CV.`)).join('');
const field = (instr, result) => `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${result}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;

function withBands({ header, footer, firstHeader, firstFooter, sections = false } = {}) {
  const parts = {};
  let rels = '';
  let sectPr = '';
  const add = (kind, type, id, inner) => {
    parts[`word/${id}.xml`] = part(kind === 'header' ? 'hdr' : 'ftr', inner);
    rels += `<Relationship Id="rId_${id}" Type="${REL}/${kind}" Target="${id}.xml"/>`;
    sectPr += `<w:${kind}Reference w:type="${type}" r:id="rId_${id}"/>`;
  };
  if (header) add('header', 'default', 'header1', header);
  if (firstHeader) add('header', 'first', 'header2', firstHeader);
  if (footer) add('footer', 'default', 'footer1', footer);
  if (firstFooter) add('footer', 'first', 'footer2', firstFooter);
  if (firstHeader || firstFooter) sectPr += '<w:titlePg/>';
  let body = LONG;
  if (sections) {
    // A first section with a header of its own, then the last.
    parts['word/header3.xml'] = part('hdr', para('First section header'));
    rels += `<Relationship Id="rId_header3" Type="${REL}/header" Target="header3.xml"/>`;
    body = para('Section one') + `<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rId_header3"/><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr><w:r><w:t>End</w:t></w:r></w:p>` + LONG;
  }
  return buildDocx(body, { parts, rels, sectPr });
}

const pdfPages = (pdf) => (pdf.toString('latin1').match(/\/Type\s*\/Page[^s]/g) || []).length;
/** The print stylesheet as set just before printing. */
async function beforePrint(page) {
  await page.evaluate(() => window.dispatchEvent(new Event('beforeprint')));
  return page.evaluate(() => [...document.head.querySelectorAll('style')].map((s) => s.textContent).find((t) => t.includes('@page')));
}

test('Download\'s menu saves a PDF from the print window, saying how and what it shows', async ({ page }) => {
  await openInEditor(page, await withBands({ footer: para('Footer') }), 'Alex_Morgan_CV.docx');
  await page.evaluate(() => { window.printed = 0; window.print = () => { window.printed++; }; });
  await page.getByRole('button', { name: 'More ways to save' }).click();
  const menu = page.getByRole('menu', { name: 'Save' });
  // Save as… where the browser can pick a file to save to (Chrome and Edge: see save-in-place.spec.js).
  const saveAs = await page.evaluate(() => typeof window.showSaveFilePicker === 'function');
  await expect(menu.getByRole('menuitem')).toHaveText([/Word document \(\.docx\)/, ...(saveAs ? [/Save as…/] : []), /PDF…/, /Mail merge…/]);
  await menu.getByRole('menuitem', { name: /PDF/ }).click();
  const dialog = page.getByRole('dialog', { name: 'Save as PDF' });
  await expect(dialog).toContainText('Save as PDF');
  await expect(dialog).toContainText('Alex_Morgan_CV.pdf');
  await expect(dialog).toContainText(/About 2 pages on A4 paper, with the footer on each/);
  await expect(dialog).not.toContainText('tracked change');
  await dialog.getByRole('button', { name: 'Open print window' }).click();
  await expect(dialog).toBeHidden();
  await expect.poll(() => page.evaluate(() => window.printed)).toBe(1);
});

test('the PDF dialog warns that tracked changes and highlights print as they show', async ({ page }) => {
  const body = '<w:p><w:ins w:id="1" w:author="Mentor" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Added by my mentor</w:t></w:r></w:ins></w:p>'
    + '<w:p><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>Check this</w:t></w:r></w:p>';
  await openInEditor(page, await buildDocx(body));
  await page.getByRole('button', { name: 'More ways to save' }).click();
  await page.getByRole('menuitem', { name: /PDF/ }).click();
  const dialog = page.getByRole('dialog', { name: 'Save as PDF' });
  await expect(dialog).toContainText('A tracked change prints marked, as it shows here. For a clean copy, accept or reject it first.');
  await expect(dialog).toContainText('Highlighted text prints highlighted.');
  await dialog.getByRole('button', { name: 'Cancel' }).click();
  await expect(dialog).toBeHidden();
});

test('the Word document item downloads as Download does', async ({ page }) => {
  await openInEditor(page, await buildDocx(para('Text')));
  await page.getByRole('button', { name: 'More ways to save' }).click();
  await page.getByRole('menuitem', { name: /Word document/ }).click();
  await expect(page.getByRole('dialog', { name: 'Review your changes' })).toBeVisible();
});

test('printing suggests the file\'s name for the PDF, and gives the title back after', async ({ page }) => {
  await openInEditor(page, await buildDocx(para('Text')), 'Cover_Letter_Initech.docx');
  const title = await page.title();
  await page.evaluate(() => window.dispatchEvent(new Event('beforeprint')));
  await expect(page).toHaveTitle('Cover_Letter_Initech');
  await page.evaluate(() => window.dispatchEvent(new Event('afterprint')));
  await expect(page).toHaveTitle(title);
});

test('a header and footer the same on every page are printed on every page', async ({ page }) => {
  await openInEditor(page, await withBands({ header: para('Alex Morgan – CV'), footer: '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Confidential</w:t></w:r></w:p>' }));
  const css = await beforePrint(page);
  // The margins become padding repeated on each page, the paper's edge the page's.
  expect(css).toMatch(/@page \{ size: [\d.]+px [\d.]+px; margin: 0px [\d.]+px 0px/);
  await page.emulateMedia({ media: 'print' });
  for (const which of ['header', 'footer']) await expect(page.locator(`.docx-hf-band.${which}`)).toHaveCSS('position', 'fixed');
  await expect(page.locator('.docx-hf-band.header')).toHaveCSS('top', '47.2px'); // 708 twips
  await expect(page.locator('.docx-page')).toHaveCSS('padding-top', '96px');
  await expect(page.locator('.docx-page')).toHaveCSS('box-decoration-break', 'clone');
  await expect(page.locator('.docx-hf-band .docx-hf-part').first()).toHaveCSS('opacity', '1');
  await page.emulateMedia({ media: 'screen' });
  await page.evaluate(() => window.dispatchEvent(new Event('afterprint')));
  await expect(page.locator('.docx-hf-band.header')).not.toHaveCSS('position', 'fixed');
});

test('page numbers print numbered on each page, in the page\'s margin boxes', async ({ page, browserName }) => {
  test.skip(browserName !== 'chromium', 'Page-margin boxes: Chromium');
  const footer = `<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t xml:space="preserve">Page </w:t></w:r>${field('PAGE', '1')}<w:r><w:t xml:space="preserve"> of </w:t></w:r>${field('NUMPAGES', '1')}</w:p>`;
  await openInEditor(page, await withBands({ header: para('Alex Morgan – CV'), footer }));
  const css = await beforePrint(page);
  expect(css).toContain('@bottom-right { content: "Page " counter(page) " of " counter(pages);');
  // The header, without a number, repeats as it is.
  await page.emulateMedia({ media: 'print' });
  await expect(page.locator('.docx-hf-band.header')).toHaveCSS('position', 'fixed');
  await expect(page.locator('.docx-hf-band.footer')).toBeHidden();
  await page.emulateMedia({ media: null }); // (a PDF made under emulated screen media is of the screen)
  expect(pdfPages(await page.pdf({ preferCSSPageSize: true }))).toBe(2);
});

test('a simple page-number field and Word\'s absolute tabs are numbered too', async ({ page, browserName }) => {
  test.skip(browserName !== 'chromium', 'Page-margin boxes: Chromium');
  const footer = '<w:p><w:r><w:t>Alex Morgan</w:t></w:r><w:r><w:ptab w:relativeTo="margin" w:alignment="right" w:leader="none"/></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>';
  await openInEditor(page, await withBands({ footer }));
  const css = await beforePrint(page);
  expect(css).toContain('@bottom-left { content: "Alex Morgan";');
  expect(css).toMatch(/@bottom-(center|right) \{ content: counter\(page\);/);
});

test('a different first page has its own header, and none where it has none', async ({ page, browserName }) => {
  test.skip(browserName !== 'chromium', 'Page-margin boxes: Chromium');
  await openInEditor(page, await withBands({ header: para('Alex Morgan – page 2 onwards'), firstHeader: '<w:p/>', footer: '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>' + field('PAGE', '1') + '</w:p>', firstFooter: '<w:p/>' }));
  const css = await beforePrint(page);
  expect(css).toContain('@top-left { content: "Alex Morgan – page 2 onwards";');
  expect(css).toMatch(/@page :first \{[^}]*@top-left \{ content: none; \}/);
  expect(css).toMatch(/@page :first \{.*@bottom-center \{ content: none; \}/);
});

test('headers of several sections print once, as before, rather than on the wrong pages', async ({ page }) => {
  await openInEditor(page, await withBands({ header: para('Last section header'), sections: true }));
  const css = await beforePrint(page);
  expect(css).toMatch(/@page \{ size: [\d.]+px [\d.]+px; margin: 96px/);
  await page.emulateMedia({ media: 'print' });
  await expect(page.locator('.docx-hf-band.header')).not.toHaveCSS('position', 'fixed');
  await page.emulateMedia({ media: 'screen' });
  await page.getByRole('button', { name: 'More ways to save' }).click();
  await page.getByRole('menuitem', { name: /PDF/ }).click();
  await expect(page.getByRole('dialog', { name: 'Save as PDF' })).toContainText('The header prints once, at the start');
});

test('a narrow paper doesn\'t print in the phone\'s layout', async ({ page }) => {
  // The print viewport is the paper's text width, under the phone layout's 860 px.
  await page.setViewportSize({ width: 700, height: 900 });
  await openInEditor(page, await buildDocx(para('Text')));
  await page.emulateMedia({ media: 'print' });
  await expect(page.locator('.docx-page')).toHaveCSS('padding-left', '0px');
  await expect(page.locator('.docx-page')).toHaveCSS('padding-top', '0px');
});
