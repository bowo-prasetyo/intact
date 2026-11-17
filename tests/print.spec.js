// Printing (Ctrl+P) prints the document, not the editor: no toolbar, header bar, status bar or
// comments; the document's own paper size and margins; and a new page at each page break.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor } from './helpers.js';

const DOC = '<w:p><w:r><w:t>First page</w:t></w:r></w:p>'
  + '<w:p><w:r><w:t>Before the break</w:t></w:r><w:r><w:br w:type="page"/></w:r><w:r><w:t>Second page</w:t></w:r></w:p>'
  + '<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>Third page</w:t></w:r></w:p>';

test('only the page is printed', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await page.emulateMedia({ media: 'print' });
  for (const chrome of ['.app-header', '.toolbar', '.statusbar']) await expect(page.locator(chrome)).toBeHidden();
  await expect(page.locator('.docx-page')).toBeVisible();
  await expect(page.locator('.docx-sheet')).toHaveCSS('box-shadow', 'none');
  // The page's margins come from @page, so the sheet adds none of its own.
  await expect(page.locator('.docx-page')).toHaveCSS('padding-left', '0px');
  await expect(page.locator('.docx-page .page-break')).toHaveCSS('break-after', 'page');
  await expect(page.locator('.docx-page .page-break')).toHaveCSS('border-top-style', 'none');
  await page.emulateMedia({ media: 'screen' });
  await expect(page.locator('.toolbar')).toBeVisible();
});

test('the printout has the document\'s paper size and a page per page break', async ({ page, browserName }) => {
  test.skip(browserName !== 'chromium', 'Only Chromium makes PDFs');
  await openInEditor(page, await buildDocx(DOC));
  const pdf = (await page.pdf({ preferCSSPageSize: true })).toString('latin1');
  const pages = pdf.match(/\/Type\s*\/Page[^s]/g) || [];
  expect(pages).toHaveLength(3);
  // A4 (11906 × 16838 twips) is 595.3 × 841.9 points.
  const [, w, h] = pdf.match(/\/MediaBox\s*\[\s*0 0 ([\d.]+) ([\d.]+)\s*\]/);
  expect(Math.round(Number(w))).toBe(595);
  expect(Math.round(Number(h))).toBe(842);
});
