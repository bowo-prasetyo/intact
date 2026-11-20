// An answer key, as Word makes one: the answers made hidden text (w:vanish, Ctrl+Shift+H or the Hidden
// text button), shown here faintly with a dotted line under them, and printed only when Print hidden
// text is chosen: the test prints without its answers, the key with them.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const BODY = '<w:p><w:r><w:t xml:space="preserve">1. The capital of France is </w:t></w:r><w:r><w:t>Paris</w:t></w:r><w:r><w:t>.</w:t></w:r></w:p>'
  + '<w:p><w:r><w:t xml:space="preserve">2. Two and two make </w:t></w:r><w:r><w:t>four</w:t></w:r><w:r><w:t>.</w:t></w:r></w:p>';

/** Selects `word` in the page. */
async function select(page, word) {
  await page.evaluate((w) => {
    const walker = document.createTreeWalker(document.querySelector('.docx-page'), NodeFilter.SHOW_TEXT);
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      const at = n.data.indexOf(w);
      if (at >= 0) {
        document.querySelector('.docx-page').focus();
        getSelection().setBaseAndExtent(n, at, n, at + w.length);
        return;
      }
    }
  }, word);
}

test('hidden text: the answers hidden with the button and Ctrl+Shift+H, written as w:vanish, shown again', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await select(page, 'Paris');
  await page.getByRole('button', { name: 'Hidden text' }).click();
  const paris = page.locator('.docx-page [data-hidden]', { hasText: 'Paris' });
  await expect(paris).toHaveCount(1);
  await expect(paris).toHaveCSS('text-decoration-style', 'dotted');
  await expect(page.locator('.statusbar [role=status]')).toContainText('Hidden text');
  await select(page, 'four');
  await page.keyboard.press('Control+Shift+H');
  await expect(page.locator('.docx-page [data-hidden]', { hasText: 'four' })).toHaveCount(1);

  let { xml } = await download(page);
  expect(xml).toMatch(/<w:r><w:rPr><w:vanish\/><\/w:rPr><w:t[^>]*>Paris<\/w:t><\/w:r>/);
  expect(xml).toMatch(/<w:vanish\/><\/w:rPr><w:t[^>]*>four<\/w:t>/);

  // Pressed again on hidden text, it is shown again; the button says which it is.
  await select(page, 'four');
  await expect(page.getByRole('button', { name: 'Hidden text' })).toHaveAttribute('aria-pressed', 'true');
  await page.keyboard.press('Control+Shift+H');
  await expect(page.locator('.docx-page [data-hidden]', { hasText: 'four' })).toHaveCount(0);
  ({ xml } = await download(page));
  expect(xml).not.toMatch(/<w:vanish\/><\/w:rPr><w:t[^>]*>four<\/w:t>/);
});

test('printing: hidden text left out, unless Print hidden text is chosen (the answer key)', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await select(page, 'Paris');
  await page.keyboard.press('Control+Shift+H');
  await page.evaluate(() => { window.printed = []; window.print = () => { window.printed.push(document.documentElement.classList.contains('print-hidden-text')); }; });
  const paris = page.locator('.docx-page [data-hidden]', { hasText: 'Paris' });

  // The test: printed without the answer.
  await page.emulateMedia({ media: 'print' });
  await expect(paris).toHaveCSS('display', 'none');
  await page.emulateMedia({ media: 'screen' });

  // The answer key: Print hidden text, in the PDF and print dialog.
  await page.getByRole('button', { name: 'More ways to save' }).click();
  await page.getByRole('menu', { name: 'Save' }).getByRole('menuitem', { name: /PDF/ }).click();
  const dialog = page.getByRole('dialog', { name: 'Save as PDF' });
  const option = dialog.getByLabel(/Print hidden text/);
  await expect(option).not.toBeChecked();
  await option.check();
  await dialog.getByRole('button', { name: 'Open print window' }).click();
  await expect.poll(() => page.evaluate(() => window.printed)).toEqual([true]);
  await page.emulateMedia({ media: 'print' });
  await expect(paris).toHaveCSS('display', 'inline');
  await expect(paris).toHaveCSS('opacity', '1');
  // After printing, it is left out again.
  await page.evaluate(() => window.dispatchEvent(new Event('afterprint')));
  await expect(paris).toHaveCSS('display', 'none');
});

test('a document without hidden text doesn’t offer to print it', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await page.getByRole('button', { name: 'More ways to save' }).click();
  await page.getByRole('menu', { name: 'Save' }).getByRole('menuitem', { name: /PDF/ }).click();
  await expect(page.getByRole('dialog', { name: 'Save as PDF' }).getByLabel(/Print hidden text/)).toHaveCount(0);
});
