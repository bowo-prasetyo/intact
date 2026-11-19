// Draw signature: a signature drawn (or typed) on the pad goes in at the caret, or in place of the
// selected "____" line, as an ordinary picture in line with the text, alt text "Signature".
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const FORM = '<w:p><w:r><w:t xml:space="preserve">Parent signature: ______________</w:t></w:r></w:p><w:p><w:r><w:t>Date:</w:t></w:r></w:p>';
const status = (page) => page.locator('.statusbar [role=status]');

/** Draws a squiggle across the pad. */
async function sign(page) {
  const pad = page.locator('.sig canvas');
  const box = await pad.boundingBox();
  await page.mouse.move(box.x + box.width * 0.15, box.y + box.height * 0.6);
  await page.mouse.down();
  for (let i = 1; i <= 20; i++) {
    await page.mouse.move(box.x + box.width * (0.15 + i * 0.03), box.y + box.height * (0.6 - 0.25 * Math.sin(i / 2)));
  }
  await page.mouse.up();
}

/** Selects `text` in the page. */
async function select(page, text) {
  await page.evaluate((t) => {
    const walker = document.createTreeWalker(document.querySelector('.docx-page'), NodeFilter.SHOW_TEXT);
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      const i = n.data.indexOf(t);
      if (i < 0) continue;
      const r = document.createRange();
      r.setStart(n, i);
      r.setEnd(n, i + t.length);
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(r);
      return;
    }
  }, text);
}

test('a drawn signature replaces the selected line, as a picture with alt text', async ({ page }) => {
  await openInEditor(page, await buildDocx(FORM));
  await page.locator('.docx-page p').first().click();
  await select(page, '______________');
  await page.getByRole('button', { name: 'Draw signature' }).click();
  const dialog = page.getByRole('dialog', { name: 'Draw your signature' });
  await expect(dialog).toBeVisible();
  const insert = dialog.getByRole('button', { name: 'Insert signature' });
  await expect(insert).toBeDisabled();
  await sign(page);
  await expect(insert).toBeEnabled();
  await dialog.getByLabel('Blue').check();
  await insert.click();
  await expect(dialog).toBeHidden();
  await expect(status(page)).toContainText('Signature inserted');

  const first = page.locator('.docx-page p').first();
  await expect(first).toHaveText('Parent signature: ');
  const img = first.locator('img');
  await expect(img).toHaveAttribute('alt', 'Signature');
  // Cropped to the ink and inserted at half the pad's size: a signature, not a page-wide picture.
  const width = await img.evaluate((el) => el.getBoundingClientRect().width);
  expect(width).toBeGreaterThan(80);
  expect(width).toBeLessThan(260);

  const { zip, xml } = await download(page);
  expect(xml).toMatch(/<w:t xml:space="preserve">Parent signature: <\/w:t><\/w:r><w:r><w:drawing><wp:inline[^>]*>.*descr="Signature"/);
  const media = Object.keys(zip.files).filter((f) => f.startsWith('word/media/'));
  const png = media.find((f) => f.endsWith('.png') && f !== 'word/media/image1.png');
  expect(png).toBeTruthy();
  // A PNG with a transparent background, so the line or text beneath shows through.
  const bytes = await zip.file(png).async('uint8array');
  expect([...bytes.slice(1, 4)].map((b) => String.fromCharCode(b)).join('')).toBe('PNG');
  expect(bytes[25]).toBe(6); // colour type 6: RGBA
});

test('a typed name can be inserted instead, and Cancel leaves the text as it was', async ({ page }) => {
  await openInEditor(page, await buildDocx(FORM));
  await page.locator('.docx-page p').nth(1).click();
  await page.keyboard.press('End');
  await page.getByRole('button', { name: 'Draw signature' }).click();
  const dialog = page.getByRole('dialog', { name: 'Draw your signature' });
  await dialog.getByRole('button', { name: 'Cancel' }).click();
  await expect(dialog).toBeHidden();
  await expect(page.locator('.docx-page img')).toHaveCount(0);

  await page.getByRole('button', { name: 'Draw signature' }).click();
  await dialog.getByText('Or type your name').click();
  await dialog.getByLabel('Name, written in a handwriting font').fill('Alex Morgan');
  await dialog.getByRole('button', { name: 'Insert signature' }).click();
  await expect(page.locator('.docx-page p').nth(1).locator('img[alt="Signature"]')).toHaveCount(1);
  await expect(page.locator('.docx-page p').nth(1)).toHaveText('Date:');

  await page.keyboard.press('Control+z');
  await expect(page.locator('.docx-page img')).toHaveCount(0);
});
