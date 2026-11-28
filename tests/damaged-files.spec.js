// A damaged file is refused with a plain reason, and never half-opened: a package cut short, a byte
// changed (its checksum doesn't match, so a changed letter can't open as changed text and be saved
// so), a part that isn't well-formed XML (named), no main document part, or no ZIP at all. The page
// stays on the start page, with the file as it was.
import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import { buildDocx } from './helpers.js';

const TEXT = '<w:p><w:r><w:t>The quick brown fox jumps over the lazy dog.</w:t></w:r></w:p>';

async function refused(page, name, buffer) {
  await page.goto('./');
  await page.setInputFiles('input[type=file]', { name, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer });
  const error = page.locator('.error');
  await expect(error).toBeVisible();
  await expect(page.locator('.docx-page')).toHaveCount(0); // never half-opened
  expect(page.url()).not.toContain('/edit');
  return error;
}

const rezip = async (buffer, change) => {
  const zip = await JSZip.loadAsync(buffer);
  change(zip);
  return zip.generateAsync({ type: 'nodebuffer' });
};

test('a package cut short is refused as damaged', async ({ page }) => {
  const good = await buildDocx(TEXT.repeat(20));
  const error = await refused(page, 'cut.docx', good.subarray(0, Math.floor(good.length * 0.6)));
  await expect(error).toContainText('cut.docx is damaged: its package is incomplete');
  await expect(error).toContainText('Nothing was opened, and the file is as it was');
});

test('a changed byte is caught by its checksum, even where it would still read as text', async ({ page }) => {
  // Stored without compression, so one letter of the text can be changed and the XML stays valid:
  // without the checksum, "quick" would open as "quack" and be saved so.
  const zip = await JSZip.loadAsync(await buildDocx(TEXT));
  const stored = Buffer.from(await zip.generateAsync({ type: 'nodebuffer', compression: 'STORE' }));
  const at = stored.indexOf(Buffer.from('quick'));
  expect(at).toBeGreaterThan(0);
  stored[at + 2] = 'a'.charCodeAt(0);
  const error = await refused(page, 'changed.docx', stored);
  await expect(error).toContainText("its contents don't match their checksums");
});

test('a part that is not well-formed XML is named', async ({ page }) => {
  const broken = await rezip(await buildDocx(TEXT), (z) => z.file('word/styles.xml', '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style>'));
  const error = await refused(page, 'styles.docx', broken);
  await expect(error).toContainText("styles.docx is damaged: its part word/styles.xml isn't well-formed XML");
});

test('a ZIP without a main document part, and a file that is no ZIP at all, say what they are', async ({ page }) => {
  const noMain = await rezip(await buildDocx(TEXT), (z) => z.remove('word/document.xml'));
  await expect(await refused(page, 'empty.docx', noMain)).toContainText('empty.docx has no main document part');
  await expect(await refused(page, 'text.docx', Buffer.from('Not a ZIP package at all, only text.'))).toContainText('Every .docx is a ZIP package, and this one isn’t');
});
