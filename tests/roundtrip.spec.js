// The product promise: "nothing you didn't touch changes". These tests open documents, save them
// (with and without edits) and check that everything the user did not edit means the same thing.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, canonicalize, documentXmlOf, KITCHEN_SINK } from './helpers.js';

test('opening and downloading without edits keeps the document unchanged', async ({ page }) => {
  const docx = await buildDocx(KITCHEN_SINK);
  await openInEditor(page, docx);
  const { xml } = await download(page);
  const original = await canonicalize(page, await documentXmlOf(docx));
  expect(await canonicalize(page, xml)).toEqual(original);
});

test('a second save of a saved file changes nothing', async ({ page }) => {
  await openInEditor(page, await buildDocx(KITCHEN_SINK));
  const first = await download(page);
  await page.getByRole('button', { name: 'Close' }).click();
  await openInEditor(page, await first.zip.generateAsync({ type: 'nodebuffer' }));
  const second = await download(page);
  expect(await canonicalize(page, second.xml)).toEqual(await canonicalize(page, first.xml));
});

test('editing one paragraph leaves every other paragraph unchanged', async ({ page }) => {
  const docx = await buildDocx(KITCHEN_SINK);
  await openInEditor(page, docx);
  const last = page.locator('.docx-page > p').last();
  await last.click();
  await page.keyboard.press('End');
  await page.keyboard.type(' edited');
  const { xml } = await download(page);

  const before = await canonicalize(page, await documentXmlOf(docx));
  const after = await canonicalize(page, xml);
  expect(after.length).toBe(before.length);
  const edited = before.findIndex((b) => b.p?.inlines.some((i) => i.text === 'Last paragraph'));
  after.forEach((block, i) => {
    if (i !== edited) expect(block, `block ${i}`).toEqual(before[i]);
  });
  expect(after[edited].p.inlines).toEqual([{ rPr: '', text: 'Last paragraph edited' }]);
});

test('tracked changes keep their author and date, and are shown read-only', async ({ page }) => {
  await openInEditor(page, await buildDocx(`<w:p><w:r><w:t xml:space="preserve">Hello </w:t></w:r><w:ins w:id="1" w:author="Ann" w:date="2026-01-01T00:00:00Z"><w:r><w:t>world</w:t></w:r></w:ins></w:p>`));
  const ins = page.locator('.docx-page .raw-inline.inserted');
  await expect(ins).toHaveText('world');
  await expect(ins).toHaveAttribute('contenteditable', 'false');
  await expect(ins).toHaveAttribute('title', /Inserted by Ann/);

  await page.locator('.docx-page p').first().click();
  await page.keyboard.press('Home');
  await page.keyboard.type('Oh, ');
  const { xml } = await download(page);
  expect(xml).toContain('w:author="Ann"');
  expect(xml).toMatch(/<w:ins [^>]*><w:r><w:t>world<\/w:t><\/w:r><\/w:ins>/);
  expect(xml).toContain('Oh, Hello ');
});
