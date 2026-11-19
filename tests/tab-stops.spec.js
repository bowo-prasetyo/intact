// Tab stops, as Word's Tabs dialog sets them: the Paragraph dialog lists a paragraph's stops (its
// style's and its own), sets and removes them, and writes only the paragraph's own <w:tabs>, a clear
// for a style's stop taken off. The page puts tabs at the paragraph's first left stop, or at Word's
// default stops (w:defaultTabStop).
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const STYLES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
  + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
  + '<w:style w:type="paragraph" w:styleId="Form"><w:name w:val="Form"/><w:basedOn w:val="Normal"/><w:pPr><w:tabs><w:tab w:val="left" w:pos="2835"/><w:tab w:val="right" w:leader="dot" w:pos="9000"/></w:tabs></w:pPr></w:style>'
  + '</w:styles>';
const DOC = '<w:p><w:r><w:t>No</w:t></w:r><w:r><w:tab/><w:t>: 12/ADM/2026</w:t></w:r></w:p>'
  + '<w:p><w:pPr><w:pStyle w:val="Form"/></w:pPr><w:r><w:t>Perihal</w:t></w:r><w:r><w:tab/><w:t>: Undangan</w:t></w:r></w:p>'
  + '<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="1701"/></w:tabs></w:pPr><w:r><w:t>Lampiran</w:t></w:r><w:r><w:tab/><w:t>: 1 berkas</w:t></w:r></w:p>';
const OPTS = { parts: { 'word/styles.xml': STYLES } };
const savedParagraph = (xml, text) => xml.match(/<w:p[ >].*?<\/w:p>/gs).find((p) => p.includes(`>${text}<`));
const para = (page, text) => page.locator('.docx-page p', { hasText: text });

async function openParagraphDialog(page, text) {
  await para(page, text).click();
  await page.keyboard.press('End');
  await page.getByRole('button', { name: 'Line and paragraph spacing' }).click();
  await page.getByRole('menu', { name: 'Line and paragraph spacing' }).getByRole('menuitem', { name: 'Paragraph…' }).click();
  return page.getByRole('dialog', { name: 'Paragraph' });
}

/** Where the text after the tab starts, in cm from where the paragraph's text starts. */
const afterTab = (page, text) => para(page, text).evaluate((p) => {
  const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
  let t = walker.nextNode();
  while (t && !t.data.includes(':')) t = walker.nextNode();
  const r = document.createRange();
  r.setStart(t, t.data.indexOf(':'));
  r.setEnd(t, t.data.indexOf(':') + 1);
  const start = p.getBoundingClientRect().left + parseFloat(getComputedStyle(p).paddingLeft);
  return Math.round(((r.getBoundingClientRect().left - start) / 96) * 2.54 * 100) / 100;
});

test('tabs go to the paragraph’s first left stop, or to Word’s default stops', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC, OPTS));
  expect(await afterTab(page, 'No')).toBeCloseTo(1.27, 1); // the default stop, half an inch (short text: before it)
  expect(await afterTab(page, 'Perihal')).toBeCloseTo(5, 1); // the style's 2835 twips
  expect(await afterTab(page, 'Lampiran')).toBeCloseTo(3, 1); // its own 1701 twips
});

test('the Paragraph dialog lists the tab stops and sets a new one, written as Word writes it', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC, OPTS));
  const dialog = await openParagraphDialog(page, 'No');
  await expect(dialog.getByText('None: tabs go to Word’s default stops, every 1.27 cm.')).toBeVisible();
  await dialog.getByLabel('Tab stop position (cm)').fill('4');
  await dialog.getByRole('button', { name: 'Set' }).click();
  await expect(dialog.getByRole('list', { name: 'Tab stops' }).getByRole('listitem')).toHaveText([/4 cm\s*Left/]);
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect.poll(() => afterTab(page, 'No')).toBeCloseTo(4, 1);

  const { xml } = await download(page);
  expect(savedParagraph(xml, 'No')).toContain('<w:tabs><w:tab w:val="left" w:pos="2268"/></w:tabs>');
  // The others are as they were.
  expect(savedParagraph(xml, 'Lampiran')).toContain('<w:tabs><w:tab w:val="left" w:pos="1701"/></w:tabs>');
});

test('a style’s tab stop removed in the dialog is cleared in the paragraph, and a right stop gets its leader', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC, OPTS));
  const dialog = await openParagraphDialog(page, 'Perihal');
  const stops = dialog.getByRole('list', { name: 'Tab stops' }).getByRole('listitem');
  await expect(stops).toHaveText([/5 cm\s*Left/, /15\.88 cm\s*Right\s*……… dots/]);
  await dialog.getByRole('button', { name: 'Remove the tab stop at 5 cm' }).click();
  await dialog.getByLabel('Tab stop position (cm)').fill('3.5');
  await dialog.getByLabel('Alignment').selectOption('center');
  await dialog.getByLabel('Leader').selectOption('hyphen');
  // OK sets the stop typed, as Word's does.
  await dialog.getByRole('button', { name: 'OK' }).click();

  const { xml } = await download(page);
  const saved = savedParagraph(xml, 'Perihal');
  expect(saved).toContain('<w:pStyle w:val="Form"/><w:tabs><w:tab w:val="center" w:leader="hyphen" w:pos="1984"/><w:tab w:val="clear" w:pos="2835"/></w:tabs>');
  expect(saved).not.toContain('w:pos="9000"'); // still the style's: not written again
});

test('Clear all, while tracking, is a tracked paragraph formatting change; Ctrl+Z puts the stops back', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC, OPTS));
  await page.getByRole('button', { name: 'Track changes' }).click();
  const named = page.getByRole('dialog', { name: 'Your name on tracked changes' });
  await named.getByRole('textbox').fill('Rita');
  await named.getByRole('button', { name: 'Start tracking' }).click();
  const dialog = await openParagraphDialog(page, 'Lampiran');
  await dialog.getByRole('button', { name: 'Clear all' }).click();
  await expect(dialog.getByText(/None: tabs go to Word’s default stops/)).toBeVisible();
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect.poll(() => afterTab(page, 'Lampiran')).toBeCloseTo(2.54, 1); // the second default stop

  const { xml } = await download(page);
  const saved = savedParagraph(xml, 'Lampiran');
  expect(saved).toMatch(/<w:pPrChange [^>]*w:author="Rita"[^>]*><w:pPr><w:tabs><w:tab w:val="left" w:pos="1701"\/><\/w:tabs><\/w:pPr><\/w:pPrChange>/);
  expect(saved).not.toMatch(/<w:pPr><w:tabs>.*<w:pPrChange/);

  await page.locator('.docx-page').focus();
  await page.keyboard.press('Control+z');
  await expect.poll(() => afterTab(page, 'Lampiran')).toBeCloseTo(3, 1);
});
