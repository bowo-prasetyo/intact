// Formatting changed while tracking is a tracked change, as in Word: text formatting records the old
// run properties (<w:rPrChange>), paragraph style, alignment and lists the old paragraph properties
// (<w:pPrChange>). They show and are accepted or rejected like any tracked formatting change, the
// review lists them as tracked, and formatting changed back records nothing.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const BODY = '<w:p><w:r><w:t xml:space="preserve">Make this word bold.</w:t></w:r></w:p>'
  + '<w:p><w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>Red text</w:t></w:r></w:p>'
  + '<w:p><w:r><w:t>A paragraph to restyle.</w:t></w:r></w:p>';
const group = (page) => page.getByRole('group', { name: 'Tracked changes' });
const paragraphXml = (xml, n) => [...xml.matchAll(/<w:p>(.*?)<\/w:p>/g)][n][1];

async function startTracking(page, name = 'Rita') {
  await page.getByRole('button', { name: 'Track changes' }).click();
  const dialog = page.getByRole('dialog', { name: 'Your name on tracked changes' });
  await dialog.getByRole('textbox').fill(name);
  await dialog.getByRole('button', { name: 'Start tracking' }).click();
}

async function select(page, n, word) {
  await page.evaluate(([i, w]) => {
    const p = document.querySelectorAll('.docx-page > :is(p, h1, h2, h3)')[i];
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    for (let t = walker.nextNode(); t; t = walker.nextNode()) {
      const at = t.data.indexOf(w);
      if (at < 0) continue;
      document.querySelector('.docx-page').focus();
      getSelection().setBaseAndExtent(t, at, t, at + w.length);
      return;
    }
  }, [n, word]);
}

test('bold while tracking records the old formatting; it shows as a tracked change and is listed as tracked', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await startTracking(page);
  await select(page, 0, 'word');
  await page.keyboard.press('ControlOrMeta+b');
  const changed = page.locator('.docx-page [data-fmt="run"]');
  await expect(changed).toHaveText('word');
  await expect(changed).toHaveAttribute('title', /^Formatting changed by Rita on \d{4}-\d\d-\d\d – tracked change/);
  await expect(group(page).locator('.badge')).toHaveText('1');
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  const review = page.getByRole('dialog', { name: 'Review your changes' });
  await expect(review).toContainText('Recorded as tracked changes (1 formatting change)');
  await expect(review).not.toContainText('aren’t tracked');
  await review.getByRole('button', { name: 'Keep editing' }).click();
  const { xml } = await download(page);
  expect(paragraphXml(xml, 0)).toMatch(/<w:r><w:rPr><w:b\/><w:rPrChange w:id="\d+" w:author="Rita" w:date="[^"]+"><w:rPr\/><\/w:rPrChange><\/w:rPr><w:t xml:space="preserve">word<\/w:t><\/w:r>/);
  expect(paragraphXml(xml, 0)).toMatch(/^<w:r><w:t xml:space="preserve">Make this <\/w:t><\/w:r>/);
});

test('formatting changed back while tracking records nothing; changed twice, the first formatting is kept', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await startTracking(page);
  await select(page, 0, 'word');
  await page.keyboard.press('ControlOrMeta+b');
  await select(page, 0, 'word');
  await page.keyboard.press('ControlOrMeta+b');
  await expect(page.locator('.docx-page [data-fmt]')).toHaveCount(0);
  await select(page, 1, 'Red');
  await page.keyboard.press('ControlOrMeta+i');
  await select(page, 1, 'Red');
  await page.keyboard.press('ControlOrMeta+u');
  const { xml } = await download(page);
  expect(paragraphXml(xml, 0)).not.toContain('rPrChange');
  // Italic, then underline: one change, recording the formatting before both (red only).
  expect(paragraphXml(xml, 1)).toMatch(/<w:rPr><w:i\/><w:color w:val="FF0000"\/><w:u w:val="single"\/><w:rPrChange w:id="\d+" w:author="Rita" w:date="[^"]+"><w:rPr><w:color w:val="FF0000"\/><\/w:rPr><\/w:rPrChange><\/w:rPr><w:t xml:space="preserve">Red<\/w:t>/);
});

test('a font colour or size changed while tracking is tracked too', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await startTracking(page);
  await select(page, 1, 'text');
  await page.getByRole('combobox', { name: 'Font size' }).selectOption('16');
  const { xml } = await download(page);
  expect(paragraphXml(xml, 1)).toMatch(/<w:rPr><w:color w:val="FF0000"\/><w:sz w:val="32"\/><w:szCs w:val="32"\/><w:rPrChange w:id="\d+" w:author="Rita" w:date="[^"]+"><w:rPr><w:color w:val="FF0000"\/><\/w:rPr><\/w:rPrChange><\/w:rPr><w:t xml:space="preserve">text<\/w:t>/);
});

test('rejecting a tracked formatting change brings the old formatting back; accepting keeps the new', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await startTracking(page);
  await select(page, 0, 'word');
  await page.keyboard.press('ControlOrMeta+b');
  await select(page, 1, 'Red');
  await page.keyboard.press('ControlOrMeta+i');
  await page.locator('.docx-page [data-fmt="run"]', { hasText: 'word' }).click();
  await group(page).getByRole('button', { name: 'Reject' }).first().click();
  await page.locator('.docx-page [data-fmt="run"]', { hasText: 'Red' }).click();
  await group(page).getByRole('button', { name: 'Accept' }).first().click();
  await expect(page.locator('.docx-page [data-fmt]')).toHaveCount(0);
  const { xml } = await download(page);
  expect(paragraphXml(xml, 0)).toBe('<w:r><w:t xml:space="preserve">Make this word bold.</w:t></w:r>');
  expect(paragraphXml(xml, 1)).toContain('<w:r><w:rPr><w:i/><w:color w:val="FF0000"/></w:rPr><w:t xml:space="preserve">Red</w:t></w:r>');
});

test('a paragraph style and alignment changed while tracking record the old paragraph formatting', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await startTracking(page);
  await select(page, 2, 'paragraph');
  await page.getByRole('combobox', { name: 'Paragraph style' }).selectOption('h1');
  await select(page, 2, 'paragraph');
  await page.keyboard.press('ControlOrMeta+e');
  await expect(page.locator('.docx-page [data-fmt-para]')).toHaveText('A paragraph to restyle.');
  const { xml } = await download(page);
  expect(paragraphXml(xml, 2)).toMatch(/^<w:pPr><w:pStyle w:val="Heading1"\/><w:jc w:val="center"\/><w:pPrChange w:id="\d+" w:author="Rita" w:date="[^"]+"><w:pPr\/><\/w:pPrChange><\/w:pPr>/);
});

test('Undo takes a tracked formatting change back', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await startTracking(page);
  await select(page, 0, 'word');
  await page.keyboard.press('ControlOrMeta+b');
  await page.keyboard.press('ControlOrMeta+z');
  await expect(page.locator('.docx-page [data-fmt]')).toHaveCount(0);
  await expect(page.locator('.docx-page strong')).toHaveCount(0);
  const { xml } = await download(page);
  expect(xml).not.toContain('rPrChange');
});
