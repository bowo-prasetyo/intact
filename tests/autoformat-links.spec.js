// AutoFormat's links, as Word's "Internet and network paths with hyperlinks": a web or e-mail
// address typed in the text becomes a link when a space or Enter follows it. Punctuation around it
// stays outside; Ctrl+Z straight after gives the address back as plain text; the AutoFormat switch
// turns it off; typed while tracking, the link goes with the tracked insertion.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const DOC = '<w:p><w:r><w:t>Start</w:t></w:r></w:p>';
const para = (page) => page.locator('.docx-page > p').first();
const links = (page) => page.locator('.docx-page a').evaluateAll((as) => as.map((a) => [a.textContent, a.getAttribute('href')]));

async function typeAtEnd(page, text) {
  await para(page).click();
  await page.keyboard.press('End');
  await page.keyboard.type(text);
}

test('web and e-mail addresses become links when a space follows, punctuation left out', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await typeAtEnd(page, ' (portfolio: www.example.com/work). Mail alex.morgan@example.com, or see https://en.wikipedia.org/wiki/Word_(software) and mailto:hr@example.org ok');
  expect(await links(page)).toEqual([
    ['www.example.com/work', 'https://www.example.com/work'],
    ['alex.morgan@example.com', 'mailto:alex.morgan@example.com'],
    ['https://en.wikipedia.org/wiki/Word_(software)', 'https://en.wikipedia.org/wiki/Word_(software)'],
    ['mailto:hr@example.org', 'mailto:hr@example.org'],
  ]);
  await expect(para(page)).toHaveText('Start (portfolio: www.example.com/work). Mail alex.morgan@example.com, or see https://en.wikipedia.org/wiki/Word_(software) and mailto:hr@example.org ok');
  const { zip, xml } = await download(page);
  expect(xml).toMatch(/<w:hyperlink r:id="rIdLink\d+" w:history="1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"\/><\/w:rPr><w:t[^>]*>www\.example\.com\/work<\/w:t><\/w:r><\/w:hyperlink>/);
  const rels = await zip.file('word/_rels/document.xml.rels').async('string');
  for (const target of ['https://www.example.com/work', 'mailto:alex.morgan@example.com', 'mailto:hr@example.org']) {
    expect(rels).toContain(`Target="${target}" TargetMode="External"`);
  }
});

test('Enter after an address makes it a link, then a new paragraph', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await typeAtEnd(page, ' alex@example.com');
  await page.keyboard.press('Enter');
  await page.keyboard.type('Next');
  expect(await links(page)).toEqual([['alex@example.com', 'mailto:alex@example.com']]);
  await expect(page.locator('.docx-page > p')).toHaveText(['Start alex@example.com', 'Next']);
});

test('what isn\'t an address stays as it is', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await typeAtEnd(page, ' version 1.2.3 e.g. Node.js linkedin.com/in/alex a@b www. done ');
  expect(await links(page)).toEqual([]);
});

test('Ctrl+Z straight after gives the address back as plain text', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await typeAtEnd(page, ' www.example.com ');
  expect(await links(page)).toHaveLength(1);
  await expect(page.locator('.statusbar [role=status]')).toHaveText('Made a link to https://www.example.com – Ctrl+Z takes it back');
  await page.keyboard.press('ControlOrMeta+z');
  expect(await links(page)).toEqual([]);
  await expect(para(page)).toHaveText('Start www.example.com ');
  await page.keyboard.type('kept');
  await expect(para(page)).toHaveText('Start www.example.com kept');
});

test('the AutoFormat switch turns it off', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await page.getByRole('button', { name: /AutoFormat: on/ }).click();
  await typeAtEnd(page, ' www.example.com ');
  expect(await links(page)).toEqual([]);
});

test('typed in a link already, nothing more happens', async ({ page }) => {
  const body = '<w:p><w:hyperlink r:id="rId3"><w:r><w:t>example</w:t></w:r></w:hyperlink></w:p>';
  await openInEditor(page, await buildDocx(body));
  // In the middle of the link's text (typed at its end, text goes after the link, as in Word).
  await page.locator('.docx-page a').click();
  await page.keyboard.press('Home');
  for (let i = 0; i < 3; i++) await page.keyboard.press('ArrowRight');
  await page.keyboard.type(' www.inside.example.com ');
  await expect(page.locator('.docx-page a')).toHaveText('exa www.inside.example.com mple');
  expect((await links(page)).map(([, href]) => href)).toEqual(['https://example.com/']);
});

test('typed while tracking, the link goes with the tracked insertion', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await page.keyboard.press('ControlOrMeta+Shift+E');
  await page.locator('#author-name').fill('Alex');
  await page.getByRole('button', { name: 'Start tracking' }).click();
  await typeAtEnd(page, ' https://tracked.example.com ');
  expect(await links(page)).toEqual([['https://tracked.example.com', 'https://tracked.example.com']]);
  const { xml } = await download(page);
  expect(xml).toMatch(/<w:hyperlink r:id="rIdLink\d+" w:history="1"><w:ins [^>]*w:author="Alex"[^>]*><w:r>.*?https:\/\/tracked\.example\.com<\/w:t><\/w:r><\/w:ins><\/w:hyperlink>/);
});

test('in a footnote, the link is written to the notes\' relationships', async ({ page }) => {
  const NOTES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Source</w:t></w:r></w:p></w:footnote></w:footnotes>`;
  const docx = await buildDocx('<w:p><w:r><w:t>Text</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>', {
    parts: { 'word/footnotes.xml': NOTES },
    rels: '<Relationship Id="rIdFn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>',
  });
  await openInEditor(page, docx);
  const note = page.locator('.docx-note p').first();
  await note.click();
  await page.keyboard.press('End');
  await page.keyboard.type(': https://doi.example.org/x ');
  await expect(note.locator('a')).toHaveAttribute('href', 'https://doi.example.org/x');
  const { zip } = await download(page);
  expect(await zip.file('word/footnotes.xml').async('string')).toContain('<w:hyperlink r:id="rIdLink1"');
  expect(await zip.file('word/_rels/footnotes.xml.rels').async('string')).toContain('Target="https://doi.example.org/x"');
});
