// Links: pasted links to web and e-mail addresses are kept, and Ctrl+K inserts a link (to a web or
// e-mail address or a DOI), changes the one at the caret or removes it. They are written as Word
// writes them: <w:hyperlink r:id> in Word's Hyperlink style, with an external relationship in the
// part the link is in, taken out again if the link is deleted.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, W_NS } from './helpers.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const DOC = '<w:p><w:r><w:t xml:space="preserve">See the data here.</w:t></w:r></w:p><w:p><w:r><w:t>Other</w:t></w:r></w:p>';

async function pasteHtml(page, html, text, where = 'See the data ') {
  await page.evaluate(({ html, text, where }) => {
    const pg = document.querySelector('.docx-page');
    pg.focus();
    const walker = document.createTreeWalker(pg, NodeFilter.SHOW_TEXT);
    let t = walker.nextNode();
    while (t && !t.data.includes(where)) t = walker.nextNode();
    t.parentElement.closest('[contenteditable="true"]').focus();
    const range = document.createRange();
    range.setStart(t, t.data.indexOf(where) + where.length);
    getSelection().removeAllRanges();
    getSelection().addRange(range);
    const dt = new DataTransfer();
    dt.setData('text/html', html);
    dt.setData('text/plain', text);
    const event = new ClipboardEvent('paste', { bubbles: true, cancelable: true });
    Object.defineProperty(event, 'clipboardData', { value: dt });
    t.parentElement.dispatchEvent(event);
  }, { html, text, where });
}

/** Selects `text` in the page. */
async function select(page, text) {
  await page.evaluate((text) => {
    const walker = document.createTreeWalker(document.querySelector('.docx-page'), NodeFilter.SHOW_TEXT);
    for (let t = walker.nextNode(); t; t = walker.nextNode()) {
      const i = t.data.indexOf(text);
      if (i < 0) continue;
      t.parentElement.closest('[contenteditable="true"]').focus();
      const range = document.createRange();
      range.setStart(t, i);
      range.setEnd(t, i + text.length);
      getSelection().removeAllRanges();
      getSelection().addRange(range);
      return;
    }
    throw new Error(`no ${text}`);
  }, text);
}

const dialog = (page) => page.getByRole('dialog', { name: 'Link' });
const rels = async (zip, path = 'word/_rels/document.xml.rels') => zip.file(path).async('string');
const relFor = (relsXml, href) => relsXml.match(new RegExp(`<Relationship Id="([^"]+)" Type="${REL}/hyperlink" Target="${href.replace(/[.?/]/g, '\\$&')}" TargetMode="External"/>`))?.[1];

test('a pasted link is kept as a Word hyperlink, in the Hyperlink style', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await pasteHtml(page, 'from <a href="https://example.org/data?id=7">the archive</a>, ', 'from the archive, ');
  await expect(page.locator('.docx-page a[data-link="new"]')).toHaveText('the archive');
  const { zip, xml } = await download(page);
  const rid = relFor(await rels(zip), 'https://example.org/data?id=7');
  expect(rid).toBeTruthy();
  expect(xml).toContain(`<w:hyperlink r:id="${rid}" w:history="1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t xml:space="preserve">the archive</w:t></w:r></w:hyperlink>`);
  const styles = await zip.file('word/styles.xml').async('string');
  expect(styles).toContain('<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/>');
});

test('Ctrl+K makes the selection a link; a DOI becomes a doi.org address', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await select(page, 'data');
  await page.keyboard.press('ControlOrMeta+k');
  await expect(dialog(page)).toBeVisible();
  await dialog(page).getByRole('textbox', { name: /Address/ }).fill('not a link');
  await dialog(page).getByRole('button', { name: 'Insert' }).click();
  await expect(dialog(page).getByRole('alert')).toContainText('Enter a web address');
  await dialog(page).getByRole('textbox', { name: /Address/ }).fill('doi:10.1000/xyz123');
  await dialog(page).getByRole('button', { name: 'Insert' }).click();
  await expect(page.locator('.docx-page a[data-link="new"]')).toHaveText('data');
  await expect(page.locator('.docx-page a[data-link="new"]')).toHaveAttribute('href', 'https://doi.org/10.1000/xyz123');
  await expect(page.locator('.docx-page p').first()).toHaveText('See the data here.');
  const { zip, xml } = await download(page);
  const rid = relFor(await rels(zip), 'https://doi.org/10.1000/xyz123');
  expect(xml).toContain(`<w:r><w:t xml:space="preserve">See the </w:t></w:r><w:hyperlink r:id="${rid}" w:history="1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t xml:space="preserve">data</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> here.</w:t></w:r>`);
});

test('at the caret, Ctrl+K inserts new text as a link; while tracking it is a tracked insertion inside the link', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await page.getByRole('button', { name: 'Track changes' }).click();
  const author = page.getByRole('dialog', { name: 'Your name on tracked changes' });
  await author.getByRole('textbox').fill('Rita');
  await author.getByRole('button', { name: 'Start tracking' }).click();
  await page.getByText('Other').click();
  await page.keyboard.press('End');
  await page.keyboard.press('ControlOrMeta+k');
  await dialog(page).getByRole('textbox', { name: /Address/ }).fill('www.example.com');
  await dialog(page).getByRole('textbox', { name: 'Text to show' }).fill(' (website)');
  await dialog(page).getByRole('button', { name: 'Insert' }).click();
  await expect(page.locator('.statusbar [role=status]')).toHaveText('Link inserted (tracked)');
  const { zip, xml } = await download(page);
  const rid = relFor(await rels(zip), 'https://www.example.com');
  expect(xml).toMatch(new RegExp(`<w:hyperlink r:id="${rid}" w:history="1"><w:ins w:id="\\d+" w:author="Rita" w:date="[^"]+"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t xml:space="preserve">\\(website\\)</w:t></w:r></w:ins></w:hyperlink>`));
});

test('a link in the file can be given a new address or removed; Undo brings it back', async ({ page }) => {
  const body = '<w:p><w:r><w:t xml:space="preserve">Read </w:t></w:r><w:hyperlink r:id="rId3" w:history="1" w:tooltip="Old"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>the paper</w:t></w:r></w:hyperlink><w:r><w:t>.</w:t></w:r></w:p>';
  await openInEditor(page, await buildDocx(body));
  await page.getByText('the paper').click();
  await page.keyboard.press('ControlOrMeta+k');
  await expect(dialog(page).getByRole('textbox', { name: /Address/ })).toHaveValue('https://example.com/');
  await dialog(page).getByRole('textbox', { name: /Address/ }).fill('https://journal.example/paper');
  await dialog(page).getByRole('button', { name: 'Update' }).click();
  let { zip, xml } = await download(page);
  const rid = relFor(await rels(zip), 'https://journal.example/paper');
  // The link keeps its own settings (its tooltip), with the new address.
  expect(xml).toContain(`<w:hyperlink w:history="1" w:tooltip="Old" r:id="${rid}">`);
  expect(await rels(zip)).toContain('Id="rId3"'); // the old relationship is left as it was

  await page.getByText('the paper').click();
  await page.keyboard.press('ControlOrMeta+k');
  await dialog(page).getByRole('button', { name: 'Remove link' }).click();
  await expect(page.locator('.docx-page a')).toHaveCount(0);
  await expect(page.locator('.docx-page p')).toHaveText('Read the paper.');
  ({ zip, xml } = await download(page));
  expect(xml).not.toContain('hyperlink');
  expect(relFor(await rels(zip), 'https://journal.example/paper')).toBeUndefined(); // no longer used: gone

  await page.keyboard.press('ControlOrMeta+z');
  await expect(page.locator('.docx-page a')).toHaveText('the paper');
});

test('a link pasted into a footnote is related from the footnotes part', async ({ page }) => {
  const notes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes ${W_NS}><w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> See the data </w:t></w:r></w:p></w:footnote></w:footnotes>`;
  await openInEditor(page, await buildDocx('<w:p><w:r><w:t>Body</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>', {
    parts: { 'word/footnotes.xml': notes },
    rels: `<Relationship Id="rIdF" Type="${REL}/footnotes" Target="footnotes.xml"/>`,
  }));
  await pasteHtml(page, '<a href="https://zenodo.example/record/1">on Zenodo</a>', 'on Zenodo');
  await expect(page.locator('.docx-note a[data-link="new"]')).toHaveText('on Zenodo');
  const { zip } = await download(page);
  const noteRels = await rels(zip, 'word/_rels/footnotes.xml.rels');
  const rid = relFor(noteRels, 'https://zenodo.example/record/1');
  expect(await zip.file('word/footnotes.xml').async('string')).toContain(`<w:hyperlink r:id="${rid}" w:history="1">`);
  expect(relFor(await rels(zip), 'https://zenodo.example/record/1')).toBeUndefined(); // not in the document's
});

test('a deleted pasted link leaves no relationship behind; the file is as it was', async ({ page }) => {
  const buffer = await buildDocx(DOC);
  await openInEditor(page, buffer);
  await pasteHtml(page, '<a href="https://example.org/x">gone</a>', 'gone');
  await expect(page.locator('.docx-page a')).toHaveText('gone');
  let { zip } = await download(page);
  expect(relFor(await rels(zip), 'https://example.org/x')).toBeTruthy();
  await page.keyboard.press('ControlOrMeta+z');
  await expect(page.locator('.docx-page a')).toHaveCount(0);
  ({ zip } = await download(page));
  expect(relFor(await rels(zip), 'https://example.org/x')).toBeUndefined();
});
