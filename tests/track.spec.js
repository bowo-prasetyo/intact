// Track Changes for your own edits: typing, deleting, Enter, paste and Replace are recorded as
// tracked changes (<w:ins>, <w:del>, paragraph-mark revisions) with your name and the date, as Word
// records them, and Word's own setting is saved so the recipient keeps tracking.
import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import { buildDocx, openInEditor, download, downloadFile, W_NS } from './helpers.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const status = (page) => page.locator('.statusbar [role=status]');
const DATE = '\\d{4}-\\d\\d-\\d\\dT\\d\\d:\\d\\d:\\d\\dZ';
/** A regex for a revision element's opening tag by Rita (any id and date). */
const rev = (kind, author = 'Rita') => `<w:${kind} w:id="(\\d+)" w:author="${author}" w:date="${DATE}"`;

// A bookmark with id 7: new revisions must get ids above it.
const DOC = '<w:p><w:bookmarkStart w:id="7" w:name="Fee"/><w:r><w:t xml:space="preserve">The fee is </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>due</w:t></w:r><w:r><w:t>.</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p>'
  + '<w:p><w:r><w:t>Second paragraph.</w:t></w:r></w:p>';

/** Places the caret in the page at `offset` into the text node containing `text`, or selects to `to`. */
async function select(page, [a, ao], [b, bo] = [a, ao]) {
  await page.evaluate(([a, ao, b, bo]) => {
    const root = document.querySelector('.docx-page');
    const find = (t) => {
      const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
      const seen = [];
      for (let n = w.nextNode(); n; n = w.nextNode()) {
        if (n.data.includes(t) && !n.parentElement.closest('.tc-del')) return n;
        seen.push(n.data.replace(/[^\x20-\x7e]/g, (c) => `<U+${c.codePointAt(0).toString(16).toUpperCase().padStart(4, '0')}>`));
      }
      // The text nodes there are, other spaces than U+0020 written out: an engine that splits or
      // re-spaces text when editing (WebKit's non-breaking spaces) shows here, not as a bare miss.
      throw new Error(`no text "${t}" among ${JSON.stringify(seen)}`);
    };
    const s = find(a);
    const e = find(b);
    const range = document.createRange();
    range.setStart(s, s.data.indexOf(a) + ao);
    range.setEnd(e, e.data.indexOf(b) + bo);
    root.focus();
    getSelection().removeAllRanges();
    getSelection().addRange(range);
  }, [a, ao, b, bo]);
}

async function startTracking(page, name = 'Rita') {
  await page.getByRole('button', { name: 'Track changes' }).click();
  const dialog = page.getByRole('dialog', { name: 'Your name on tracked changes' });
  // Firefox's autocomplete list for earlier names would open over "Start tracking" and take the click.
  await expect(dialog.getByRole('textbox')).toHaveAttribute('autocomplete', 'off');
  await dialog.getByRole('textbox').fill(name);
  await dialog.getByRole('button', { name: 'Start tracking' }).click();
  await expect(page.getByRole('button', { name: 'Track changes' })).toHaveAttribute('aria-pressed', 'true');
  await expect(page.locator('.statusbar')).toContainText(`Tracking changes as ${name}`);
}

const settingsOf = (zip) => zip.file('word/settings.xml')?.async('string');

test('typing while tracking records an insertion with your name, and Word keeps tracking', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await startTracking(page);
  await select(page, ['The fee is ', 4]);
  await page.keyboard.type('total ');
  await expect(page.locator('.docx-page .tc-ins')).toHaveText('total ');
  await expect(page.locator('.docx-page .tc-ins')).toHaveAttribute('title', /^Inserted by Rita – tracked change/);
  const { zip, xml } = await download(page);
  const m = xml.match(new RegExp(`<w:t xml:space="preserve">The </w:t></w:r>${rev('ins')}><w:r><w:t xml:space="preserve">total </w:t></w:r></w:ins><w:r><w:t xml:space="preserve">fee is </w:t>`));
  expect(m, xml).not.toBeNull();
  expect(Number(m[1])).toBeGreaterThan(7);
  expect(await settingsOf(zip)).toContain('<w:trackRevisions/>');
});

test('Backspace and Delete mark text deleted, keeping it and its formatting', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await startTracking(page);
  await select(page, ['.', 0]); // after "due" (bold)
  for (let i = 0; i < 4; i++) await page.keyboard.press('Backspace'); // "due" and the space before it
  // One deletion, in a span per formatting (the bold word and the plain space).
  expect((await page.locator('.docx-page .tc-del').allTextContents()).join('')).toBe(' due');
  await page.keyboard.press('Delete'); // the "." after the deletion
  await expect(page.locator('.docx-page p').first()).toHaveText('The fee is due.'); // nothing disappears
  const { xml } = await download(page);
  expect(xml).toMatch(new RegExp(`<w:t xml:space="preserve">The fee is</w:t></w:r>${rev('del')}><w:r><w:delText xml:space="preserve"> </w:delText></w:r><w:r><w:rPr><w:b/></w:rPr><w:delText xml:space="preserve">due</w:delText></w:r><w:r><w:delText xml:space="preserve">.</w:delText></w:r></w:del>`));
});

test('deleting your own insertion removes it', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await startTracking(page);
  await select(page, ['Second paragraph.', 0]);
  await page.keyboard.type('New ');
  await page.keyboard.press('Backspace');
  await page.keyboard.press('Backspace');
  await expect(page.locator('.docx-page .tc-ins')).toHaveText('Ne');
  await expect(page.locator('.docx-page .tc-del')).toHaveCount(0);
  const { xml } = await download(page);
  expect(xml).toMatch(new RegExp(`<w:p>${rev('ins')}><w:r><w:t xml:space="preserve">Ne</w:t></w:r></w:ins><w:r><w:t xml:space="preserve">Second paragraph.</w:t>`));
  expect(xml).not.toContain('<w:del ');
});

test('typing over a selection records a deletion followed by an insertion', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await startTracking(page);
  await select(page, ['Second paragraph.', 0], ['Second paragraph.', 6]);
  await page.keyboard.type('Third');
  await expect(page.locator('.docx-page p').nth(1)).toHaveText('SecondThird paragraph.');
  const { xml } = await download(page);
  expect(xml).toMatch(new RegExp(`<w:p>${rev('del')}><w:r><w:delText xml:space="preserve">Second</w:delText></w:r></w:del>${rev('ins')}><w:r><w:t xml:space="preserve">Third</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> paragraph.</w:t>`));
});

test('Enter inserts a tracked paragraph break; Backspace deletes one without merging', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await startTracking(page);
  await select(page, ['Second paragraph.', 7]);
  await page.keyboard.press('Enter');
  await expect(page.locator('.docx-page p')).toHaveCount(3);
  await expect(page.locator('.docx-page p').nth(1)).toHaveAttribute('data-tc-mark', 'ins');
  // Backspace at the start of "Second ": the break of the first paragraph is deleted, not merged.
  await select(page, ['Second ', 0]);
  await page.keyboard.press('Backspace');
  await expect(page.locator('.docx-page p')).toHaveCount(3);
  await expect(page.locator('.docx-page p').first()).toHaveAttribute('data-tc-mark', 'del');
  let { xml } = await download(page);
  expect(xml).toMatch(new RegExp(`<w:p><w:pPr><w:rPr>${rev('del')}/></w:rPr></w:pPr><w:bookmarkStart w:id="7"`));
  expect(xml).toMatch(new RegExp(`<w:p><w:pPr><w:rPr>${rev('ins')}/></w:rPr></w:pPr><w:r><w:t xml:space="preserve">Second </w:t></w:r></w:p><w:p><w:r><w:t xml:space="preserve">paragraph.</w:t>`));

  // Backspace at the start of "paragraph.": that break was inserted here, so it goes.
  await select(page, ['paragraph.', 0]);
  await page.keyboard.press('Backspace');
  await expect(page.locator('.docx-page p')).toHaveCount(2);
  await expect(page.locator('.docx-page p').nth(1)).toHaveText('Second paragraph.');
  await expect(page.locator('.docx-page p').nth(1)).not.toHaveAttribute('data-tc-mark');
  ({ xml } = await download(page));
  expect(xml).toMatch(/<w:p><w:r><w:t xml:space="preserve">Second (<\/w:t>(<\/w:r><w:r>)?<w:t xml:space="preserve">)?paragraph\.<\/w:t><\/w:r><\/w:p>/);
});

test('Undo takes a tracked change back', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await startTracking(page);
  await select(page, ['Second paragraph.', 6], ['Second paragraph.', 16]);
  await page.keyboard.press('Backspace');
  await expect(page.locator('.docx-page .tc-del')).toHaveText(' paragraph');
  await page.keyboard.press('ControlOrMeta+z');
  await expect(page.locator('.docx-page .tc-del')).toHaveCount(0);
});

test('paste and Replace are tracked too', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await startTracking(page);
  await select(page, ['Second paragraph.', 0]);
  await page.evaluate(() => {
    const event = new ClipboardEvent('paste', { bubbles: true, cancelable: true });
    const dt = new DataTransfer();
    dt.setData('text/plain', 'Pasted ');
    Object.defineProperty(event, 'clipboardData', { value: dt });
    document.querySelector('.docx-page p:nth-child(2)').dispatchEvent(event);
  });
  await page.keyboard.press('ControlOrMeta+h');
  await page.getByRole('textbox', { name: 'Find' }).fill('fee');
  await page.getByRole('textbox', { name: 'Replace' }).fill('price');
  await page.getByRole('button', { name: 'Replace all' }).click();
  const { xml } = await download(page);
  expect(xml).toMatch(new RegExp(`<w:t xml:space="preserve">The </w:t></w:r>${rev('del')}><w:r><w:delText xml:space="preserve">fee</w:delText></w:r></w:del>${rev('ins')}><w:r><w:t xml:space="preserve">price</w:t></w:r></w:ins>`));
  expect(xml).toMatch(new RegExp(`${rev('ins')}><w:r><w:t xml:space="preserve">Pasted </w:t></w:r></w:ins><w:r><w:t xml:space="preserve">Second paragraph.</w:t>`));
});

test('a document with Track Changes on opens with it on; switching it off is saved', async ({ page }) => {
  const settings = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ${W_NS}><w:zoom w:percent="100"/><w:trackRevisions/><w:defaultTabStop w:val="720"/></w:settings>`;
  const docx = await buildDocx(DOC, { parts: { 'word/settings.xml': settings }, rels: `<Relationship Id="rIdS" Type="${REL}/settings" Target="settings.xml"/>` });
  await openInEditor(page, docx);
  await expect(page.getByRole('button', { name: 'Track changes' })).toHaveAttribute('aria-pressed', 'true');
  await expect(page.locator('.statusbar')).toContainText('Tracking changes as Author');
  await select(page, ['Second paragraph.', 0]);
  await page.keyboard.type('A ');
  let { zip, xml } = await download(page);
  expect(xml).toMatch(new RegExp(`${rev('ins', 'Author')}><w:r><w:t xml:space="preserve">A </w:t>`));
  expect(await settingsOf(zip)).toBe(settings); // unchanged: it was already on

  await page.keyboard.press('ControlOrMeta+Shift+E');
  await expect(status(page)).toHaveText('Track Changes is off: your edits change the text directly');
  await select(page, ['Second paragraph.', 0]);
  await page.keyboard.type('B');
  ({ zip, xml } = await download(page));
  // Typed right after the earlier insertion, but untracked: outside it.
  expect(xml).toMatch(/A <\/w:t><\/w:r><\/w:ins><w:r><w:t xml:space="preserve">B(<\/w:t><w:t xml:space="preserve">)?Second paragraph\.<\/w:t>/);
  const off = await settingsOf(zip); // rewritten: only the setting itself changes
  expect(off).not.toContain('trackRevisions');
  expect(off).toContain('<w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/></w:settings>');
});

test('others\' tracked changes stay protected, and comments survive a tracked deletion', async ({ page }) => {
  const comments = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments ${W_NS}><w:comment w:id="0" w:author="Bob"><w:p><w:r><w:t>Check.</w:t></w:r></w:p></w:comment></w:comments>`;
  const docx = await buildDocx('<w:p><w:r><w:t xml:space="preserve">Pay </w:t></w:r><w:ins w:id="1" w:author="Alice" w:date="2026-01-01T00:00:00Z"><w:r><w:t>promptly</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> now.</w:t></w:r></w:p>'
    + '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Commented text</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>',
  { parts: { 'word/comments.xml': comments }, rels: `<Relationship Id="rIdC" Type="${REL}/comments" Target="comments.xml"/>` });
  await openInEditor(page, docx);
  await startTracking(page);
  await select(page, ['Pay ', 0], [' now.', 2]);
  await page.keyboard.press('Backspace');
  await expect(status(page)).toContainText('a tracked insertion by Alice');
  await select(page, ['Commented text', 0], ['Commented text', 14]);
  await page.keyboard.press('Backspace');
  await page.keyboard.press('Backspace');
  const { zip, xml } = await download(page);
  expect(xml).toMatch(new RegExp(`<w:commentRangeStart w:id="0"/>${rev('del')}><w:r><w:delText xml:space="preserve">Commented text</w:delText></w:r></w:del><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>`));
  expect(await zip.file('word/comments.xml').async('string')).toBe(comments);
});

test('a saved file\'s tracked changes reopen as tracked changes', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await startTracking(page);
  await select(page, ['Second paragraph.', 0], ['Second paragraph.', 7]);
  await page.keyboard.type('Next ');
  const saved = await (await import('node:fs')).promises.readFile(await downloadFile(page));
  await openInEditor(page, saved);
  await expect(page.locator('.docx-page .inserted')).toHaveAttribute('title', /^Inserted by Rita on \d{4}-/);
  await expect(page.locator('.docx-page .deleted')).toHaveAttribute('title', /^Deleted by Rita on \d{4}-/);
  expect(await JSZip.loadAsync(saved).then((z) => z.file('word/settings.xml')?.async('string'))).toContain('trackRevisions');
});
