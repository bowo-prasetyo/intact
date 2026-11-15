// Deleting text never silently takes tracked changes or comments with it. Edits that would remove a
// tracked change are refused, with a reason; deleting the text a comment is attached to deletes the
// comment completely (as in Word), visibly, and Undo brings it back.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, canonicalize, documentXmlOf, W_NS } from './helpers.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const status = (page) => page.locator('.statusbar [role=status]');

/** Selects from `from` (text, offset in it) to `to` in the editable page, or places the caret. */
async function select(page, from, to = from) {
  await page.evaluate(([[a, ao], [b, bo]]) => {
    const root = document.querySelector('.docx-page');
    const find = (t) => {
      const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
      for (let n = w.nextNode(); n; n = w.nextNode()) if (n.data.includes(t)) return n;
      throw new Error(`no text "${t}"`);
    };
    const s = find(a);
    const e = find(b);
    const range = document.createRange();
    range.setStart(s, s.data.indexOf(a) + ao);
    range.setEnd(e, e.data.indexOf(b) + bo);
    root.focus();
    getSelection().removeAllRanges();
    getSelection().addRange(range);
  }, [from, to]);
}

// ------------------------------------------------------------------ tracked changes

const TRACKED = `<w:p><w:r><w:t xml:space="preserve">The client </w:t></w:r><w:ins w:id="1" w:author="Alice" w:date="2026-01-01T00:00:00Z"><w:r><w:t>shall promptly</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> pay </w:t></w:r><w:del w:id="2" w:author="Bob" w:date="2026-01-02T00:00:00Z"><w:r><w:delText>all</w:delText></w:r></w:del><w:r><w:t xml:space="preserve"> invoices.</w:t></w:r></w:p>
<w:p><w:pPr><w:rPr><w:ins w:id="3" w:author="Carol" w:date="2026-01-03T00:00:00Z"/></w:rPr></w:pPr><w:r><w:t>Split by Carol.</w:t></w:r></w:p>
<w:p><w:r><w:t>Plain paragraph.</w:t></w:r></w:p>
<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:trPr><w:ins w:id="4" w:author="Dan" w:date="2026-01-04T00:00:00Z"/></w:trPr><w:tc><w:p><w:r><w:t>Row by Dan</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
<w:p><w:r><w:t>After the table.</w:t></w:r></w:p>`;

/** The document must come back exactly as it was opened. */
async function expectUnchanged(page, docx) {
  const { xml } = await download(page);
  expect(await canonicalize(page, xml)).toEqual(await canonicalize(page, await documentXmlOf(docx)));
}

test('Backspace and Delete next to a tracked change leave it alone and say why', async ({ page }) => {
  const docx = await buildDocx(TRACKED);
  await openInEditor(page, docx);
  await select(page, [' pay ', 0]);
  await page.keyboard.press('Backspace');
  await expect(status(page)).toHaveText('Not deleted: that would remove a tracked insertion by Alice. Accept or reject it first (on the toolbar), or select only the text around it.');
  await select(page, [' invoices.', 0]);
  await page.keyboard.press('Backspace');
  await expect(status(page)).toContainText('a tracked deletion by Bob');
  await select(page, ['The client ', 11]);
  await page.keyboard.press('Delete');
  await page.keyboard.press('ControlOrMeta+Delete');
  await expect(status(page)).toContainText('a tracked insertion by Alice');
  await expectUnchanged(page, docx);
});

test('typing over, cutting or pasting over a selection with a tracked change is refused', async ({ page }) => {
  const docx = await buildDocx(TRACKED);
  await openInEditor(page, docx);
  await select(page, ['The client ', 4], [' pay ', 2]);
  await page.keyboard.type('x');
  await expect(status(page)).toContainText('a tracked insertion by Alice');
  await select(page, [' pay ', 1], [' invoices.', 2]);
  await page.keyboard.press('ControlOrMeta+x');
  await expect(status(page)).toContainText('a tracked deletion by Bob');
  await select(page, ['The client ', 4], [' pay ', 2]);
  await page.keyboard.press('Enter');
  await page.keyboard.press('Tab');
  await page.evaluate(() => {
    const event = new ClipboardEvent('paste', { bubbles: true, cancelable: true });
    const dt = new DataTransfer();
    dt.setData('text/plain', 'pasted');
    Object.defineProperty(event, 'clipboardData', { value: dt });
    document.querySelector('.docx-page p').dispatchEvent(event);
  });
  await expect(page.locator('.docx-page')).not.toContainText('pasted');
  await expectUnchanged(page, docx);
});

test('deleting across paragraphs that hold a tracked change is refused', async ({ page }) => {
  const docx = await buildDocx(TRACKED);
  await openInEditor(page, docx);
  await select(page, ['The client ', 0], ['Plain paragraph.', 5]);
  await page.keyboard.press('Delete');
  await expect(status(page)).toContainText('a tracked insertion by Alice');
  await page.keyboard.press('ControlOrMeta+a');
  await page.keyboard.press('Backspace');
  await expect(status(page)).toContainText('Not deleted');
  await expectUnchanged(page, docx);
});

test('a paragraph whose break is a tracked change can\'t be merged away', async ({ page }) => {
  const docx = await buildDocx(TRACKED);
  await openInEditor(page, docx);
  await select(page, ['Plain paragraph.', 0]);
  await page.keyboard.press('Backspace'); // would merge into Carol's paragraph
  await expect(status(page)).toHaveText(/^Not deleted: that would remove a tracked paragraph break by Carol\./);
  await select(page, ['Split by Carol.', 15]);
  await page.keyboard.press('Delete');
  await expect(status(page)).toContainText('a tracked paragraph break by Carol');
  await expectUnchanged(page, docx);

  // Inside the paragraph, editing is as usual.
  await select(page, ['Split by Carol.', 0], ['Split by Carol.', 5]);
  await page.keyboard.type('Added');
  const { xml } = await download(page);
  expect(xml).toContain('<w:ins w:id="3" w:author="Carol" w:date="2026-01-03T00:00:00Z"/></w:rPr></w:pPr><w:r><w:t xml:space="preserve">Added by Carol.</w:t>');
});

test('a table with tracked rows can\'t be deleted with the text around it', async ({ page }) => {
  const docx = await buildDocx(TRACKED);
  await openInEditor(page, docx);
  await select(page, ['Plain paragraph.', 6], ['After the table.', 5]);
  await page.keyboard.press('Backspace');
  await expect(status(page)).toContainText('a table with tracked changes by Dan');
  await expectUnchanged(page, docx);

  // Editing a cell of that table is fine.
  await select(page, ['Row by Dan', 10]);
  await page.keyboard.type('!');
  const { xml } = await download(page);
  expect(xml).toContain('<w:t xml:space="preserve">Row by Dan!</w:t>');
});

test('text around a tracked change can still be deleted', async ({ page }) => {
  await openInEditor(page, await buildDocx(TRACKED));
  await select(page, ['The client ', 0], ['The client ', 4]);
  await page.keyboard.press('Delete');
  await select(page, [' invoices.', 1], [' invoices.', 10]);
  await page.keyboard.press('Backspace');
  await page.keyboard.type('bills.');
  await expect(status(page)).not.toContainText('Not deleted');
  const { xml } = await download(page);
  expect(xml).toContain('<w:t xml:space="preserve">client </w:t></w:r><w:ins w:id="1" w:author="Alice" w:date="2026-01-01T00:00:00Z"><w:r><w:t>shall promptly</w:t></w:r></w:ins>');
  expect(xml).toContain('<w:del w:id="2" w:author="Bob" w:date="2026-01-02T00:00:00Z"><w:r><w:delText>all</w:delText></w:r></w:del><w:r><w:t xml:space="preserve"> bills.</w:t>');
});

// ------------------------------------------------------------------ comments

const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const W16CID = 'http://schemas.microsoft.com/office/word/2016/wordml/cid';
const W16CEX = 'http://schemas.microsoft.com/office/word/2018/wordml/cex';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

// Comment 0 (Ann) has a reply, 1 (Bob); comment 2 (Ann) is on another paragraph.
const PARTS = {
  'word/comments.xml': `${XML}<w:comments ${W_NS}>`
    + '<w:comment w:id="0" w:author="Ann Lee" w:date="2026-01-05T10:00:00Z"><w:p w14:paraId="00000001"><w:r><w:t>Check the amount.</w:t></w:r></w:p></w:comment>'
    + '<w:comment w:id="1" w:author="Bob" w:date="2026-01-06T10:00:00Z"><w:p w14:paraId="00000002"><w:r><w:t>Confirmed with finance.</w:t></w:r></w:p></w:comment>'
    + '<w:comment w:id="2" w:author="Ann Lee" w:date="2026-01-07T10:00:00Z"><w:p w14:paraId="00000003"><w:r><w:t>Typo here?</w:t></w:r></w:p></w:comment>'
    + '</w:comments>',
  'word/commentsExtended.xml': `${XML}<w15:commentsEx xmlns:w15="${W15}">`
    + '<w15:commentEx w15:paraId="00000001" w15:done="0"/><w15:commentEx w15:paraId="00000002" w15:paraIdParent="00000001" w15:done="0"/><w15:commentEx w15:paraId="00000003" w15:done="0"/>'
    + '</w15:commentsEx>',
  'word/commentsIds.xml': `${XML}<w16cid:commentsIds xmlns:w16cid="${W16CID}">`
    + '<w16cid:commentId w16cid:paraId="00000001" w16cid:durableId="1A000001"/><w16cid:commentId w16cid:paraId="00000002" w16cid:durableId="1A000002"/><w16cid:commentId w16cid:paraId="00000003" w16cid:durableId="1A000003"/>'
    + '</w16cid:commentsIds>',
  'word/commentsExtensible.xml': `${XML}<w16cex:commentsExtensible xmlns:w16cex="${W16CEX}">`
    + '<w16cex:commentExtensible w16cex:durableId="1A000001" w16cex:dateUtc="2026-01-05T10:00:00Z"/><w16cex:commentExtensible w16cex:durableId="1A000002" w16cex:dateUtc="2026-01-06T10:00:00Z"/><w16cex:commentExtensible w16cex:durableId="1A000003" w16cex:dateUtc="2026-01-07T10:00:00Z"/>'
    + '</w16cex:commentsExtensible>',
};
const COMMENTS = {
  parts: PARTS,
  rels: `<Relationship Id="rIdC" Type="${REL}/comments" Target="comments.xml"/>`
    + '<Relationship Id="rIdCx" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>'
    + '<Relationship Id="rIdCi" Type="http://schemas.microsoft.com/office/2016/09/relationships/commentsIds" Target="commentsIds.xml"/>'
    + '<Relationship Id="rIdCe" Type="http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible" Target="commentsExtensible.xml"/>',
};
const COMMENTED = '<w:p><w:r><w:t xml:space="preserve">The fee is </w:t></w:r><w:commentRangeStart w:id="0"/><w:commentRangeStart w:id="1"/><w:r><w:t>EUR 1,200</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r><w:r><w:t xml:space="preserve"> per month.</w:t></w:r></w:p>'
  + '<w:p><w:commentRangeStart w:id="2"/><w:r><w:t>Recieved</w:t></w:r><w:commentRangeEnd w:id="2"/><w:r><w:commentReference w:id="2"/></w:r><w:r><w:t xml:space="preserve"> on time.</w:t></w:r></w:p>'
  + '<w:p><w:r><w:t>Last paragraph.</w:t></w:r></w:p>';

const parts = async (zip) => Object.fromEntries(await Promise.all(Object.keys(PARTS).map(async (p) => [p, await zip.file(p).async('string')])));

test('deleting commented text deletes its comment and replies from every comment part', async ({ page }) => {
  await page.setViewportSize({ width: 1280, height: 800 });
  await openInEditor(page, await buildDocx(COMMENTED, COMMENTS));
  const threads = page.getByRole('complementary', { name: 'Comments' }).locator('.thread');
  await expect(threads).toHaveCount(2);

  await select(page, ['The fee is ', 0], ['Recieved', 0]);
  await page.keyboard.press('Backspace');
  await expect(status(page)).toHaveText('Deleted the comment by Ann Lee and its reply with the text they were attached to – Undo (Ctrl+Z) brings them back');
  await expect(threads).toHaveCount(1);
  await expect(threads.first()).toContainText('Typo here?');
  await expect(page.getByTitle('Comments (1)')).toBeVisible();

  const { zip, xml } = await download(page);
  await expect(status(page)).toHaveText(/ – 2 comments removed with their text$/);
  expect(xml).not.toMatch(/w:id="[01]"/);
  // Comment 2's range started where the deletion ended: its start marker is kept.
  expect(xml).toContain('<w:commentRangeStart w:id="2"/><w:r><w:t xml:space="preserve">Recieved</w:t></w:r><w:commentRangeEnd w:id="2"/><w:r><w:commentReference w:id="2"/></w:r>');
  const saved = await parts(zip);
  expect(saved['word/comments.xml']).not.toMatch(/w:id="[01]"/);
  expect(saved['word/comments.xml']).toContain('Typo here?');
  expect(saved['word/commentsExtended.xml']).not.toMatch(/0000000[12]/);
  expect(saved['word/commentsExtended.xml']).toContain('w15:paraId="00000003"');
  expect(saved['word/commentsIds.xml']).not.toMatch(/0000000[12]/);
  expect(saved['word/commentsIds.xml']).toContain('w16cid:durableId="1A000003"');
  expect(saved['word/commentsExtensible.xml']).not.toMatch(/1A00000[12]/);
  expect(saved['word/commentsExtensible.xml']).toContain('w16cex:durableId="1A000003"');
});

test('Undo brings a deleted comment back, and the comment parts come back unchanged', async ({ page }) => {
  await page.setViewportSize({ width: 1280, height: 800 });
  await openInEditor(page, await buildDocx(COMMENTED, COMMENTS));
  await select(page, ['Recieved', 0], [' on time.', 9]);
  await page.keyboard.press('Delete');
  await expect(status(page)).toHaveText('Deleted the comment by Ann Lee with the text it was attached to – Undo (Ctrl+Z) brings it back');
  const threads = page.getByRole('complementary', { name: 'Comments' }).locator('.thread');
  await expect(threads).toHaveCount(1);
  let saved = await parts((await download(page)).zip);
  expect(saved['word/comments.xml']).not.toContain('Typo here?');

  await page.keyboard.press('ControlOrMeta+z');
  await expect(threads).toHaveCount(2);
  saved = await parts((await download(page)).zip);
  expect(saved).toEqual(PARTS); // byte for byte, as opened
});

test('editing elsewhere leaves the comment parts byte for byte', async ({ page }) => {
  const docx = await buildDocx(COMMENTED, COMMENTS);
  await openInEditor(page, docx);
  await select(page, ['Last paragraph.', 15]);
  await page.keyboard.type(' More.');
  const { zip } = await download(page);
  expect(await parts(zip)).toEqual(PARTS);
});

test('deleting one end of a comment range with the text around it keeps the range', async ({ page }) => {
  await openInEditor(page, await buildDocx(COMMENTED, COMMENTS));
  // From the end of the first paragraph into the commented word: the paragraphs merge and the
  // comment's start marker, deleted with the text, is put back where the deletion happened.
  await select(page, [' per month.', 5], ['Recieved', 3]);
  await page.keyboard.press('Delete');
  await expect(page.locator('.docx-page p').first()).toHaveText(/^The fee is EUR 1,200 per ieved on time\.$/);
  await expect(page.locator('[data-comment-start="2"]')).toHaveCount(1);
  const { zip, xml } = await download(page);
  expect(xml).toContain('<w:t xml:space="preserve"> per </w:t></w:r><w:commentRangeStart w:id="2"/><w:r><w:t xml:space="preserve">ieved</w:t></w:r><w:commentRangeEnd w:id="2"/><w:r><w:commentReference w:id="2"/></w:r>');
  expect(await parts(zip)).toEqual(PARTS);
});

test('deleting a comment\'s mark deletes the comment, and no half of its range is left', async ({ page }) => {
  await openInEditor(page, await buildDocx(COMMENTED, COMMENTS));
  // From inside "EUR 1,200" to the end of the paragraph: both marks and range ends go, the starts stay.
  await select(page, ['EUR 1,200', 4], [' per month.', 11]);
  await page.keyboard.press('Backspace');
  await expect(status(page)).toHaveText(/^Deleted the comment by Ann Lee and its reply /);
  const { zip, xml } = await download(page);
  expect(xml).toContain('<w:p><w:r><w:t xml:space="preserve">The fee is </w:t></w:r><w:r><w:t xml:space="preserve">EUR </w:t></w:r></w:p>');
  expect(xml).not.toMatch(/w:id="[01]"/);
  expect((await parts(zip))['word/comments.xml']).not.toMatch(/w:id="[01]"/);
});

test('Backspace next to a comment\'s range marker deletes text, not the marker', async ({ page }) => {
  await openInEditor(page, await buildDocx(COMMENTED, COMMENTS));
  await select(page, ['EUR 1,200', 0]); // just after the start markers
  await page.keyboard.press('Backspace');
  await page.keyboard.press('Backspace');
  await select(page, [' per month.', 0]); // just after the end marker and the comment mark
  await page.keyboard.press('Delete');
  const { zip, xml } = await download(page);
  expect(xml).toContain('<w:t xml:space="preserve">The fee i</w:t></w:r><w:commentRangeStart w:id="0"/><w:commentRangeStart w:id="1"/><w:r><w:t xml:space="preserve">EUR 1,200</w:t></w:r><w:commentRangeEnd w:id="0"/>');
  expect(xml).toContain('<w:commentReference w:id="1"/></w:r><w:r><w:t xml:space="preserve">per month.</w:t>');
  expect(await parts(zip)).toEqual(PARTS);
});
