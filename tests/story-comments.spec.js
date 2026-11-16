// Comments anchored in a header, footer or note can be replied to, edited, resolved and deleted, as
// those in the text: a reply is written next to its comment in that part (inside its range), a
// deleted comment's marks go from that part, and Undo covers it all.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, W_NS } from './helpers.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const HEADER = `${XML}<w:hdr ${W_NS}><w:p><w:commentRangeStart w:id="5"/><w:r><w:t>Header text</w:t></w:r><w:commentRangeEnd w:id="5"/><w:r><w:commentReference w:id="5"/></w:r></w:p></w:hdr>`;
const FOOTNOTES = `${XML}<w:footnotes ${W_NS}><w:footnote w:id="1"><w:p><w:commentRangeStart w:id="6"/><w:r><w:t>A note.</w:t></w:r><w:commentRangeEnd w:id="6"/><w:r><w:commentReference w:id="6"/></w:r></w:p></w:footnote></w:footnotes>`;
const COMMENTS = `${XML}<w:comments ${W_NS}>`
  + '<w:comment w:id="5" w:author="Bob" w:date="2026-01-01T00:00:00Z"><w:p w14:paraId="00000005"><w:r><w:t>About the header.</w:t></w:r></w:p></w:comment>'
  + '<w:comment w:id="6" w:author="Cy" w:date="2026-01-02T00:00:00Z"><w:p w14:paraId="00000006"><w:r><w:t>About the note.</w:t></w:r></w:p></w:comment></w:comments>';
const EXTENDED = `${XML}<w15:commentsEx xmlns:w15="${W15}"><w15:commentEx w15:paraId="00000005" w15:done="0"/><w15:commentEx w15:paraId="00000006" w15:done="0"/></w15:commentsEx>`;
const OPTS = {
  parts: { 'word/header1.xml': HEADER, 'word/footnotes.xml': FOOTNOTES, 'word/comments.xml': COMMENTS, 'word/commentsExtended.xml': EXTENDED },
  rels: `<Relationship Id="rIdH" Type="${REL}/header" Target="header1.xml"/><Relationship Id="rIdF" Type="${REL}/footnotes" Target="footnotes.xml"/>`
    + `<Relationship Id="rIdC" Type="${REL}/comments" Target="comments.xml"/><Relationship Id="rIdCx" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>`,
  sectPr: '<w:headerReference w:type="default" r:id="rIdH"/>',
};
const BODY = '<w:p><w:r><w:t xml:space="preserve">Body</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>';

const panel = (page) => page.getByRole('complementary', { name: 'Comments' });
const thread = (page, text) => panel(page).locator('.thread', { hasText: text });
const status = (page) => page.locator('.statusbar [role=status]');

test.beforeEach(async ({ page }) => page.setViewportSize({ width: 1280, height: 800 }));

test('a comment in a header says where it is, and can be replied to: the reply goes into the header', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY, OPTS));
  await expect(thread(page, 'About the header.').locator('.tag.muted')).toHaveText('In a header');
  await expect(thread(page, 'About the note.').locator('.tag.muted')).toHaveText('In a footnote');
  await thread(page, 'About the header.').getByRole('button', { name: 'Reply' }).click();
  const form = thread(page, 'About the header.').getByRole('form', { name: 'Reply' });
  await form.getByRole('textbox', { name: 'Your name' }).fill('Rita');
  await form.getByRole('textbox', { name: 'Reply text' }).fill('Fixed the header.');
  await form.getByRole('button', { name: 'Reply' }).click();
  await expect(status(page)).toHaveText('Reply added');
  await expect(thread(page, 'About the header.').locator('.reply')).toContainText('Fixed the header.');

  const { zip } = await download(page);
  const header = await zip.file('word/header1.xml').async('string');
  const id = header.match(/<w:commentRangeStart w:id="5"\/><w:commentRangeStart w:id="(\d+)"\/>/)[1];
  expect(header).toContain(`<w:commentRangeEnd w:id="5"/><w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="5"/></w:r><w:r><w:commentReference w:id="${id}"/></w:r>`);
  const comments = await zip.file('word/comments.xml').async('string');
  const paraId = comments.match(new RegExp(`<w:comment w:id="${id}" w:author="Rita"[^>]*><w:p w14:paraId="([0-9A-F]{8})">`))[1];
  const extended = await zip.file('word/commentsExtended.xml').async('string');
  expect(extended).toContain(`<w15:commentEx w15:paraId="${paraId}" w15:done="0" w15:paraIdParent="00000005"/>`);
});

test('editing and resolving a comment in a footnote; Undo takes them back', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY, OPTS));
  await thread(page, 'About the note.').getByRole('button', { name: 'Edit' }).click();
  const form = panel(page).getByRole('form', { name: 'Edit comment' });
  await form.getByRole('textbox', { name: 'Comment text' }).fill('About the footnote.');
  await form.getByRole('button', { name: 'Save' }).click();
  await expect(status(page)).toHaveText('Comment edited');
  await thread(page, 'About the footnote.').getByRole('button', { name: 'Resolve' }).click();
  await expect(thread(page, 'About the footnote.').locator('.tag.resolved')).toHaveText('Resolved');
  let { zip } = await download(page);
  expect(await zip.file('word/comments.xml').async('string')).toContain('<w:comment w:id="6" w:author="Cy" w:date="2026-01-02T00:00:00Z"><w:p w14:paraId="00000006"><w:r><w:t xml:space="preserve">About the footnote.</w:t></w:r></w:p></w:comment>');
  expect(await zip.file('word/commentsExtended.xml').async('string')).toContain('<w15:commentEx w15:paraId="00000006" w15:done="1"/>');

  await page.locator('.docx-page p').first().click();
  await page.keyboard.press('ControlOrMeta+z');
  await page.keyboard.press('ControlOrMeta+z');
  await expect(thread(page, 'About the note.')).toBeVisible();
  await expect(thread(page, 'About the note.').locator('.tag.resolved')).toHaveCount(0);
  ({ zip } = await download(page));
  expect(await zip.file('word/comments.xml').async('string')).toBe(COMMENTS); // as it was, byte for byte
  expect(await zip.file('word/commentsExtended.xml').async('string')).toBe(EXTENDED);
});

test('deleting a comment in a header takes its marks out of the header; Undo puts it back as it was', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY, OPTS));
  await thread(page, 'About the header.').locator('.thread-actions').getByRole('button', { name: 'Delete' }).click();
  await expect(status(page)).toHaveText('Comment deleted – Undo (Ctrl+Z) brings it back');
  await expect(thread(page, 'About the header.')).toHaveCount(0);
  await expect(page.locator('.docx-hf [data-comment-ref="5"]')).toBeHidden();
  let { zip } = await download(page);
  expect(await zip.file('word/header1.xml').async('string')).not.toContain('comment');
  expect(await zip.file('word/comments.xml').async('string')).not.toContain('About the header.');

  await page.locator('.docx-page p').first().click();
  await page.keyboard.press('ControlOrMeta+z');
  await expect(thread(page, 'About the header.')).toBeVisible();
  ({ zip } = await download(page));
  expect(await zip.file('word/header1.xml').async('string')).toBe(HEADER);
  expect(await zip.file('word/comments.xml').async('string')).toBe(COMMENTS);
});

test('the review lists changes to comments in headers and notes', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY, OPTS));
  await thread(page, 'About the note.').getByRole('button', { name: 'Resolve' }).click();
  await thread(page, 'About the header.').locator('.thread-actions').getByRole('button', { name: 'Delete' }).click();
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  const notes = page.getByRole('dialog', { name: 'Review your changes' }).locator('.notes');
  await expect(notes).toContainText('Comment by Cy resolved.');
  await expect(notes).toContainText('Comment by Bob removed');
});
