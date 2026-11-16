// Replying to, editing, resolving and deleting comments that are already in the document, as Word
// writes them: replies are comments linked to their parent in commentsExtended.xml (created if need
// be, with paragraph ids added where old comments lack them); resolved is w15:done there.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, W_NS } from './helpers.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const panel = (page) => page.getByRole('complementary', { name: 'Comments' });
const thread = (page, text) => panel(page).locator('.thread', { hasText: text });
const status = (page) => page.locator('.statusbar [role=status]');

const BODY = '<w:p><w:r><w:t xml:space="preserve">The fee is </w:t></w:r><w:commentRangeStart w:id="0"/><w:r><w:t>EUR 1,200</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r><w:r><w:t xml:space="preserve"> per month.</w:t></w:r></w:p>'
  + '<w:p><w:r><w:t>Second paragraph.</w:t></w:r></w:p>';
// With paragraph ids and commentsExtended.xml, as Word 2013 and later write them.
const COMMENTS = `${XML}<w:comments ${W_NS}><w:comment w:id="0" w:author="Ann Lee" w:date="2026-01-05T10:00:00Z" w:initials="AL"><w:p w14:paraId="00000001"><w:r><w:annotationRef/></w:r><w:r><w:t>Check the amount.</w:t></w:r></w:p></w:comment></w:comments>`;
const EXTENDED = `${XML}<w15:commentsEx xmlns:w15="${W15}"><w15:commentEx w15:paraId="00000001" w15:done="0"/></w15:commentsEx>`;
const MODERN = {
  parts: { 'word/comments.xml': COMMENTS, 'word/commentsExtended.xml': EXTENDED },
  rels: `<Relationship Id="rIdC" Type="${REL}/comments" Target="comments.xml"/><Relationship Id="rIdCx" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>`,
};
// An older file: no paragraph ids, no commentsExtended.xml.
const OLD = { parts: { 'word/comments.xml': COMMENTS.replace(' w14:paraId="00000001"', '') }, rels: `<Relationship Id="rIdC" Type="${REL}/comments" Target="comments.xml"/>` };

async function reply(page, threadText, text, name = 'Rita') {
  await thread(page, threadText).getByRole('button', { name: 'Reply' }).click();
  const form = thread(page, threadText).getByRole('form', { name: 'Reply' });
  if (name) await form.getByRole('textbox', { name: 'Your name' }).fill(name);
  await form.getByRole('textbox', { name: 'Reply text' }).fill(text);
  await form.getByRole('button', { name: 'Reply' }).click();
}

test.beforeEach(async ({ page }) => page.setViewportSize({ width: 1280, height: 800 }));

test('a reply is a real Word reply: its own comment, inside the parent\'s range, linked to it', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY, MODERN));
  await reply(page, 'Check the amount.', 'Confirmed with finance.');
  await expect(status(page)).toHaveText('Reply added');
  await expect(thread(page, 'Check the amount.').locator('.reply')).toContainText('Confirmed with finance.');
  await expect(thread(page, 'Check the amount.').locator('.reply')).toContainText('Rita');
  await expect(page.locator('.docx-page .comment-ref')).toHaveCount(1); // a reply's mark is hidden

  const { zip, xml } = await download(page);
  const id = xml.match(/<w:commentRangeStart w:id="0"\/><w:commentRangeStart w:id="(\d+)"\/>/)[1];
  expect(xml).toContain(`<w:commentRangeEnd w:id="0"/><w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="0"/></w:r><w:r><w:commentReference w:id="${id}"/></w:r>`);
  const comments = await zip.file('word/comments.xml').async('string');
  const paraId = comments.match(new RegExp(`<w:comment w:id="${id}" w:author="Rita"[^>]*><w:p w14:paraId="([0-9A-F]{8})"><w:r><w:annotationRef/></w:r><w:r><w:t xml:space="preserve">Confirmed with finance.</w:t></w:r></w:p></w:comment>`))?.[1];
  expect(paraId).toBeTruthy();
  expect(comments).toContain('<w:comment w:id="0" w:author="Ann Lee" w:date="2026-01-05T10:00:00Z" w:initials="AL"><w:p w14:paraId="00000001"><w:r><w:annotationRef/></w:r><w:r><w:t>Check the amount.</w:t></w:r></w:p></w:comment>');
  const extended = await zip.file('word/commentsExtended.xml').async('string');
  expect(extended).toContain(`<w15:commentEx w15:paraId="00000001" w15:done="0"/><w15:commentEx w15:paraId="${paraId}" w15:done="0" w15:paraIdParent="00000001"/>`);
});

test('replying in an older file adds paragraph ids and creates commentsExtended.xml', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY, OLD));
  await reply(page, 'Check the amount.', 'Agreed.');
  const { zip } = await download(page);
  const comments = await zip.file('word/comments.xml').async('string');
  const parentId = comments.match(/<w:comment w:id="0"[^>]*><w:p w14:paraId="([0-9A-F]{8})">/)[1];
  const extended = await zip.file('word/commentsExtended.xml').async('string');
  expect(extended).toMatch(new RegExp(`<w15:commentEx w15:paraId="${parentId}" w15:done="0"/><w15:commentEx w15:paraId="[0-9A-F]{8}" w15:done="0" w15:paraIdParent="${parentId}"/>`));
  expect(await zip.file('word/_rels/document.xml.rels').async('string')).toContain('relationships/commentsExtended" Target="commentsExtended.xml"');
  expect(await zip.file('[Content_Types].xml').async('string')).toContain('PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"');
});

test('editing a comment rewrites its text and keeps its id; Undo takes the edit back', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY, MODERN));
  await thread(page, 'Check the amount.').getByRole('button', { name: 'Edit' }).click();
  const form = panel(page).getByRole('form', { name: 'Edit comment' });
  await form.getByRole('textbox', { name: 'Comment text' }).fill('Check the amount\nand the currency.');
  await form.getByRole('button', { name: 'Save' }).click();
  await expect(status(page)).toHaveText('Comment edited');
  await expect(panel(page).locator('.thread .text')).toHaveText(['Check the amount', 'and the currency.']);
  let { zip } = await download(page);
  expect(await zip.file('word/comments.xml').async('string')).toContain('<w:comment w:id="0" w:author="Ann Lee" w:date="2026-01-05T10:00:00Z" w:initials="AL"><w:p><w:r><w:annotationRef/></w:r><w:r><w:t xml:space="preserve">Check the amount</w:t></w:r></w:p><w:p w14:paraId="00000001"><w:r><w:t xml:space="preserve">and the currency.</w:t></w:r></w:p></w:comment>');

  await page.locator('.docx-page p').nth(1).click();
  await page.keyboard.press('ControlOrMeta+z');
  await expect(panel(page).locator('.thread .text')).toHaveText(['Check the amount.']);
  ({ zip } = await download(page));
  expect(await zip.file('word/comments.xml').async('string')).toBe(COMMENTS); // as it was, byte for byte
});

test('resolving and reopening a thread: w15:done, and the review says so', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY, MODERN));
  await thread(page, 'Check the amount.').getByRole('button', { name: 'Resolve' }).click();
  await expect(thread(page, 'Check the amount.').locator('.tag.resolved')).toHaveText('Resolved');
  await expect(thread(page, 'Check the amount.').getByRole('button', { name: 'Reopen' })).toBeVisible();
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  const review = page.getByRole('dialog', { name: 'Review your changes' });
  await expect(review.locator('.notes')).toContainText('Comment by Ann Lee resolved.');
  await review.getByRole('button', { name: 'Keep editing' }).click();
  let { zip } = await download(page);
  expect(await zip.file('word/commentsExtended.xml').async('string')).toContain('<w15:commentEx w15:paraId="00000001" w15:done="1"/>');

  await thread(page, 'Check the amount.').getByRole('button', { name: 'Reopen' }).click();
  ({ zip } = await download(page));
  expect(await zip.file('word/commentsExtended.xml').async('string')).toBe(EXTENDED); // back as it was
});

test('deleting a thread deletes its comment and replies; a comment in a header stays', async ({ page }) => {
  const header = `${XML}<w:hdr ${W_NS}><w:p><w:commentRangeStart w:id="5"/><w:r><w:t>Header text</w:t></w:r><w:commentRangeEnd w:id="5"/><w:r><w:commentReference w:id="5"/></w:r></w:p></w:hdr>`;
  const comments = COMMENTS.replace('</w:comments>', '<w:comment w:id="5" w:author="Bob"><w:p><w:r><w:t>About the header.</w:t></w:r></w:p></w:comment></w:comments>');
  await openInEditor(page, await buildDocx(BODY, {
    parts: { ...MODERN.parts, 'word/comments.xml': comments, 'word/header1.xml': header },
    rels: `${MODERN.rels}<Relationship Id="rIdH" Type="${REL}/header" Target="header1.xml"/>`,
    sectPr: '<w:headerReference w:type="default" r:id="rIdH"/>',
  }));
  await expect(thread(page, 'About the header.').locator('.tag.muted')).toHaveText('In a header');

  await reply(page, 'Check the amount.', 'A reply.');
  await thread(page, 'Check the amount.').locator('.thread-actions').getByRole('button', { name: 'Delete' }).click();
  await expect(panel(page).locator('.thread')).toHaveCount(1);
  const { zip, xml } = await download(page);
  expect(xml).not.toContain('commentRangeStart');
  const saved = await zip.file('word/comments.xml').async('string');
  expect(saved).not.toContain('Check the amount.');
  expect(saved).not.toContain('A reply.');
  expect(saved).toContain('About the header.');
});
