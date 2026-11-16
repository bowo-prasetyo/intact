// Adding comments, as in Word: on the selection (or the word at the caret), with your name; written
// to comments.xml (created if need be) next to the comments already there, which stay as they were.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, W_NS } from './helpers.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const DOC = '<w:p><w:r><w:t xml:space="preserve">The fee is due within thirty days.</w:t></w:r></w:p><w:p><w:r><w:t>Second paragraph.</w:t></w:r></w:p>';
const panel = (page) => page.getByRole('complementary', { name: 'Comments' });
const composer = (page) => panel(page).getByRole('form', { name: 'New comment' });

async function select(page, text, offset = 0, length = text.length - offset) {
  await page.evaluate(([t, o, l]) => {
    const root = document.querySelector('.docx-page');
    const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    for (let n = w.nextNode(); n; n = w.nextNode()) {
      const i = n.data.indexOf(t);
      if (i < 0) continue;
      const range = document.createRange();
      range.setStart(n, i + o);
      range.setEnd(n, i + o + l);
      root.focus();
      getSelection().removeAllRanges();
      getSelection().addRange(range);
      return;
    }
    throw new Error(`no text "${t}"`);
  }, [text, offset, length]);
}

async function comment(page, text, name = null) {
  if (name) await composer(page).getByRole('textbox', { name: 'Your name' }).fill(name);
  await composer(page).getByRole('textbox', { name: 'New comment' }).fill(text);
  await composer(page).getByRole('button', { name: 'Comment' }).click();
}

test('a comment on the selected text, in a document without comments: comments.xml is created', async ({ page }) => {
  await page.setViewportSize({ width: 1280, height: 800 });
  await openInEditor(page, await buildDocx(DOC));
  await select(page, 'thirty days');
  await page.getByTitle('New comment (Ctrl+Alt+M)').click();
  await expect(composer(page).getByRole('textbox', { name: 'Your name' })).toBeFocused(); // asked once
  await comment(page, 'Is thirty right?\nThe contract says fourteen.', 'Rita Gomez');
  const thread = panel(page).locator('.thread');
  await expect(thread).toHaveCount(1);
  await expect(thread).toContainText('Rita Gomez');
  await expect(thread).toContainText('Is thirty right?');
  await expect(thread.locator('.tag.new')).toHaveText('New');
  await expect(page.locator('.docx-page .comment-ref')).toHaveAttribute('data-n', '1');
  expect(await page.evaluate(() => [...(CSS.highlights.get('comment') || [])].map((r) => r.toString()))).toEqual(['thirty days']);

  const { zip, xml } = await download(page);
  const id = xml.match(/<w:commentRangeStart w:id="(\d+)"\/>/)[1];
  expect(xml).toContain(`<w:t xml:space="preserve">The fee is due within </w:t></w:r><w:commentRangeStart w:id="${id}"/><w:r><w:t xml:space="preserve">thirty days</w:t></w:r><w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r>`);
  const comments = await zip.file('word/comments.xml').async('string');
  // The last paragraph has an id, so the comment can be replied to and resolved (commentsExtended.xml).
  expect(comments).toMatch(new RegExp(`<w:comment w:id="${id}" w:author="Rita Gomez" w:date="\\d{4}-\\d\\d-\\d\\dT[\\d:]+Z" w:initials="RG"><w:p><w:r><w:annotationRef/></w:r><w:r><w:t xml:space="preserve">Is thirty right\\?</w:t></w:r></w:p><w:p w14:paraId="[0-9A-F]{8}"><w:r><w:t xml:space="preserve">The contract says fourteen.</w:t></w:r></w:p></w:comment>`));
  expect(await zip.file('word/_rels/document.xml.rels').async('string')).toContain('relationships/comments" Target="comments.xml"');
  expect(await zip.file('[Content_Types].xml').async('string')).toContain('PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"');
});

test('next to existing comments: they stay as they were, and ids don\'t clash', async ({ page }) => {
  await page.setViewportSize({ width: 1280, height: 800 });
  const existing = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments ${W_NS}><w:comment w:id="12" w:author="Bob" w:date="2026-01-01T00:00:00Z"><w:p><w:r><w:t>Old comment.</w:t></w:r></w:p></w:comment></w:comments>`;
  const docx = await buildDocx('<w:p><w:bookmarkStart w:id="40" w:name="b"/><w:commentRangeStart w:id="12"/><w:r><w:t xml:space="preserve">Commented </w:t></w:r><w:commentRangeEnd w:id="12"/><w:r><w:commentReference w:id="12"/></w:r><w:r><w:t>plain words.</w:t></w:r><w:bookmarkEnd w:id="40"/></w:p>',
    { parts: { 'word/comments.xml': existing }, rels: `<Relationship Id="rIdC" Type="${REL}/comments" Target="comments.xml"/>` });
  await openInEditor(page, docx);
  await select(page, 'plain', 2, 0); // a caret in "plain": the comment goes on the word
  await page.keyboard.press('ControlOrMeta+Alt+KeyM');
  await comment(page, 'Which words?', 'Rita');
  await expect(panel(page).locator('.thread')).toHaveCount(2);
  await expect(page.locator('.docx-page .comment-ref')).toHaveCount(2);
  const { zip, xml } = await download(page);
  expect(xml).toContain('<w:commentRangeStart w:id="41"/><w:r><w:t xml:space="preserve">plain</w:t></w:r><w:commentRangeEnd w:id="41"/><w:r><w:commentReference w:id="41"/></w:r>');
  const comments = await zip.file('word/comments.xml').async('string');
  expect(comments).toContain('<w:comment w:id="12" w:author="Bob" w:date="2026-01-01T00:00:00Z"><w:p><w:r><w:t>Old comment.</w:t></w:r></w:p></w:comment><w:comment w:id="41" w:author="Rita"');
});

test('Cancel adds nothing; a new comment can be deleted, and Undo brings it back', async ({ page }) => {
  await page.setViewportSize({ width: 1280, height: 800 });
  await openInEditor(page, await buildDocx(DOC));
  await select(page, 'Second');
  await page.getByTitle('New comment (Ctrl+Alt+M)').click();
  await composer(page).getByRole('textbox', { name: 'New comment' }).press('Escape');
  await expect(composer(page)).toHaveCount(0);
  await expect(page.locator('.docx-page .comment-ref')).toHaveCount(0);

  await select(page, 'Second');
  await page.getByTitle('New comment (Ctrl+Alt+M)').click();
  await composer(page).getByRole('textbox', { name: 'Your name' }).fill('Rita');
  await composer(page).getByRole('textbox', { name: 'New comment' }).fill('Remove?');
  await composer(page).getByRole('textbox', { name: 'New comment' }).press('Control+Enter');
  await expect(panel(page).locator('.thread')).toHaveCount(1);
  await panel(page).getByRole('button', { name: 'Delete' }).click();
  await expect(page.locator('.statusbar [role=status]')).toHaveText('Comment deleted – Undo (Ctrl+Z) brings it back');
  await expect(page.locator('.docx-page .comment-ref')).toHaveCount(0);
  await page.locator('.docx-page p').first().click();
  await page.keyboard.press('ControlOrMeta+z');
  await expect(page.locator('.docx-page .comment-ref')).toHaveCount(1);
  await expect(panel(page).locator('.thread')).toHaveCount(1);
});

test('a new comment whose text is deleted isn\'t saved; one kept is in the review', async ({ page }) => {
  await page.setViewportSize({ width: 1280, height: 800 });
  await openInEditor(page, await buildDocx(DOC));
  await select(page, 'Second');
  await page.getByTitle('New comment (Ctrl+Alt+M)').click();
  await comment(page, 'Gone soon', 'Rita');
  await select(page, 'fee');
  await page.getByTitle('New comment (Ctrl+Alt+M)').click();
  await comment(page, 'Which fee?'); // the name is known now
  // Delete the second paragraph, with its comment.
  await page.evaluate(() => {
    const p = document.querySelectorAll('.docx-page p')[1];
    const range = document.createRange();
    range.selectNodeContents(p);
    document.querySelector('.docx-page').focus();
    getSelection().removeAllRanges();
    getSelection().addRange(range);
  });
  await page.keyboard.press('Delete');
  await expect(panel(page).locator('.thread')).toHaveCount(1);

  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  const review = page.getByRole('dialog', { name: 'Review your changes' });
  await expect(review.locator('.notes')).toContainText('Comment added by Rita: “Which fee?”');
  await expect(review.locator('.notes')).not.toContainText('Gone soon');
  await expect(review.locator('.change', { hasText: 'The fee is due' })).toContainText('Comment added');
  await review.getByRole('button', { name: 'Keep editing' }).click();
  const { zip } = await download(page);
  const comments = await zip.file('word/comments.xml').async('string');
  expect(comments).toContain('Which fee?');
  expect(comments).not.toContain('Gone soon');
});
