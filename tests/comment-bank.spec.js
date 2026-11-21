// The comment bank, for marking: a comment given often is saved for reuse, then put into a new
// comment with a click or Alt+1 to Alt+9, at the caret. Kept in this browser (never in the file),
// and exported or imported as a CSV file.
import fs from 'node:fs';
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const DOC = '<w:p><w:r><w:t xml:space="preserve">The results clearly show that encapsulation is very important.</w:t></w:r></w:p><w:p><w:r><w:t>Second paragraph says more.</w:t></w:r></w:p>';
const panel = (page) => page.getByRole('complementary', { name: 'Comments' });
const composer = (page) => panel(page).getByRole('form', { name: 'New comment' });
const box = (page) => composer(page).getByRole('textbox', { name: 'New comment' });

async function select(page, text) {
  await page.evaluate((t) => {
    const root = document.querySelector('.docx-page');
    const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    for (let n = w.nextNode(); n; n = w.nextNode()) {
      const i = n.data.indexOf(t);
      if (i < 0) continue;
      const range = document.createRange();
      range.setStart(n, i);
      range.setEnd(n, i + t.length);
      root.focus();
      getSelection().removeAllRanges();
      getSelection().addRange(range);
      return;
    }
    throw new Error(`no text "${t}"`);
  }, text);
}
async function newComment(page, on) {
  await select(page, on);
  await page.getByTitle('New comment (Ctrl+Alt+M)').click();
  const name = composer(page).getByRole('textbox', { name: 'Your name' });
  if (await name.count()) await name.fill('Prof. Reviewer');
}

test('a comment saved for reuse is put into the next with a click, or Alt+1, at the caret; it stays in this browser', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await newComment(page, 'clearly');
  await box(page).fill('Avoid intensifiers: quantify the effect instead.');
  await composer(page).getByRole('button', { name: 'Save for reuse' }).click();
  await expect(composer(page).getByRole('status')).toHaveText('Saved: Alt+1 puts it in a comment');
  await composer(page).getByRole('button', { name: 'Comment', exact: true }).click();

  // The next comment: the saved one with a click, after what was typed.
  await newComment(page, 'very important');
  await composer(page).getByRole('button', { name: /Saved comments/ }).click();
  const saved = composer(page).getByRole('list', { name: 'Saved comments' }).getByRole('listitem');
  await expect(saved).toHaveText([/1Avoid intensifiers/]);
  await box(page).fill('Here too.');
  await saved.first().getByRole('button', { name: /^Avoid intensifiers/ }).click();
  await expect(box(page)).toHaveValue('Here too. Avoid intensifiers: quantify the effect instead.');
  await composer(page).getByRole('button', { name: 'Comment', exact: true }).click();

  // Alt+1 in the box, then the comment as posted.
  await newComment(page, 'Second paragraph');
  await box(page).focus();
  await page.keyboard.press('Alt+Digit1');
  await expect(box(page)).toHaveValue('Avoid intensifiers: quantify the effect instead.');
  await composer(page).getByRole('button', { name: 'Comment', exact: true }).click();
  await expect(panel(page).locator('.thread')).toHaveCount(3);

  // Saved in this browser, not in the document.
  const { zip } = await download(page);
  const comments = await zip.file('word/comments.xml').async('string');
  expect(comments.match(/Avoid intensifiers/g)).toHaveLength(3);
  expect(JSON.parse(await page.evaluate(() => localStorage.getItem('intact.commentBank')))).toEqual(['Avoid intensifiers: quantify the effect instead.']);
  await page.reload();
  await page.goto('./');
  expect(JSON.parse(await page.evaluate(() => localStorage.getItem('intact.commentBank')))).toHaveLength(1);
});

test('saved comments are removed, found, exported and imported as a CSV file', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await page.evaluate(() => localStorage.setItem('intact.commentBank', JSON.stringify(['Cite a primary source.', 'Define this term, on first use.', 'Good point, well argued.'])));
  await newComment(page, 'clearly');
  await composer(page).getByRole('button', { name: /Saved comments/ }).click();
  const list = composer(page).getByRole('list', { name: 'Saved comments' });
  await expect(list.getByRole('listitem')).toHaveCount(3);
  // Saving one already there says so.
  await box(page).fill('cite a primary source.');
  await composer(page).getByRole('button', { name: 'Save for reuse' }).click();
  await expect(composer(page).getByRole('status')).toHaveText('Already saved');
  // Removed.
  await list.getByRole('button', { name: 'Remove saved comment: Good point, well argued.' }).click();
  await expect(list.getByRole('listitem')).toHaveCount(2);

  // Exported: a CSV with a header, the comma quoted.
  const [dl] = await Promise.all([page.waitForEvent('download'), composer(page).getByRole('button', { name: 'Export…' }).click()]);
  expect(dl.suggestedFilename()).toBe('saved-comments.csv');
  expect(fs.readFileSync(await dl.path(), 'utf8')).toBe('﻿Comment\r\nCite a primary source.\r\n"Define this term, on first use."\r\n');

  // Imported: a column from a spreadsheet, without a header; one already saved left out.
  const csv = 'Cite a primary source.\n"Explain the method in more detail,\nwith an example."\nCheck the units.\n';
  await composer(page).locator('input[type=file]').setInputFiles({ name: 'feedback.csv', mimeType: 'text/csv', buffer: Buffer.from(csv) });
  await expect(composer(page).getByRole('status')).toHaveText('2 comments imported, 1 already saved');
  await expect(list.getByRole('listitem')).toHaveCount(4);
  await expect(list.getByRole('listitem').nth(2)).toContainText('Explain the method in more detail,\nwith an example.');
  // More than six: a box to find one (the list stays open, as last time).
  await page.evaluate(() => localStorage.setItem('intact.commentBank', JSON.stringify(Array.from({ length: 8 }, (_, i) => `Feedback number ${i + 1}`))));
  await composer(page).getByRole('button', { name: 'Cancel' }).click();
  await newComment(page, 'Second paragraph');
  await composer(page).getByRole('searchbox', { name: 'Find a saved comment' }).fill('number 7');
  await expect(list.getByRole('listitem')).toHaveText([/7Feedback number 7/]);
});
