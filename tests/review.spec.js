// "Review my changes": before Download, a list of what changed (paragraph by paragraph, with the
// words), and of anything else that changes. An untouched document must say "No changes".
import { test, expect } from '@playwright/test';
import fs from 'node:fs';
import { buildDocx, openInEditor, download, KITCHEN_SINK, W_NS } from './helpers.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const DOC = '<w:p><w:r><w:t xml:space="preserve">The fee is due on the first working day of each month.</w:t></w:r></w:p>'
  + '<w:p><w:r><w:t>Second paragraph.</w:t></w:r></w:p><w:p><w:r><w:t>Third paragraph.</w:t></w:r></w:p>';

const dialog = (page) => page.getByRole('dialog', { name: 'Review your changes' });
const changes = (page) => dialog(page).locator('.change');

async function openReview(page) {
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  await expect(dialog(page)).toBeVisible();
}

async function select(page, [a, ao], [b, bo] = [a, ao]) {
  await page.evaluate(([a, ao, b, bo]) => {
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
  }, [a, ao, b, bo]);
}

for (const [name, load] of [
  ['a real Word letter', () => fs.readFileSync('tests/fixtures/word-letter-id.docx')],
  ['the kitchen-sink document', () => buildDocx(KITCHEN_SINK)],
]) {
  test(`opening ${name} and downloading it says there are no changes`, async ({ page }) => {
    await openInEditor(page, await load());
    await openReview(page);
    await expect(dialog(page).locator('.summary')).toHaveText('No changes: the file is the same as the one you opened.');
    await expect(changes(page)).toHaveCount(0);
    await expect(dialog(page).locator('.notes li')).toHaveCount(0);
  });
}

test('an edited word shows in its paragraph, with a warning that it isn\'t tracked', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await select(page, ['due on', 0], ['due on', 3]);
  await page.keyboard.type('payable');
  await openReview(page);
  await expect(dialog(page).locator('.summary')).toHaveText('1 paragraph changed.');
  await expect(changes(page)).toHaveCount(1);
  await expect(changes(page).locator('.where')).toHaveText(/Paragraph 1\s*Changed/);
  await expect(changes(page).locator('del')).toHaveText('due');
  await expect(changes(page).locator('ins')).toHaveText('payable');
  await expect(changes(page).locator('.words')).toHaveText('The fee is duepayable on the first working day of each month.');
  await expect(dialog(page).locator('.notes')).toContainText('Some text changes aren’t tracked');

  // Keep editing closes it without downloading; Download downloads.
  await dialog(page).getByRole('button', { name: 'Keep editing' }).click();
  await expect(dialog(page)).toHaveCount(0);
  const { xml } = await download(page);
  expect(xml).toContain('The fee is payable on the first');
});

test('tracked edits are listed as tracked, and turning Track Changes on is mentioned', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await page.getByRole('button', { name: 'Track changes' }).click();
  await page.getByRole('dialog', { name: 'Your name on tracked changes' }).getByRole('textbox').fill('Rita');
  await page.getByRole('button', { name: 'Start tracking' }).click();
  await select(page, ['Second paragraph.', 0], ['Second paragraph.', 6]);
  await page.keyboard.type('Next');
  await page.keyboard.press('ControlOrMeta+s'); // Ctrl+S reviews too
  await expect(dialog(page)).toBeVisible();
  await expect(changes(page).locator('.tag')).toHaveText('Tracked');
  const notes = dialog(page).locator('.notes');
  await expect(notes).toContainText('Recorded as tracked changes (1 insertion and 1 deletion)');
  await expect(notes).toContainText('Track Changes will be on when the file is opened in Word.');
  await expect(notes).not.toContainText('aren’t tracked');
});

test('formatting, added and removed paragraphs and removed comments are all listed', async ({ page }) => {
  const comments = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments ${W_NS}><w:comment w:id="0" w:author="Ann Lee"><w:p><w:r><w:t>Check.</w:t></w:r></w:p></w:comment></w:comments>`;
  const docx = await buildDocx(`${DOC}<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Commented paragraph.</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p><w:p><w:r><w:t>Last.</w:t></w:r></w:p>`,
    { parts: { 'word/comments.xml': comments }, rels: `<Relationship Id="rIdC" Type="${REL}/comments" Target="comments.xml"/>` });
  await openInEditor(page, docx);
  // Heading 2 on the first paragraph (a style the file doesn't have yet).
  await select(page, ['The fee', 0]);
  await page.getByLabel('Paragraph style').selectOption('h2');
  // Bold on "Third".
  await select(page, ['Third paragraph.', 0], ['Third paragraph.', 5]);
  await page.keyboard.press('ControlOrMeta+b');
  // A new paragraph after "Second paragraph.".
  await select(page, ['Second paragraph.', 17]);
  await page.keyboard.press('Enter');
  await page.keyboard.type('Brand new.');
  // The commented paragraph, with its break, deleted.
  await select(page, ['Commented paragraph.', 0], ['Last.', 0]);
  await page.keyboard.press('Delete');
  // The paragraphs join (no empty paragraph is left holding the comment's range start).
  expect(await page.locator('.docx-page > *:not(.raw-hidden)').allTextContents()).toEqual([
    'The fee is due on the first working day of each month.', 'Second paragraph.', 'Brand new.', 'Third paragraph.', 'Last.',
  ]);

  await openReview(page);
  // The whole list first, so a failure shows everything that is listed.
  expect((await changes(page).allInnerTexts()).map((t) => t.replace(/\s+/g, ' ').trim()).sort()).toEqual([
    'Paragraph 1 Changed The fee is due on the first working day of each month. Style: Normal → Heading 2',
    'Paragraph 3 Added Brand new.',
    'Paragraph 4 Changed Third paragraph. Text formatting changed (bold, italic, underline or strikethrough)',
    'Paragraph 4 Removed Commented paragraph.',
  ]);
  await expect(dialog(page).locator('.summary')).toHaveText('2 paragraphs changed, 1 added, 1 removed.');
  const entry = (text) => changes(page).filter({ hasText: text });
  await expect(entry('The fee is due')).toContainText('Style: Normal → Heading 2');
  await expect(entry('Brand new.').locator('.kind')).toHaveText('Added');
  await expect(entry('Third paragraph.')).toContainText('Text formatting changed');
  await expect(entry('Commented paragraph.').locator('.kind')).toHaveText('Removed');
  await expect(entry('Commented paragraph.').locator('del')).toHaveText('Commented paragraph.');
  const notes = dialog(page).locator('.notes');
  await expect(notes).toContainText('Comment by Ann Lee removed, with the text it was attached to.');
  await expect(notes).toContainText('Formatting changes are applied directly, never tracked.');
  await expect(notes).toContainText('Paragraph styles you used are added to the file’s style list.');
});

test('the review can be skipped, and still opened from the toolbar', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await openReview(page);
  await dialog(page).getByRole('checkbox', { name: 'Download without this review next time' }).check();
  await dialog(page).getByRole('button', { name: 'Keep editing' }).click();
  const [dl] = await Promise.all([page.waitForEvent('download'), page.locator('.app-header').getByRole('button', { name: 'Download' }).click()]);
  expect(dl.suggestedFilename()).toBe('fixture.docx');
  await expect(dialog(page)).toHaveCount(0);
  await page.getByRole('button', { name: 'Review' }).click();
  await expect(dialog(page)).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(dialog(page)).toHaveCount(0);
});
