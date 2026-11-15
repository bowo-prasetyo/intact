// Accepting and rejecting tracked changes, as in Word: anyone's insertions, deletions and
// paragraph breaks, and this session's own, one at a time (Previous, Next) or all at once. Changes
// only Word can resolve (formatting, table rows) are counted and named, never dropped.
import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import { buildDocx, openInEditor, download, canonicalize, documentXmlOf, W_NS } from './helpers.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const status = (page) => page.locator('.statusbar [role=status]');
const group = (page) => page.getByRole('group', { name: 'Tracked changes' });
const ins = (author, text, id) => `<w:ins w:id="${id}" w:author="${author}" w:date="2026-01-0${id}T00:00:00Z"><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:ins>`;
const del = (author, text, id) => `<w:del w:id="${id}" w:author="${author}" w:date="2026-01-0${id}T00:00:00Z"><w:r><w:delText xml:space="preserve">${text}</w:delText></w:r></w:del>`;

const DOC = `<w:p><w:r><w:t xml:space="preserve">Pay </w:t></w:r>${ins('Alice', 'promptly ', 1)}<w:r><w:t xml:space="preserve">the </w:t></w:r>${del('Bob', 'full ', 2)}<w:r><w:t>fee.</w:t></w:r></w:p>`
  + '<w:p><w:pPr><w:rPr><w:ins w:id="3" w:author="Carol" w:date="2026-01-03T00:00:00Z"/></w:rPr></w:pPr><w:r><w:t>Split part one</w:t></w:r></w:p>'
  + '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>part two.</w:t></w:r></w:p>'
  // Only Word: a tracked table row and a tracked formatting change.
  + '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:trPr><w:ins w:id="4" w:author="Dan" w:date="2026-01-04T00:00:00Z"/></w:trPr><w:tc><w:p><w:r><w:t>Row</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
  + '<w:p><w:r><w:rPr><w:b/><w:rPrChange w:id="5" w:author="Eve" w:date="2026-01-05T00:00:00Z"><w:rPr/></w:rPrChange></w:rPr><w:t>Bold by Eve.</w:t></w:r></w:p>';

async function caretAt(page, text, offset = 0) {
  await page.evaluate(([t, o]) => {
    const root = document.querySelector('.docx-page');
    const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    for (let n = w.nextNode(); n; n = w.nextNode()) {
      if (!n.data.includes(t)) continue;
      const range = document.createRange();
      range.setStart(n, n.data.indexOf(t) + o);
      root.focus();
      getSelection().removeAllRanges();
      getSelection().addRange(range);
      return;
    }
    throw new Error(`no text "${t}"`);
  }, [text, offset]);
}

test('the toolbar counts the changes, and names those only Word can resolve', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await expect(group(page)).toBeVisible();
  await expect(group(page).locator('.badge')).toHaveText('3');
  await expect(group(page).locator('.badge')).toHaveAttribute('title', '3 tracked changes to accept or reject here; 2 more (formatting, table rows or locked content) only in Word');
  // Carol's tracked paragraph break shows as a ¶.
  await expect(page.locator('.docx-page p[data-rev-mark="ins"]')).toHaveText('Split part one');
});

test('Next and Previous go from change to change, wrapping round', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await caretAt(page, 'Pay ');
  const next = group(page).getByTitle('Next change');
  await next.click();
  await expect(status(page)).toHaveText('The insertion by Alice, 2026-01-01');
  await next.click();
  await expect(status(page)).toHaveText('The deletion by Bob, 2026-01-02');
  await next.click();
  await expect(status(page)).toHaveText('The paragraph break insertion by Carol, 2026-01-03');
  await next.click();
  await expect(status(page)).toHaveText('The insertion by Alice, 2026-01-01');
  await group(page).getByTitle('Previous change').click();
  await expect(status(page)).toHaveText('The paragraph break insertion by Carol, 2026-01-03');
});

test('accepting an insertion and rejecting a deletion leave ordinary text, then go to the next', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await page.locator('.docx-page .inserted').click(); // the caret goes beside it
  await group(page).getByRole('button', { name: 'Accept' }).first().click();
  await expect(status(page)).toHaveText('Accepted the insertion by Alice. Next: the deletion by Bob');
  await expect(page.locator('.docx-page .inserted')).toHaveCount(0);
  await group(page).getByRole('button', { name: 'Reject' }).first().click(); // Bob's deletion, now selected
  await expect(status(page)).toHaveText('Rejected the deletion by Bob. Next: the paragraph break insertion by Carol');
  await expect(page.locator('.docx-page p').first()).toHaveText('Pay promptly the full fee.');
  await expect(group(page).locator('.badge')).toHaveText('1');
  const { xml } = await download(page);
  expect(xml).toMatch(/<w:p><w:r><w:t xml:space="preserve">Pay (<\/w:t><\/w:r><w:r><w:t xml:space="preserve">)?promptly (<\/w:t><\/w:r><w:r><w:t xml:space="preserve">)?the (<\/w:t><\/w:r><w:r><w:t xml:space="preserve">)?full (<\/w:t><\/w:r><w:r><w:t xml:space="preserve">)?fee\.<\/w:t><\/w:r><\/w:p>/);
  expect(xml).not.toMatch(/w:author="(Alice|Bob)"/);
});

test('a tracked paragraph break: rejected, its paragraph joins the next; accepted, it stays', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await caretAt(page, 'Split part one', 14); // at the end of the paragraph: on its break
  await group(page).getByRole('button', { name: 'Reject' }).first().click();
  await expect(status(page)).toHaveText(/^Rejected the paragraph break insertion by Carol\./);
  let { xml } = await download(page);
  // Joined, with the next paragraph's properties (its break is the one that stays).
  expect(xml).toMatch(/<w:p><w:pPr><w:jc w:val="center"\/><\/w:pPr><w:r><w:t xml:space="preserve">Split part one(<\/w:t><\/w:r><w:r><w:t xml:space="preserve">)?part two\.<\/w:t><\/w:r><\/w:p>/);

  await page.keyboard.press('ControlOrMeta+z');
  await caretAt(page, 'Split part one', 14);
  await page.getByTitle('More ways to accept').click();
  await page.getByRole('menuitem', { name: 'Accept, and stay here' }).click();
  ({ xml } = await download(page));
  expect(xml).toContain('<w:p><w:r><w:t xml:space="preserve">Split part one</w:t></w:r></w:p>');
  await expect(page.locator('.docx-page p[data-rev-mark]')).toHaveCount(0);
});

test('Accept all accepts every change it can, says which are left for Word, and the review agrees', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await page.getByTitle('More ways to accept').click();
  await page.getByRole('menuitem', { name: 'Accept all changes' }).click();
  await expect(status(page)).toHaveText('Accepted all 3 changes. 2 more (formatting, table rows or locked content) can only be accepted in Word');
  await expect(group(page).locator('.badge')).toHaveText('0');
  await expect(group(page).getByRole('button', { name: 'Accept' }).first()).toBeDisabled();
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  const review = page.getByRole('dialog', { name: 'Review your changes' });
  await expect(review.locator('.notes')).toContainText('You accepted 3 tracked changes. As in Word, accepting and rejecting aren’t themselves tracked.');
  await expect(review.locator('.notes')).not.toContainText('aren’t tracked: the recipient');
  await expect(review.locator('.tag.resolved').first()).toHaveText('Change accepted or rejected');
  const [dl] = await Promise.all([page.waitForEvent('download'), review.getByRole('button', { name: 'Download' }).click()]);
  const xml = await (await JSZip.loadAsync(await (await import('node:fs')).promises.readFile(await dl.path()))).file('word/document.xml').async('string');
  expect(xml).not.toMatch(/w:author="(Alice|Bob|Carol)"/);
  expect(xml).toContain('w:author="Dan"'); // the table row: only in Word
  expect(xml).toContain('<w:rPrChange w:id="5" w:author="Eve"');
});

test('a selection accepts the changes in it; Undo takes it all back', async ({ page }) => {
  const docx = await buildDocx(DOC);
  await openInEditor(page, docx);
  await page.evaluate(() => {
    const p = document.querySelector('.docx-page p');
    const range = document.createRange();
    range.selectNodeContents(p);
    document.querySelector('.docx-page').focus();
    getSelection().removeAllRanges();
    getSelection().addRange(range);
  });
  await group(page).getByRole('button', { name: 'Reject' }).first().click();
  await expect(status(page)).toHaveText('Rejected 2 changes');
  await expect(page.locator('.docx-page p').first()).toHaveText('Pay the full fee.'); // Alice's insertion gone, Bob's deletion undone
  await page.keyboard.press('ControlOrMeta+z');
  await expect(page.locator('.docx-page p').first()).toHaveText('Pay promptly the full fee.');
  const { xml } = await download(page);
  expect(await canonicalize(page, xml)).toEqual(await canonicalize(page, await documentXmlOf(docx)));
});

test('your own tracked changes: Reject all puts the document back as it was', async ({ page }) => {
  const docx = await buildDocx('<w:p><w:r><w:t>Original text here.</w:t></w:r></w:p><w:p><w:r><w:t>Second.</w:t></w:r></w:p>');
  await openInEditor(page, docx);
  await page.getByRole('button', { name: 'Track changes' }).click();
  await page.getByRole('dialog', { name: 'Your name on tracked changes' }).getByRole('textbox').fill('Rita');
  await page.getByRole('button', { name: 'Start tracking' }).click();
  await caretAt(page, 'here.', 0);
  await page.keyboard.type('right ');
  await caretAt(page, 'Original', 8);
  for (let i = 0; i < 8; i++) await page.keyboard.press('Backspace');
  await caretAt(page, 'Second.', 7);
  await page.keyboard.press('Enter');
  await page.keyboard.type('Third.');
  await expect(group(page).locator('.badge')).toHaveText('4');
  await page.getByTitle('More ways to reject').click();
  await page.getByRole('menuitem', { name: 'Reject all changes' }).click();
  await expect(page.locator('.docx-page').locator('.tc-ins, .tc-del, [data-tc-mark]')).toHaveCount(0);
  const { xml } = await download(page);
  expect(await canonicalize(page, xml)).toEqual(await canonicalize(page, await documentXmlOf(docx)));
});

test('a comment inside an accepted insertion stays', async ({ page }) => {
  const comments = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments ${W_NS}><w:comment w:id="0" w:author="Bob"><w:p><w:r><w:t>Why?</w:t></w:r></w:p></w:comment></w:comments>`;
  const docx = await buildDocx('<w:p><w:r><w:t xml:space="preserve">Pay </w:t></w:r><w:commentRangeStart w:id="0"/><w:ins w:id="7" w:author="Alice" w:date="2026-01-01T00:00:00Z"><w:r><w:t>now</w:t></w:r><w:r><w:commentReference w:id="0"/></w:r></w:ins><w:commentRangeEnd w:id="0"/><w:r><w:t>.</w:t></w:r></w:p>',
    { parts: { 'word/comments.xml': comments }, rels: `<Relationship Id="rIdC" Type="${REL}/comments" Target="comments.xml"/>` });
  await openInEditor(page, docx);
  await page.getByTitle('More ways to accept').click();
  await page.getByRole('menuitem', { name: 'Accept all changes' }).click();
  await expect(page.locator('.docx-page .comment-ref')).toHaveCount(1);
  const { zip, xml } = await download(page);
  expect(xml).toContain('<w:commentReference w:id="0"/>');
  expect(await zip.file('word/comments.xml').async('string')).toBe(comments);
});
