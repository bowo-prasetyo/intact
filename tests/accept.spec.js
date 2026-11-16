// Accepting and rejecting tracked changes, as in Word: anyone's insertions, deletions, paragraph
// breaks, formatting changes and table rows, and this session's own, one at a time (Previous, Next)
// or all at once. Changes inside locked content are resolved together, as one item.
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
  // A tracked row (and, only for Word, an insertion in locked content), then a tracked formatting change.
  + '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>'
  + `<w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:sdt><w:sdtPr><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent><w:p>${ins('Fay', 'Header', 6)}</w:p></w:sdtContent></w:sdt><w:p/></w:tc></w:tr>`
  + '<w:tr><w:trPr><w:ins w:id="4" w:author="Dan" w:date="2026-01-04T00:00:00Z"/></w:trPr><w:tc><w:p><w:r><w:t>Row</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
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

test('the toolbar counts the changes, those in locked content too', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await expect(group(page)).toBeVisible();
  await expect(group(page).locator('.badge')).toHaveText('6');
  await expect(group(page).locator('.badge')).toHaveAttribute('title', '6 tracked changes to accept or reject here');
  // How they show: Dan's row tinted, Eve's formatting change underlined.
  await expect(page.locator('.docx-page tr[data-rev-row="ins"]')).toHaveText('Row');
  await expect(page.locator('.docx-page [data-fmt="run"]')).toHaveAttribute('title', 'Formatting changed by Eve on 2026-01-05 – tracked change: Accept or Reject on the toolbar');
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
  await expect(status(page)).toHaveText('The tracked change in the locked field by Fay');
  await next.click();
  await expect(status(page)).toHaveText('The inserted row by Dan, 2026-01-04');
  await next.click();
  await expect(status(page)).toHaveText('The formatting change by Eve, 2026-01-05');
  await next.click();
  await expect(status(page)).toHaveText('The insertion by Alice, 2026-01-01');
  await group(page).getByTitle('Previous change').click();
  await expect(status(page)).toHaveText('The formatting change by Eve, 2026-01-05');
});

test('accepting an insertion and rejecting a deletion leave ordinary text, then go to the next', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await page.locator('.docx-page .inserted').first().click(); // the caret goes beside it
  await group(page).getByRole('button', { name: 'Accept' }).first().click();
  await expect(status(page)).toHaveText('Accepted the insertion by Alice. Next: the deletion by Bob');
  await expect(page.locator('.docx-page p').first().locator('.inserted')).toHaveCount(0);
  await group(page).getByRole('button', { name: 'Reject' }).first().click(); // Bob's deletion, now selected
  await expect(status(page)).toHaveText('Rejected the deletion by Bob. Next: the paragraph break insertion by Carol');
  await expect(page.locator('.docx-page p').first()).toHaveText('Pay promptly the full fee.');
  await expect(group(page).locator('.badge')).toHaveText('4');
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

test('Accept all accepts every change, those in locked content too, and the review agrees', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await page.getByTitle('More ways to accept').click();
  await page.getByRole('menuitem', { name: 'Accept all changes' }).click();
  await expect(status(page)).toHaveText('Accepted all 6 changes');
  await expect(group(page)).toHaveCount(0); // nothing left to accept or reject
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  const review = page.getByRole('dialog', { name: 'Review your changes' });
  await expect(review.locator('.notes')).toContainText('You accepted 6 tracked changes. As in Word, accepting and rejecting aren’t themselves tracked.');
  await expect(review.locator('.notes')).not.toContainText('aren’t tracked: the recipient');
  await expect(review.locator('.tag.resolved').first()).toHaveText('Change accepted or rejected');
  const [dl] = await Promise.all([page.waitForEvent('download'), review.getByRole('button', { name: 'Download' }).click()]);
  const xml = await (await JSZip.loadAsync(await (await import('node:fs')).promises.readFile(await dl.path()))).file('word/document.xml').async('string');
  expect(xml).not.toMatch(/w:author="(Alice|Bob|Carol|Dan|Eve|Fay)"/);
  expect(xml).toContain('<w:tr><w:tc><w:p><w:r><w:t xml:space="preserve">Row</w:t></w:r></w:p></w:tc></w:tr>'); // Dan's row, kept
  expect(xml).toContain('<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Bold by Eve.</w:t></w:r>'); // Eve's bold, kept
  expect(xml).toContain('<w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t xml:space="preserve">Header</w:t></w:r></w:p></w:sdtContent>'); // Fay's insertion accepted, the field still locked
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

test('formatting changes: rejected, the old formatting comes back; accepted, the record goes', async ({ page }) => {
  const docx = await buildDocx('<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="center"/><w:pPrChange w:id="8" w:author="Gus" w:date="2026-02-01T00:00:00Z"><w:pPr/></w:pPrChange></w:pPr><w:r><w:t>Was a plain paragraph</w:t></w:r></w:p>'
    + '<w:p><w:r><w:t xml:space="preserve">Some </w:t></w:r><w:r><w:rPr><w:b/><w:i/><w:rPrChange w:id="9" w:author="Eve" w:date="2026-01-05T00:00:00Z"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>words</w:t></w:r><w:r><w:t>.</w:t></w:r></w:p>');
  await openInEditor(page, docx);
  await expect(page.locator('.docx-page h1[data-fmt-para]')).toHaveText('Was a plain paragraph');
  // Reject Eve's change (bold and italic, was italic only): back to italic.
  await page.locator('.docx-page [data-fmt="run"]').click();
  await group(page).getByRole('button', { name: 'Reject' }).first().click();
  await expect(status(page)).toHaveText(/^Rejected the formatting change by Eve\./);
  await expect(page.locator('.docx-page p em')).toHaveText('words');
  await expect(page.locator('.docx-page p strong')).toHaveCount(0);
  // Reject Gus's paragraph change: back to a plain, left-aligned paragraph.
  await caretAt(page, 'Was a plain', 3);
  await page.getByTitle('More ways to reject').click();
  await page.getByRole('menuitem', { name: 'Reject, and stay here' }).click();
  await expect(page.locator('.docx-page p').first()).toHaveText('Was a plain paragraph');
  await expect(page.locator('.docx-page h1')).toHaveCount(0);
  let { xml } = await download(page);
  expect(xml).toContain('<w:body><w:p><w:r><w:t xml:space="preserve">Was a plain paragraph</w:t></w:r></w:p>');
  expect(xml).toContain('<w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">words</w:t></w:r>');

  // Accepting instead keeps the new formatting, without the record.
  await page.keyboard.press('ControlOrMeta+z');
  await page.keyboard.press('ControlOrMeta+z');
  await page.getByTitle('More ways to accept').click();
  await page.getByRole('menuitem', { name: 'Accept all changes' }).click();
  ({ xml } = await download(page));
  expect(xml).toContain('<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">Was a plain paragraph</w:t>');
  expect(xml).toContain('<w:r><w:rPr><w:b/><w:i/></w:rPr><w:t xml:space="preserve">words</w:t></w:r>');
});

test('tracked rows: an inserted row rejected goes (and an empty table with it); a deleted row accepted goes', async ({ page }) => {
  const row = (text, rev = '') => `<w:tr>${rev ? `<w:trPr>${rev}</w:trPr>` : ''}<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc></w:tr>`;
  const table = (rows) => `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>${rows}</w:tbl>`;
  const docx = await buildDocx(table(row('Kept') + row('Added by Dan', '<w:ins w:id="1" w:author="Dan" w:date="2026-01-04T00:00:00Z"/>') + row('Removed by Dan', '<w:del w:id="2" w:author="Dan" w:date="2026-01-04T00:00:00Z"/>')) + '<w:p/>'
    + table(row('Only row', '<w:ins w:id="3" w:author="Dan" w:date="2026-01-04T00:00:00Z"/>')) + '<w:p><w:r><w:t>End.</w:t></w:r></w:p>');
  await openInEditor(page, docx);
  await expect(page.locator('.docx-page tr[data-rev-row="del"]')).toHaveText('Removed by Dan');
  await page.locator('td', { hasText: 'Added by Dan' }).click();
  await group(page).getByRole('button', { name: 'Reject' }).first().click();
  await expect(status(page)).toHaveText('Rejected the inserted row by Dan. Next: the deleted row by Dan');
  await group(page).getByRole('button', { name: 'Accept' }).first().click(); // the deleted row, now selected
  await page.locator('td', { hasText: 'Only row' }).click();
  await group(page).getByRole('button', { name: 'Reject' }).first().click();
  await expect(page.locator('.docx-page table')).toHaveCount(1);
  await expect(page.locator('.docx-page td')).toHaveText(['Kept']);
  const { xml } = await download(page);
  expect(xml).toContain('<w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t xml:space="preserve">Kept</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/><w:p><w:r><w:t xml:space="preserve">End.</w:t>');
  expect(xml).not.toContain('Dan');

  await page.keyboard.press('ControlOrMeta+z');
  await expect(page.locator('.docx-page table')).toHaveCount(2);
});
