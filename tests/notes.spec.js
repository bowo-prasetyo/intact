// Footnotes and endnotes are edited in place, below the text, as in Word's draft view: typing,
// formatting, Find and Replace, Undo and Track Changes work in them as in the text. Only notes that
// changed are rewritten in their part (footnotes.xml / endnotes.xml); untouched, the parts come back
// byte for byte. A note can't be deleted with the text around it; it goes when its reference does.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, W_NS } from './helpers.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const SEPARATORS = '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>';
const mark = (local) => `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:${local}/></w:r>`;
const note = (local, id, text) => `<w:${local} w:id="${id}"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>${mark(`${local}Ref`)}<w:r><w:t xml:space="preserve"> ${text}</w:t></w:r></w:p></w:${local}>`;
const FOOTNOTES = `${XML}<w:footnotes ${W_NS}>${SEPARATORS}${note('footnote', 1, 'Smith (2020), p. 4.')}${note('footnote', 2, 'The data were recieved in 2021.')}</w:footnotes>`;
const ENDNOTES = `${XML}<w:endnotes ${W_NS}><w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>${note('endnote', 1, 'First endnote.')}${note('endnote', 2, 'Second endnote.')}</w:endnotes>`;
const OPTS = {
  parts: { 'word/footnotes.xml': FOOTNOTES, 'word/endnotes.xml': ENDNOTES },
  rels: `<Relationship Id="rIdF" Type="${REL}/footnotes" Target="footnotes.xml"/><Relationship Id="rIdE" Type="${REL}/endnotes" Target="endnotes.xml"/>`,
};
const ref = (local, id) => `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:${local}Reference w:id="${id}"/></w:r>`;
// The text refers to footnote 2 before footnote 1: the notes are numbered (and listed) in that order.
const BODY = `<w:p><w:r><w:t xml:space="preserve">Prior work agrees.</w:t></w:r>${ref('footnote', 2)}<w:r><w:t xml:space="preserve"> So do others.</w:t></w:r>${ref('footnote', 1)}</w:p>`
  + `<w:p><w:r><w:t xml:space="preserve">See the appendix.</w:t></w:r>${ref('endnote', 1)}<w:r><w:t xml:space="preserve"> And the method.</w:t></w:r>${ref('endnote', 2)}</w:p>`;

const status = (page) => page.locator('.statusbar [role=status]');
const footnote = (page, n) => page.locator(`.docx-notes[data-notes="footnote"] .docx-note[data-n="${n}"]`);
const endnote = (page, n) => page.locator(`.docx-notes[data-notes="endnote"] .docx-note[data-n="${n}"]`);
const reference = (page, kind, n) => page.locator(`.docx-page .note-ref[data-note-ref="${kind}"][data-n="${n}"]`);
const part = async (zip, name) => zip.file(`word/${name}.xml`).async('string');

async function open(page, body = BODY) {
  await openInEditor(page, await buildDocx(body, OPTS));
}

/** Puts the caret at the end of note `el`'s text. */
async function endOf(page, el) {
  await el.locator('p').last().click();
  await page.keyboard.press('End');
}

async function startTracking(page, name = 'Rita') {
  await page.getByRole('button', { name: 'Track changes' }).click();
  const dialog = page.getByRole('dialog', { name: 'Your name on tracked changes' });
  await dialog.getByRole('textbox').fill(name);
  await dialog.getByRole('button', { name: 'Start tracking' }).click();
  await expect(page.getByRole('button', { name: 'Track changes' })).toHaveAttribute('aria-pressed', 'true');
}

test('notes are shown below the text, numbered in the order the text refers to them', async ({ page }) => {
  await open(page);
  await expect(page.locator('.docx-notes[data-notes="footnote"]')).toHaveAttribute('data-label', 'Footnotes');
  await expect(footnote(page, '1')).toHaveText('The data were recieved in 2021.');
  await expect(footnote(page, '2')).toHaveText('Smith (2020), p. 4.');
  await expect(footnote(page, '1').locator('.note-mark')).toHaveAttribute('data-n', '1');
  await expect(endnote(page, 'i')).toHaveText('First endnote.');
  await expect(endnote(page, 'ii')).toHaveText('Second endnote.');
  await expect(page.locator('.docx-notes')).toHaveCount(2);
  // The labels and numbers are drawn, not text: nothing extra in the page or in the word count.
  await expect(page.locator('.statusbar')).toContainText('26 words');
});

test('untouched, the note parts and the text come back as they were', async ({ page }) => {
  await open(page);
  const { zip, xml } = await download(page);
  expect(await part(zip, 'footnotes')).toBe(FOOTNOTES);
  expect(await part(zip, 'endnotes')).toBe(ENDNOTES);
  expect(xml).toContain('<w:footnoteReference w:id="2"/>');
  expect(xml).not.toContain('Smith (2020)');
});

test('typing in a footnote changes only that note, and the review lists it', async ({ page }) => {
  await open(page);
  await endOf(page, footnote(page, '2'));
  await page.keyboard.type(' See also Jones (2021).');
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  const review = page.getByRole('dialog', { name: 'Review your changes' });
  await expect(review.locator('.change')).toHaveCount(1);
  await expect(review.locator('.change .where')).toContainText('Footnote 2, paragraph 1');
  await expect(review.locator('.change ins')).toHaveText(' See also Jones (2021).');
  await review.getByRole('button', { name: 'Keep editing' }).click();

  const { zip, xml } = await download(page);
  const notes = await part(zip, 'footnotes');
  expect(notes).toContain(`${SEPARATORS}${note('footnote', 1, 'Smith (2020), p. 4. See also Jones (2021).')}${note('footnote', 2, 'The data were recieved in 2021.')}`);
  expect(await part(zip, 'endnotes')).toBe(ENDNOTES);
  expect(xml).not.toContain('Jones');
});

test('Undo takes back typing in a note', async ({ page }) => {
  await open(page);
  await endOf(page, endnote(page, 'i'));
  await page.keyboard.type(' More.');
  await expect(endnote(page, 'i')).toHaveText('First endnote. More.');
  await page.keyboard.press('Control+z');
  await expect(endnote(page, 'i')).toHaveText('First endnote.');
  const { zip } = await download(page);
  expect(await part(zip, 'endnotes')).toBe(ENDNOTES);
});

test('Find and Replace reach the notes', async ({ page }) => {
  await open(page);
  await page.locator('.docx-page p').first().click();
  await page.keyboard.press('Control+h');
  await page.getByRole('textbox', { name: 'Find' }).fill('recieved');
  await page.getByRole('textbox', { name: 'Replace with' }).fill('received');
  await page.getByRole('button', { name: 'Replace all' }).click();
  await expect(footnote(page, '1')).toHaveText('The data were received in 2021.');
  const { zip } = await download(page);
  expect(await part(zip, 'footnotes')).toContain(note('footnote', 2, 'The data were received in 2021.'));
});

test('typing in a note while tracking is a tracked insertion in the note, with a new id', async ({ page }) => {
  await open(page);
  await startTracking(page);
  await endOf(page, footnote(page, '1'));
  await page.keyboard.type(' (n = 40)');
  await expect(footnote(page, '1').locator('.tc-ins')).toHaveText(' (n = 40)');
  const { zip } = await download(page);
  const ins = (await part(zip, 'footnotes')).match(/<w:ins w:id="(\d+)" w:author="Rita" w:date="[^"]+"><w:r><w:t xml:space="preserve"> \(n = 40\)<\/w:t><\/w:r><\/w:ins>/);
  expect(ins).not.toBeNull();
  // Revision ids are shared with the main document: the new one is above every id there and in the notes.
  expect(Number(ins[1])).toBeGreaterThan(2);
});

test('clicking a reference goes to its note, and the note’s number goes back', async ({ page }) => {
  await open(page);
  await reference(page, 'footnote', '2').click();
  await expect(status(page)).toHaveText('Footnote 2: type to edit it – click its number to go back to the text');
  await page.keyboard.type('!');
  await expect(footnote(page, '2')).toHaveText('Smith (2020), p. 4.!');
  await footnote(page, '2').locator('.note-mark').click();
  await expect(status(page)).toHaveText('Back in the text, after footnote 2');
  await page.keyboard.type('*');
  await expect(page.locator('.docx-page > p').first()).toHaveText('Prior work agrees. So do others.*');
});

test('a note, and its number, can’t be deleted with the text around them', async ({ page }) => {
  await open(page);
  // From the text into the notes.
  await page.locator('.docx-page > p').last().click();
  await page.evaluate(() => {
    const sel = window.getSelection();
    const range = document.createRange();
    range.setStart(document.querySelector('.docx-page > p').firstChild, 0);
    range.setEnd(document.querySelector('.docx-note p').lastChild, 3);
    sel.removeAllRanges();
    sel.addRange(range);
  });
  await page.keyboard.press('Delete');
  await expect(status(page)).toContainText('Not deleted: that would remove');
  await expect(page.locator('.docx-page > p').first()).toContainText('Prior work agrees.');
  // The number at the start of a note.
  const first = footnote(page, '1');
  await first.locator('p').click();
  await page.keyboard.press('Home');
  await page.keyboard.press('Backspace');
  await page.keyboard.press('Backspace');
  await expect(status(page)).toHaveText('Not deleted: the note’s number stays – Word numbers notes. To remove the note, delete its number in the text.');
  await expect(first.locator('.note-mark')).toHaveCount(1);
});

test('deleting a reference hides its note and renumbers the others; Undo brings it back', async ({ page }) => {
  await open(page);
  await reference(page, 'footnote', '1').click(); // goes to the note: come back beside the reference
  await footnote(page, '1').locator('.note-mark').click();
  await page.keyboard.press('Backspace');
  await expect(reference(page, 'footnote', '1')).toHaveAttribute('data-note-id', '1');
  await expect(page.locator('.docx-note[data-note="footnote"][data-note-id="2"]')).toBeHidden();
  await expect(footnote(page, '1')).toHaveText('Smith (2020), p. 4.');
  await page.keyboard.press('Control+z');
  await expect(page.locator('.docx-note[data-note="footnote"][data-note-id="2"]')).toBeVisible();
  await expect(footnote(page, '2')).toHaveText('Smith (2020), p. 4.');
});

test('pictures, tables, comments and page breaks don’t go in a note', async ({ page }) => {
  await open(page);
  await endOf(page, footnote(page, '1'));
  await page.keyboard.press('Control+Enter');
  await expect(status(page)).toHaveText('A page break can’t go in a footnote here: add it in Word');
  await page.keyboard.press('Control+Alt+m');
  await expect(status(page)).toHaveText('A new comment can’t go in a footnote here: add it in Word');
  const { zip } = await download(page);
  expect(await part(zip, 'footnotes')).toBe(FOOTNOTES);
});

test('a comment in a note keeps its place when its text is deleted, as in the text', async ({ page }) => {
  const commented = `${XML}<w:footnotes ${W_NS}>${SEPARATORS}<w:footnote w:id="1"><w:p>${mark('footnoteRef')}<w:r><w:t xml:space="preserve"> Keep </w:t></w:r><w:commentRangeStart w:id="6"/><w:r><w:t>this</w:t></w:r><w:commentRangeEnd w:id="6"/><w:r><w:commentReference w:id="6"/></w:r></w:p></w:footnote></w:footnotes>`;
  const comments = `${XML}<w:comments ${W_NS}><w:comment w:id="6" w:author="Cy" w:date="2026-01-02T00:00:00Z"><w:p><w:r><w:t>About the note.</w:t></w:r></w:p></w:comment></w:comments>`;
  const opts = {
    parts: { 'word/footnotes.xml': commented, 'word/comments.xml': comments },
    rels: `<Relationship Id="rIdF" Type="${REL}/footnotes" Target="footnotes.xml"/><Relationship Id="rIdC" Type="${REL}/comments" Target="comments.xml"/>`,
  };
  await openInEditor(page, await buildDocx(`<w:p><w:r><w:t>Body</w:t></w:r>${ref('footnote', 1)}</w:p>`, opts));
  await page.locator('.docx-note p').click();
  await page.keyboard.press('End');
  for (let i = 0; i < 5; i++) await page.keyboard.press('Backspace');
  await expect(page.locator('.docx-note')).toHaveText('Keep');
  const { zip } = await download(page);
  // Its range shrinks to where the text was; the comment is the comments panel's to delete.
  expect(await part(zip, 'footnotes')).toContain('<w:t xml:space="preserve"> Keep</w:t></w:r><w:commentRangeStart w:id="6"/><w:commentRangeEnd w:id="6"/><w:r><w:commentReference w:id="6"/></w:r>');
  expect(await part(zip, 'comments')).toBe(comments);
});

test('a note edited and a comment in it deleted from the panel: both are saved in the note', async ({ page }) => {
  const commented = `${XML}<w:footnotes ${W_NS}>${SEPARATORS}<w:footnote w:id="1"><w:p>${mark('footnoteRef')}<w:r><w:t xml:space="preserve"> A </w:t></w:r><w:commentRangeStart w:id="6"/><w:r><w:t>note</w:t></w:r><w:commentRangeEnd w:id="6"/><w:r><w:commentReference w:id="6"/></w:r></w:p></w:footnote></w:footnotes>`;
  const comments = `${XML}<w:comments ${W_NS}><w:comment w:id="6" w:author="Cy" w:date="2026-01-02T00:00:00Z"><w:p><w:r><w:t>About the note.</w:t></w:r></w:p></w:comment></w:comments>`;
  const opts = {
    parts: { 'word/footnotes.xml': commented, 'word/comments.xml': comments },
    rels: `<Relationship Id="rIdF" Type="${REL}/footnotes" Target="footnotes.xml"/><Relationship Id="rIdC" Type="${REL}/comments" Target="comments.xml"/>`,
  };
  await page.setViewportSize({ width: 1280, height: 800 });
  await openInEditor(page, await buildDocx(`<w:p><w:r><w:t>Body</w:t></w:r>${ref('footnote', 1)}</w:p>`, opts));
  await page.locator('.docx-note p').click();
  await page.keyboard.press('End');
  await page.keyboard.type(', revised');
  const thread = page.getByRole('complementary', { name: 'Comments' }).locator('.thread', { hasText: 'About the note.' });
  await expect(thread.locator('.tag.muted')).toHaveText('In a footnote');
  await thread.locator('.thread-actions').getByRole('button', { name: 'Delete' }).click();
  const { zip } = await download(page);
  const notes = await part(zip, 'footnotes');
  // The comment's marks go from between the runs (which Word reads as one text).
  expect(notes).toContain('<w:t xml:space="preserve"> A </w:t></w:r><w:r><w:t xml:space="preserve">note, revised</w:t></w:r></w:p></w:footnote>');
  expect(notes).not.toContain('comment');
  expect(await part(zip, 'comments')).not.toContain('About the note.');
});

test('Select All in the text selects the text, not the notes: typing replaces only the text', async ({ page }) => {
  await open(page);
  await page.locator('.docx-page > p').first().click();
  await page.keyboard.press('ControlOrMeta+a');
  await page.keyboard.type('New text.');
  await expect(page.locator('.docx-page > p')).toHaveText(['New text.']);
  await expect(footnote(page, '1')).toBeHidden(); // its reference went with the text
  const { zip } = await download(page);
  expect(await part(zip, 'footnotes')).toBe(FOOTNOTES);
});
