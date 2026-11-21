// Compare documents looks in the footnotes and endnotes too: each kind in the order the text refers
// to them, numbered as shown, compared word by word; a click goes to the note. The redline marks the
// body's changes only, and says so.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, W_NS } from './helpers.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const SEP = '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>';
const mark = (local) => `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:${local}/></w:r>`;
const note = (local, id, text) => `<w:${local} w:id="${id}"><w:p>${mark(`${local}Ref`)}<w:r><w:t xml:space="preserve"> ${text}</w:t></w:r></w:p></w:${local}>`;
const ref = (local, id) => `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:${local}Reference w:id="${id}"/></w:r>`;
const P = (text, refs = '') => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r>${refs}</w:p>`;

/** A version of a chapter: its footnotes' texts (by id), in the order given by `order`, and an endnote. */
function version({ footnotes, order, endnote }) {
  const body = P('Prior work agrees.', order.map((id) => ref('footnote', id)).join('')) + P('See the appendix.', ref('endnote', 1));
  return buildDocx(body, {
    parts: {
      'word/footnotes.xml': `${XML}<w:footnotes ${W_NS}>${SEP}${Object.entries(footnotes).map(([id, t]) => note('footnote', id, t)).join('')}</w:footnotes>`,
      'word/endnotes.xml': `${XML}<w:endnotes ${W_NS}>${note('endnote', 1, endnote)}</w:endnotes>`,
    },
    rels: `<Relationship Id="rIdF" Type="${REL}/footnotes" Target="footnotes.xml"/><Relationship Id="rIdE" Type="${REL}/endnotes" Target="endnotes.xml"/>`,
  });
}

async function compareWith(page, buffer, name = 'Chapter v1.docx') {
  const chooser = page.waitForEvent('filechooser');
  await page.getByRole('button', { name: 'Compare documents' }).click();
  await (await chooser).setFiles({ name, mimeType: DOCX, buffer });
}

test('changes in footnotes and endnotes are listed by note, word by word; a click goes to the note', async ({ page }) => {
  const older = await version({ footnotes: { 1: 'Smith (2020), p. 4.', 2: 'Lee (2019).', 3: 'A note later cut.' }, order: [1, 2, 3], endnote: 'Data from 2021.' });
  // The newer one: footnote 1 changed, footnote 3 cut and a new last one added; the endnote changed.
  const newer = await version({ footnotes: { 1: 'Smith (2020), p. 6.', 2: 'Lee (2019).', 5: 'Park (2023) disagrees.' }, order: [1, 2, 5], endnote: 'Data from 2022.' });
  await openInEditor(page, newer, 'Chapter v2.docx');
  await compareWith(page, older);
  const dialog = page.getByRole('dialog', { name: 'Compare documents' });
  await expect(dialog.getByRole('status')).toHaveText('The text of the body is the same in both.In the notes: 3 changed, 0 added, 0 taken out');
  const notes = dialog.getByRole('list', { name: 'Footnotes and endnotes' }).getByRole('listitem');
  await expect(notes).toHaveCount(3);
  await expect(notes.nth(0)).toContainText('Footnote 1');
  await expect(notes.nth(0).locator('del')).toHaveText('4.');
  await expect(notes.nth(0).locator('ins')).toHaveText('6.');
  await expect(notes.nth(1)).toContainText('Footnote 3');
  // Word by word: what was struck through and what was put in.
  expect((await notes.nth(1).locator('del').allTextContents()).join(' ')).toMatch(/A.*note.*later.*cut\./);
  expect((await notes.nth(1).locator('ins').allTextContents()).join(' ')).toMatch(/Park.*\(2023\).*disagrees\./);
  await expect(notes.nth(2)).toContainText('Endnote i');
  await expect(notes.nth(2).locator('ins')).toHaveText('2022.');
  // The body's changes alone make a redline: none here, so it isn't offered.
  await expect(dialog.getByRole('button', { name: 'Open redline' })).toHaveCount(0);

  await notes.nth(2).getByRole('button').click();
  await expect(dialog).toBeHidden();
  await expect(page.locator('.statusbar [role=status]')).toHaveText('Changed in this document: endnote i');
  expect(await page.evaluate(() => { const n = window.getSelection().anchorNode; return (n.nodeType === 1 ? n : n.parentElement).closest('.docx-note')?.textContent; })).toContain('Data from 2022.');
});

test('a note added or taken out, and a body change too: the redline says it marks the body only', async ({ page }) => {
  const older = await version({ footnotes: { 1: 'Smith (2020).' }, order: [1], endnote: 'Same.' });
  const newerBody = P('Prior work mostly agrees.', ref('footnote', 1) + ref('footnote', 2)) + P('See the appendix.', ref('endnote', 1));
  const newer = await buildDocx(newerBody, {
    parts: {
      'word/footnotes.xml': `${XML}<w:footnotes ${W_NS}>${SEP}${note('footnote', 1, 'Smith (2020).')}${note('footnote', 2, 'Lee (2019).')}</w:footnotes>`,
      'word/endnotes.xml': `${XML}<w:endnotes ${W_NS}>${note('endnote', 1, 'Same.')}</w:endnotes>`,
    },
    rels: `<Relationship Id="rIdF" Type="${REL}/footnotes" Target="footnotes.xml"/><Relationship Id="rIdE" Type="${REL}/endnotes" Target="endnotes.xml"/>`,
  });
  await openInEditor(page, newer, 'Chapter v2.docx');
  await compareWith(page, older);
  const dialog = page.getByRole('dialog', { name: 'Compare documents' });
  await expect(dialog.getByRole('status')).toHaveText('1 changed, 0 added, 0 taken out (paragraphs)In the notes: 0 changed, 1 added, 0 taken out');
  const notes = dialog.getByRole('list', { name: 'Footnotes and endnotes' }).getByRole('listitem');
  await expect(notes).toHaveText([/Footnote 2, added\s*Lee \(2019\)\./]);
  await expect(dialog).toContainText('Changes in footnotes and endnotes are listed here, not marked in it.');

  // The other way round: the note was taken out.
  await dialog.getByRole('button', { name: 'Close' }).last().click();
  await page.goto('about:blank');
  await openInEditor(page, older, 'Chapter v1.docx');
  await compareWith(page, newer, 'Chapter v2.docx');
  await expect(dialog.getByRole('list', { name: 'Footnotes and endnotes' }).getByRole('listitem')).toHaveText([/Footnote, taken out\s*Lee \(2019\)\./]);
});

test('documents without notes, the same in both, say the text is the same', async ({ page }) => {
  const same = await buildDocx(P('One paragraph.'));
  await openInEditor(page, same);
  await compareWith(page, same);
  const dialog = page.getByRole('dialog', { name: 'Compare documents' });
  await expect(dialog.getByRole('status')).toHaveText('The text is the same in both.');
  await expect(dialog.getByRole('list', { name: 'Footnotes and endnotes' })).toHaveCount(0);
});
