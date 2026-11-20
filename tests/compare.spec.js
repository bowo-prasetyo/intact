// Compare documents (Review > Compare), as Word's Compare does for the text: another version of the
// document, chosen on this device, against the one open; each paragraph changed, added or taken out,
// its words struck through or underlined, and a click goes to it. Nothing is changed or sent.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor } from './helpers.js';

const para = (text) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const OLD = [
  'Kepada Yth. Kepala Dinas Pendidikan',
  'Rapat koordinasi diadakan pada hari Senin di ruang rapat lantai 2.',
  'Peserta diharap membawa laptop.',
  'Atas perhatiannya kami ucapkan terima kasih.',
].map(para).join('')
  // A tracked deletion in the older version doesn't count as its text; a table cell's text does.
  + '<w:p><w:del w:id="9" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>Dihapus.</w:delText></w:r></w:del></w:p>'
  + '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Sel tabel</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>';
const NEW = [
  'Kepada Yth. Kepala Dinas Pendidikan',
  'Rapat koordinasi diadakan pada hari Selasa pukul 09.00 di ruang rapat lantai 2.',
  'Atas perhatiannya kami ucapkan terima kasih.',
  'Tembusan: Sekretaris Daerah.',
].map(para).join('') + '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Sel tabel</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>';
const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

async function compareWith(page, buffer, name = 'Undangan v1.docx') {
  const chooser = page.waitForEvent('filechooser');
  await page.getByRole('button', { name: 'Compare documents' }).click();
  await (await chooser).setFiles({ name, mimeType: DOCX, buffer });
}

test('the changes from the other version, paragraph by paragraph and word by word; a click goes to one', async ({ page }) => {
  await openInEditor(page, await buildDocx(NEW), 'Undangan v2.docx');
  await compareWith(page, await buildDocx(OLD));
  const dialog = page.getByRole('dialog', { name: 'Compare documents' });
  await expect(dialog).toContainText('From Undangan v1.docx to this document');
  await expect(dialog.getByRole('status')).toHaveText('1 changed, 1 added, 1 taken out (paragraphs)');
  const items = dialog.getByRole('listitem');
  await expect(items).toHaveCount(3);
  await expect(items.nth(0).locator('del')).toHaveText('Senin');
  await expect(items.nth(0).locator('ins')).toHaveText('Selasa pukul 09.00');
  await expect(items.nth(1)).toContainText('Taken out');
  await expect(items.nth(1).locator('del')).toHaveText('Peserta diharap membawa laptop.');
  await expect(items.nth(2)).toContainText('Added');
  await expect(items.nth(2).locator('ins')).toHaveText('Tembusan: Sekretaris Daerah.');
  // To the paragraph that changed, with the caret in it.
  await items.nth(0).getByRole('button').click();
  await expect(dialog).toBeHidden();
  await expect(page.locator('.statusbar [role=status]')).toHaveText('Changed in this document');
  expect(await page.evaluate(() => (() => { const n = window.getSelection().anchorNode; return (n.nodeType === 1 ? n : n.parentElement).closest('p').textContent; })())).toContain('Selasa pukul 09.00');
  // Nothing in the document changed.
  await expect(page.locator('.app-header .dirty')).toHaveCount(0);
});

test('the same text says so; a file that isn\'t a .docx says what it can\'t read', async ({ page }) => {
  await openInEditor(page, await buildDocx(NEW));
  await compareWith(page, await buildDocx(NEW), 'Same.docx');
  const dialog = page.getByRole('dialog', { name: 'Compare documents' });
  await expect(dialog.getByRole('status')).toHaveText('The text is the same in both.');
  await dialog.getByRole('button', { name: 'Close' }).last().click();
  // A Word 97–2003 .doc: a compound file with a WordDocument stream. It says what it is, as opening one does.
  const doc = Buffer.concat([Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]), Buffer.alloc(56), Buffer.from('WordDocument', 'utf16le')]);
  await compareWith(page, doc, 'Lama.doc');
  await expect(page.locator('.statusbar [role=status]')).toHaveText(/^Lama\.doc can’t be compared: This is a Word 97–2003 document \(\.doc\)/);
});
