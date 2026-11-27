// A redline from Compare documents, as Word's Compare makes one: a new document, this version with
// the changes since the other version written into it as tracked changes (docx/redline.js). Checked
// in its XML, and by accepting all its changes (this version's text) and rejecting them all (the
// other version's text). Compare reads both versions from their XML, so a document compared with
// itself is the same, as the real fixtures show.
import fs from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, downloadNames } from './helpers.js';

const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const p = (text, pPr = '') => `<w:p>${pPr}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const OLD = [
  p('Board paper: new office lease'),
  p('The lease runs for five years from 1 March.'),
  p('Rent is reviewed every year.'),
  p('Questions go to the facilities team.'),
].join('');
// This version: a word changed (in bold, which the inserted run keeps), a paragraph taken out, and
// one added at the end.
const NEW = [
  p('Board paper: new office lease'),
  '<w:p><w:r><w:t xml:space="preserve">The lease runs for </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>ten</w:t></w:r><w:r><w:t xml:space="preserve"> years from 1 March.</w:t></w:r></w:p>',
  p('Questions go to the facilities team.'),
  p('The landlord pays for repairs.'),
].join('');

async function compareWith(page, buffer, name = 'Lease v1.docx') {
  const chooser = page.waitForEvent('filechooser');
  await page.getByRole('button', { name: 'Compare documents' }).click();
  await (await chooser).setFiles({ name, mimeType: DOCX, buffer });
  return page.getByRole('dialog', { name: 'Compare documents' });
}

async function downloadRedline(page, dialog) {
  const [dl] = await Promise.all([page.waitForEvent('download'), dialog.getByRole('button', { name: 'Download redline' }).click()]);
  return { bytes: fs.readFileSync(await dl.path()), name: dl.suggestedFilename() };
}

/** Opens `buffer` in the page already set up by openInEditor (whose dialog handler stays). */
async function reopen(page, buffer, name) {
  await page.goto('./');
  await page.setInputFiles('input[type=file]', { name, mimeType: DOCX, buffer });
  await expect(page.locator('.docx-page')).toBeVisible();
}

const named = (page, name) => page.addInitScript((n) => localStorage.setItem('intact.author', n), name);
const bodyText = (page) => page.locator('.docx-page').evaluate((el) => [...el.querySelectorAll('p')].map((x) => x.textContent).filter((x) => x.trim()));

test('a redline marks the changes since the other version as tracked changes, in your name', async ({ page, browserName }) => {
  await named(page, 'Dana Executive');
  await openInEditor(page, await buildDocx(NEW), 'Lease v2.docx');
  const dialog = await compareWith(page, await buildDocx(OLD));
  await expect(dialog.getByRole('status')).toHaveText('1 changed, 1 added, 1 taken out (paragraphs)');
  await expect(dialog).toContainText('marked as tracked changes by Dana Executive');
  const { bytes, name } = await downloadRedline(page, dialog);
  expect(downloadNames('Lease v2 (redline).docx', browserName)).toContain(name);
  await expect(dialog).toBeHidden();
  await expect(page.locator('.statusbar [role=status]')).toHaveText('Redline downloaded: Lease v2 (redline).docx');
  // This document itself is unchanged.
  await expect(page.locator('.app-header .dirty')).toHaveCount(0);

  const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml').async('string');
  const dom = await page.evaluate((x) => {
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const d = new DOMParser().parseFromString(x, 'application/xml');
    const all = (local) => [...d.getElementsByTagNameNS(W, local)];
    const inline = (local) => all(local).filter((e) => e.parentNode.localName !== 'rPr');
    return {
      ins: inline('ins').map((e) => ({ text: e.textContent, bold: !!e.getElementsByTagNameNS(W, 'b').length })),
      del: inline('del').map((e) => [...e.getElementsByTagNameNS(W, 'delText')].map((t) => t.textContent).join('')),
      marks: all('ins').concat(all('del')).filter((e) => e.parentNode.localName === 'rPr').map((e) => e.localName),
      authors: [...new Set(all('ins').concat(all('del')).map((e) => e.getAttributeNS(W, 'author')))],
      ids: all('ins').concat(all('del')).map((e) => e.getAttributeNS(W, 'id')),
      dates: all('ins').map((e) => e.getAttributeNS(W, 'date')),
    };
  }, xml);
  expect(dom.ins).toEqual([{ text: 'ten', bold: true }, { text: 'The landlord pays for repairs.', bold: false }]);
  expect(dom.del).toEqual(['five', 'Rent is reviewed every year.']);
  expect(dom.marks.sort()).toEqual(['del', 'ins']); // the paragraph taken out, and the one added
  expect(dom.authors).toEqual(['Dana Executive']);
  expect(new Set(dom.ids).size).toBe(dom.ids.length);
  expect(dom.dates[0]).toMatch(/^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\dZ$/);

  // Accepted, it is this version; rejected, the other.
  await reopen(page, bytes, 'Lease v2 (redline).docx');
  await page.getByTitle('More ways to accept').click();
  await page.getByRole('menuitem', { name: 'Accept all changes' }).click();
  await expect(page.locator('.statusbar [role=status]')).toHaveText(/^Accepted all \d+ changes$/);
  expect(await bodyText(page)).toEqual(['Board paper: new office lease', 'The lease runs for ten years from 1 March.', 'Questions go to the facilities team.', 'The landlord pays for repairs.']);
  await reopen(page, bytes, 'Lease v2 (redline).docx');
  await page.getByTitle('More ways to reject').click();
  await page.getByRole('menuitem', { name: 'Reject all changes' }).click();
  expect(await bodyText(page)).toEqual(['Board paper: new office lease', 'The lease runs for five years from 1 March.', 'Rent is reviewed every year.', 'Questions go to the facilities team.']);
});

test('tracked changes already in this version are accepted in the redline, as Word’s Compare does', async ({ page }) => {
  const tracked = '<w:p><w:r><w:t xml:space="preserve">Rent is </w:t></w:r><w:ins w:id="1" w:author="Legal" w:date="2026-01-01T00:00:00Z"><w:r><w:t xml:space="preserve">not </w:t></w:r></w:ins>'
    + '<w:r><w:t xml:space="preserve">reviewed every year.</w:t></w:r><w:del w:id="2" w:author="Legal" w:date="2026-01-01T00:00:00Z"><w:r><w:delText> Ever.</w:delText></w:r></w:del></w:p>';
  await openInEditor(page, await buildDocx(p('Board paper') + tracked), 'Lease v3.docx');
  const dialog = await compareWith(page, await buildDocx(p('Board paper') + p('Rent is reviewed every year.')));
  const { bytes } = await downloadRedline(page, dialog);
  await expect(page.locator('.statusbar [role=status]')).toContainText('the 2 tracked changes already in this document are accepted in it');
  const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml').async('string');
  expect(xml).not.toContain('w:author="Legal"');
  expect(xml).not.toContain('Ever.');
  expect(xml).toMatch(/<w:ins [^>]*w:author="Author"[^>]*><w:r><w:t xml:space="preserve">not <\/w:t><\/w:r><\/w:ins>/);
});

test('the redline opens in a new tab, beside this document', async ({ page, context }) => {
  await openInEditor(page, await buildDocx(NEW), 'Lease v2.docx');
  const dialog = await compareWith(page, await buildDocx(OLD));
  const [tab] = await Promise.all([context.waitForEvent('page'), dialog.getByRole('button', { name: 'Open redline' }).click()]);
  await expect(tab.locator('.docx-page')).toContainText('The landlord pays for repairs.');
  await expect(tab.locator('.app-header .filename')).toHaveValue('Lease v2 (redline)');
  await expect(tab.locator('.docx-page')).toContainText('fiveten years'); // the change marked: both words shown
  await expect(page.locator('.statusbar [role=status]')).toHaveText('The redline is opening in a new tab');
});

test('the same text offers no redline', async ({ page }) => {
  await openInEditor(page, await buildDocx(NEW));
  const dialog = await compareWith(page, await buildDocx(NEW));
  await expect(dialog.getByRole('status')).toHaveText('The text is the same in both.');
  await expect(dialog.getByRole('button', { name: /redline/ })).toHaveCount(0);
});

// Real documents compared with themselves: the text is the same (it used to differ where the page
// shows a line break, a note's reference or a tracked change differently from the XML).
const FIXTURES = path.join(import.meta.dirname, 'fixtures');
for (const file of fs.readdirSync(FIXTURES).filter((f) => f.endsWith('.docx'))) {
  test(`${file} compared with itself is the same`, async ({ page }) => {
    const buffer = fs.readFileSync(path.join(FIXTURES, file));
    await openInEditor(page, buffer, file);
    const dialog = await compareWith(page, buffer, file);
    await expect(dialog.getByRole('status')).toHaveText('The text is the same in both.');
  });
}
