// Documents encrypted with a password (Word's File › Info › Protect Document › Encrypt with Password):
// opened with the password, on this device; edited; downloaded encrypted with it again, as Word saves
// them; a password given, changed or taken off; drafts kept encrypted; and a version to compare that
// is encrypted. The fixtures in tests/fixtures/encrypted were encrypted by other programs (see the
// README there): msoffcrypto-tool (Agile Encryption, as Word 2010 and later) and LibreOffice
// (Standard Encryption, as Office 2007).
import fs from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor } from './helpers.js';
import { readCfb } from '../src/docx/cfb.js';

const DIR = path.join(import.meta.dirname, 'fixtures', 'encrypted');
const AGILE = fs.readFileSync(path.join(DIR, 'report-password.docx'));
const AGILE_PASSWORD = 'Rahasia 2026!';
const STANDARD = fs.readFileSync(path.join(DIR, 'libreoffice-memo-password.docx'));
const STANDARD_PASSWORD = 'Memo-Pass 7';
const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const CFB = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

const form = (page) => page.getByRole('form', { name: 'Protected with a password' });
const bar = (page) => page.getByRole('region', { name: 'Password protection' });
const status = (page) => page.locator('.statusbar [role=status]');
const isCompoundFile = (buf) => CFB.every((b, i) => buf[i] === b);
/** Major and minor version of the file's EncryptionInfo: 4.4 is Agile, as Word 2010 and later write. */
const encryptionVersion = (buf) => {
  const info = readCfb(new Uint8Array(buf)).streams.get('EncryptionInfo');
  return `${info[0]}.${info[2]}`;
};

async function pick(page, buffer, name) {
  await page.setInputFiles('input[type=file]', { name, mimeType: DOCX, buffer });
}

async function openWithPassword(page, buffer, password, name = 'Report.docx') {
  if (!page.listenerCount('dialog')) page.on('dialog', (d) => d.accept());
  await page.goto('./');
  await pick(page, buffer, name);
  await expect(form(page)).toBeVisible();
  await page.getByLabel(`Password for ${name}`).fill(password);
  await form(page).getByRole('button', { name: 'Open' }).click();
  await expect(page.locator('.docx-page')).toBeVisible();
}

/** Download, through the review; returns the file's bytes and the review's text. */
async function downloadBytes(page) {
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  const review = page.getByRole('dialog', { name: 'Review your changes' });
  await expect(review).toBeVisible();
  const said = await review.textContent(); // folded parts too
  const [dl] = await Promise.all([page.waitForEvent('download'), review.getByRole('button', { name: 'Download' }).click()]);
  return { bytes: fs.readFileSync(await dl.path()), said };
}

async function typeAtEnd(page, text) {
  const p = page.locator('.docx-page p').filter({ hasText: /\S/ }).first();
  await p.click();
  await page.keyboard.press('End');
  await page.keyboard.type(text);
}

test('a document encrypted with a password asks for it, refuses a wrong one, and opens with the right one', async ({ page }) => {
  await page.goto('./');
  await pick(page, AGILE, 'Report.docx');
  await expect(form(page)).toContainText('never sent anywhere');
  const input = page.getByLabel('Password for Report.docx');
  await expect(input).toBeFocused();
  await expect(input).toHaveAttribute('type', 'password');
  await input.fill('rahasia 2026!'); // case matters
  await input.press('Enter');
  await expect(form(page).getByRole('alert')).toHaveText('That password isn’t right. Passwords are case-sensitive: check Caps Lock.');
  await expect(input).toHaveAttribute('aria-invalid', 'true');
  await input.fill(AGILE_PASSWORD);
  await input.press('Enter');
  await expect(page.locator('.docx-page')).toContainText('Quarterly Field Report');
  await expect(bar(page)).toContainText('Protected with a password.');
  // Cancel goes back to the start page's choices without opening anything.
  await page.goto('./');
  await pick(page, AGILE, 'Report.docx');
  await form(page).getByRole('button', { name: 'Cancel' }).click();
  await expect(form(page)).toHaveCount(0);
});

test('edited, it downloads encrypted with the same password, and opens again only with it', async ({ page }) => {
  await openWithPassword(page, AGILE, AGILE_PASSWORD);
  await typeAtEnd(page, ' Approved by the board.');
  const { bytes, said } = await downloadBytes(page);
  expect(said).toContain('The downloaded file stays encrypted with its password, as Word saves it.');
  await expect(status(page)).toContainText('encrypted with its password');
  expect(isCompoundFile(bytes)).toBe(true);
  expect(encryptionVersion(bytes)).toBe('4.4');
  // A new salt each time: not the bytes opened.
  expect(bytes.equals(AGILE)).toBe(false);
  // Reopened: the password is asked for, and the edit is there.
  await openWithPassword(page, bytes, AGILE_PASSWORD);
  await expect(page.locator('.docx-page')).toContainText('Approved by the board.');
});

test('unchanged, the encrypted file downloads exactly as it was opened', async ({ page }) => {
  await openWithPassword(page, AGILE, AGILE_PASSWORD);
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  const review = page.getByRole('dialog', { name: 'Review your changes' });
  const [dl] = await Promise.all([page.waitForEvent('download'), review.getByRole('button', { name: 'Download' }).click()]);
  expect(fs.readFileSync(await dl.path()).equals(AGILE)).toBe(true);
});

test('a document encrypted as Office 2007 and LibreOffice do opens, and is encrypted again as Word does now', async ({ page }) => {
  expect(encryptionVersion(STANDARD)).toBe('3.2');
  await openWithPassword(page, STANDARD, STANDARD_PASSWORD, 'Memo.docx');
  await expect(page.locator('.docx-page')).toContainText('Internal memo');
  const { bytes } = await downloadBytes(page);
  expect(encryptionVersion(bytes)).toBe('4.4');
  await openWithPassword(page, bytes, STANDARD_PASSWORD, 'Memo.docx');
  await expect(page.locator('.docx-page')).toContainText('Internal memo');
});

test('any document can be protected with a password, which can be changed or taken off', async ({ page }) => {
  await openInEditor(page, await buildDocx('<w:p><w:r><w:t>Board paper: confidential.</w:t></w:r></w:p>'), 'Board.docx');
  await expect(bar(page)).toHaveCount(0);
  await page.getByRole('button', { name: 'More ways to save' }).click();
  await page.getByRole('menuitem', { name: /Protect with password/ }).click();
  const dialog = page.getByRole('dialog', { name: 'Protect with password' });
  await expect(dialog).toContainText('a lost password can’t be recovered');
  await dialog.getByLabel('Password', { exact: true }).fill('Board#1');
  await dialog.getByLabel('Type it again').fill('Board#2');
  await dialog.getByRole('button', { name: 'Protect' }).click();
  await expect(dialog.getByRole('alert')).toHaveText('The two passwords aren’t the same. Type the same password in both.');
  await dialog.getByLabel('Type it again').fill('Board#1');
  await dialog.getByRole('button', { name: 'Protect' }).click();
  await expect(dialog).toHaveCount(0);
  await expect(status(page)).toHaveText('Protected with a password: the downloaded file opens only with it');
  await expect(bar(page)).toBeVisible();
  await expect(page.locator('.app-header .dirty')).toBeVisible();
  let { bytes, said } = await downloadBytes(page);
  expect(said).toContain('The downloaded file is encrypted with the password you set');
  expect(isCompoundFile(bytes)).toBe(true);
  await openWithPassword(page, bytes, 'Board#1', 'Board.docx');
  await expect(page.locator('.docx-page')).toContainText('Board paper: confidential.');

  // Taken off: the file is a plain .docx again.
  await bar(page).getByRole('button', { name: 'Change or remove password' }).click();
  await page.getByRole('dialog', { name: 'Change or remove password' }).getByRole('button', { name: 'Remove password' }).click();
  await expect(bar(page)).toHaveCount(0);
  ({ bytes, said } = await downloadBytes(page));
  expect(said).toContain('The password is removed');
  expect(bytes[0]).toBe(0x50); // PK: a ZIP package
  expect(await (await JSZip.loadAsync(bytes)).file('word/document.xml').async('string')).toContain('Board paper: confidential.');
});

test('a protected document’s unsaved changes are kept encrypted, and restoring them asks for the password', async ({ page }) => {
  await openWithPassword(page, AGILE, AGILE_PASSWORD);
  await typeAtEnd(page, ' Draft note.');
  await expect(page.locator('.statusbar').getByText('Changes kept on this device')).toBeVisible();
  // The copy in IndexedDB is a compound file, encrypted: nothing of the text in the clear.
  const head = await page.evaluate(() => new Promise((resolve) => {
    const req = indexedDB.open('intact');
    req.onsuccess = () => {
      const all = req.result.transaction('drafts').objectStore('drafts').getAll();
      all.onsuccess = async () => resolve([...new Uint8Array(await all.result[0].blob.arrayBuffer()).subarray(0, 8)]);
    };
  }));
  expect(head).toEqual(CFB);
  await page.reload();
  const drafts = page.getByRole('region', { name: 'Unsaved changes on this device' });
  await expect(drafts).toContainText('encrypted with its password');
  await drafts.getByRole('button', { name: 'Restore Report.docx' }).click();
  await page.getByLabel('Password for Report.docx').fill(AGILE_PASSWORD);
  await form(page).getByRole('button', { name: 'Open' }).click();
  await expect(page.locator('.docx-page')).toContainText('Draft note.');
  await expect(bar(page)).toBeVisible();
});

test('an encrypted version to compare asks for its password', async ({ page }) => {
  await openWithPassword(page, AGILE, AGILE_PASSWORD);
  await typeAtEnd(page, ' Revised.');
  const chooser = page.waitForEvent('filechooser');
  await page.getByRole('button', { name: 'Compare documents' }).click();
  await (await chooser).setFiles({ name: 'Report v1.docx', mimeType: DOCX, buffer: AGILE });
  const dialog = page.getByRole('dialog', { name: 'Password for Report v1.docx' });
  await dialog.getByLabel('Password', { exact: true }).fill('wrong');
  await dialog.getByRole('button', { name: 'Compare' }).click();
  await expect(dialog.getByRole('alert')).toContainText('That password isn’t right');
  await dialog.getByLabel('Password', { exact: true }).fill(AGILE_PASSWORD);
  await dialog.getByRole('button', { name: 'Compare' }).click();
  const result = page.getByRole('dialog', { name: 'Compare documents' });
  await expect(result.getByRole('status')).toHaveText('1 changed, 0 added, 0 taken out (paragraphs)');
  await expect(result.locator('ins')).toHaveText(' Revised.');
});
