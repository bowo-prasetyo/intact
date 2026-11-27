// Versioned download names (editor/versionName.js and the review's "Name it as a new version"):
// where the file opened can't be saved back to, a changed document downloads as its next version
// ("Report v2", "Report_v5_JD") rather than as the browser's "Report (1)", when asked for; the
// choice is remembered, and the header's name follows, so the next download is the version after.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, downloadNames } from './helpers.js';
import { nextVersionName, initialsOf } from '../src/editor/versionName.js';

test('the next version of a name: its number one up, written as it is, or v2 added', () => {
  const cases = [
    ['Report', '', 'Report v2'],
    ['Q3_board_paper', '', 'Q3_board_paper_v2'], // the separator the name uses
    ['Board-paper', '', 'Board-paper-v2'],
    ['Report v2', '', 'Report v3'],
    ['Lease V09', '', 'Lease V10'], // its case and its leading zero
    ['Contract rev 3', '', 'Contract rev 4'],
    ['Contract rev3', '', 'Contract rev4'],
    ['Minutes - v1.2', '', 'Minutes - v1.3'],
    ['Plan_version_7', '', 'Plan_version_8'],
    ['Budget 2026', '', 'Budget 2026 v2'], // a year isn't a version
    ['Report_v4_JD', '', 'Report_v5_JD'], // the initials there were stay…
    ['Report_v4_JD', 'DE', 'Report_v5_DE'], // …or are the new editor's
    ['Report', 'DE', 'Report v2 DE'],
    ['', '', 'document v2'],
  ];
  for (const [base, initials, next] of cases) expect(nextVersionName(base, initials), `${base} ${initials}`).toBe(next);
});

test('initials from a name, as people write theirs', () => {
  expect(initialsOf('Dana Executive')).toBe('DE');
  expect(initialsOf('Ana María de la Cruz')).toBe('AMC'); // the capitalised words
  expect(initialsOf('Jean-Luc Picard')).toBe('JP');
  expect(initialsOf('')).toBe('');
});

const BODY = '<w:p><w:r><w:t xml:space="preserve">The lease runs for five years.</w:t></w:r></w:p>';
const review = (page) => page.getByRole('dialog', { name: 'Review your changes' });

async function edit(page, text) {
  await page.locator('.docx-page p').first().click();
  await page.keyboard.press('End');
  await page.keyboard.type(text);
}

/** Download through the review, as it stands; returns the file's name. */
async function downloadName(page) {
  const [dl] = await Promise.all([page.waitForEvent('download'), review(page).getByRole('button', { name: 'Download' }).click()]);
  return dl.suggestedFilename();
}

test('a changed document downloads as its next version when asked, and the name follows', async ({ page, browserName }) => {
  await page.addInitScript(() => localStorage.setItem('intact.author', 'Dana Executive'));
  await openInEditor(page, await buildDocx(BODY), 'Lease.docx');
  await edit(page, ' Rent is reviewed yearly.');
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  const names = review(page).getByRole('group', { name: 'File name' });
  await expect(names).toContainText('Downloads as Lease.docx'); // off until asked for
  await names.getByLabel(/Name it as a new version/).check();
  await expect(names).toContainText('Downloads as Lease v2.docx');
  expect(downloadNames('Lease v2.docx', browserName)).toContain(await downloadName(page));
  await expect(page.getByRole('textbox', { name: 'File name' })).toHaveValue('Lease v2');
  await expect(page.locator('.statusbar [role=status]')).toContainText('Saved Lease v2.docx');

  // The next edit, the next version; with initials, when asked for.
  await edit(page, ' Repairs are the landlord’s.');
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  await expect(names.getByLabel(/Name it as a new version/)).toBeChecked(); // remembered
  await expect(names).toContainText('Downloads as Lease v3.docx');
  await names.getByLabel('With my initials (DE)').check();
  await expect(names).toContainText('Downloads as Lease v3 DE.docx');
  expect(downloadNames('Lease v3 DE.docx', browserName)).toContain(await downloadName(page));

  // Nothing changed since: the file is the same, and so is its name; nothing is offered.
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  await expect(names).toHaveCount(0);
  await review(page).getByRole('button', { name: 'Keep editing' }).click();

  // Remembered for the next document, and applied without the review when it is skipped.
  await openInEditor(page, await buildDocx(BODY), 'Contract_v4_JD.docx');
  await edit(page, ' Signed.');
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  await review(page).getByLabel('Download without this review next time').check();
  expect(await downloadName(page)).toBe('Contract_v5_DE.docx');
  await edit(page, ' Again.');
  const [dl] = await Promise.all([page.waitForEvent('download'), page.locator('.app-header').getByRole('button', { name: 'Download' }).click()]);
  expect(dl.suggestedFilename()).toBe('Contract_v6_DE.docx');
});

test('a new document has no earlier version: its name is as typed', async ({ page }) => {
  await page.goto('./');
  await page.getByRole('button', { name: 'Blank document' }).click();
  await expect(page.locator('.docx-page')).toBeVisible();
  await page.locator('.docx-page p').first().click();
  await page.keyboard.type('Notes.');
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  await expect(review(page)).toBeVisible();
  await expect(review(page).getByRole('group', { name: 'File name' })).toHaveCount(0);
});
