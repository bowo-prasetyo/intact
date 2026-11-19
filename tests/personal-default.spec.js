// Removing personal information (author, last saved by, company, manager) can be made the default,
// by choice, for files where nothing would be lost by it: ones you change that have no comments or
// tracked changes. Remembered in this browser; the review says when it was chosen by the default,
// and it can be kept in a file. Where your name is known, the review says when the author isn't you.
import JSZip from 'jszip';
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor } from './helpers.js';
import fs from 'node:fs';

const para = (text) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

/** `buffer` with core properties naming `creator` (a template's designer, say). */
async function withAuthor(buffer, creator = 'Template Designer') {
  const zip = await JSZip.loadAsync(buffer);
  zip.file('docProps/core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>${creator}</dc:creator><cp:lastModifiedBy>${creator}</cp:lastModifiedBy><cp:revision>3</cp:revision></cp:coreProperties>`);
  zip.file('_rels/.rels', (await zip.file('_rels/.rels').async('string')).replace('</Relationships>', '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>'));
  zip.file('[Content_Types].xml', (await zip.file('[Content_Types].xml').async('string')).replace('</Types>', '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>'));
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function edit(page) {
  await page.locator('.docx-page p').last().click();
  await page.keyboard.press('End');
  await page.keyboard.type(' edited');
}
const review = (page) => page.getByRole('dialog', { name: 'Review your changes' });
const item = (page) => review(page).locator('.release li', { hasText: 'Personal information' });
const remove = (page) => item(page).getByRole('checkbox', { name: /Remove it from the downloaded file/ });
const everyFile = (page) => item(page).getByRole('checkbox', { name: /Do this for every file I change/ });
async function openDownload(page) {
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  await expect(review(page)).toBeVisible();
}
async function savedCore(page) {
  const [dl] = await Promise.all([page.waitForEvent('download'), review(page).getByRole('button', { name: 'Download' }).click()]);
  return (await JSZip.loadAsync(fs.readFileSync(await dl.path()))).file('docProps/core.xml').async('string');
}

test('it isn\'t the default until you make it so: the name stays, and the review offers it', async ({ page }) => {
  await openInEditor(page, await withAuthor(await buildDocx(para('Alex Morgan'))));
  await edit(page);
  await openDownload(page);
  await expect(item(page)).toContainText('Personal information: Author “Template Designer”, Last saved by “Template Designer”.');
  await expect(remove(page)).not.toBeChecked();
  await expect(everyFile(page)).not.toBeChecked();
  expect(await savedCore(page)).toContain('Template Designer');
});

test('where your name is known, the review says the author isn\'t you', async ({ page }) => {
  await page.addInitScript(() => localStorage.setItem('intact.author', 'Alex Morgan'));
  await openInEditor(page, await withAuthor(await buildDocx(para('Alex Morgan'))));
  await edit(page);
  await openDownload(page);
  await expect(item(page)).toContainText('That isn’t you (Alex Morgan): a file made from someone else’s often carries their name.');
});

test('made the default, it removes the name from this file and from the next one you change', async ({ page }) => {
  await openInEditor(page, await withAuthor(await buildDocx(para('Alex Morgan'))));
  await edit(page);
  await openDownload(page);
  await everyFile(page).check();
  await expect(remove(page)).toBeChecked();
  await expect(everyFile(page)).toBeChecked();
  expect(await savedCore(page)).not.toContain('Template Designer');

  await page.getByRole('button', { name: 'Close' }).click();
  await page.setInputFiles('input[type=file]', { name: 'Cover_letter.docx', mimeType: 'application/octet-stream', buffer: await withAuthor(await buildDocx(para('Dear hiring manager')), 'Someone Else') });
  await edit(page);
  await openDownload(page);
  // Chosen by the default: the section opens, says what goes and why.
  await expect(item(page)).toContainText('Personal information is removed from the downloaded file (Author “Someone Else”, Last saved by “Someone Else”), as you chose for files you change that have no comments or tracked changes. Untick to keep it in this one.');
  await expect(remove(page)).toBeChecked();
  // Kept in this file, the default stays.
  await remove(page).uncheck();
  await expect(remove(page)).not.toBeChecked();
  await expect(everyFile(page)).toBeChecked();
  expect(await savedCore(page)).toContain('Someone Else');
});

test.describe('the default leaves alone', () => {
  test.beforeEach(async ({ page }) => {
    await page.addInitScript(() => localStorage.setItem('intact.removePersonal', '1'));
  });

  test('a file downloaded as it was opened', async ({ page }) => {
    await openInEditor(page, await withAuthor(await buildDocx(para('Text'))));
    await page.getByRole('button', { name: 'Review', exact: true }).click();
    await expect(remove(page)).not.toBeChecked();
  });

  test('a file with tracked changes, whose authors\' names it would change', async ({ page }) => {
    const tracked = '<w:p><w:ins w:id="1" w:author="Lawyer" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Clause</w:t></w:r></w:ins></w:p>' + para('Text');
    await openInEditor(page, await withAuthor(await buildDocx(tracked)));
    await edit(page);
    await openDownload(page);
    await expect(remove(page)).not.toBeChecked();
    await expect(everyFile(page)).toBeChecked(); // shown, as it is on
  });

  test('a placeholder name someone\'s removal already left', async ({ page }) => {
    await openInEditor(page, await withAuthor(await buildDocx(para('Text')), 'Author'));
    await edit(page);
    await openDownload(page);
    expect(await savedCore(page)).toContain('<dc:creator>Author</dc:creator>');
  });
});

test('the default leaves the properties alone when recording the edit is switched off', async ({ page }) => {
  await page.addInitScript(() => { localStorage.setItem('intact.removePersonal', '1'); localStorage.setItem('intact.updateProperties', '0'); });
  await openInEditor(page, await withAuthor(await buildDocx(para('Text'))));
  await edit(page);
  await openDownload(page);
  await expect(remove(page)).not.toBeChecked();
  expect(await savedCore(page)).toContain('Template Designer');
});
