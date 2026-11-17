// Saving an edit records it in the document's core properties, as Word does: modified now, last
// modified by whoever edited (if they gave their name), the revision one up. Optional, and only
// when something changed.
import fs from 'node:fs';
import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import { buildDocx, openInEditor } from './helpers.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CORE = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>Officer Alpha</dc:creator><cp:lastModifiedBy>Officer Bravo</cp:lastModifiedBy><cp:revision>4</cp:revision><dcterms:created xsi:type="dcterms:W3CDTF">2025-12-01T00:00:00Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2026-01-05T00:00:00Z</dcterms:modified></cp:coreProperties>';

async function withCore(core = CORE) {
  const zip = await JSZip.loadAsync(await buildDocx('<w:p><w:r><w:t>The agreed text.</w:t></w:r></w:p>'));
  let ct = await zip.file('[Content_Types].xml').async('string');
  zip.file('[Content_Types].xml', ct.replace('</Types>', '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>'));
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PKG}"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="${PKG}/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`);
  zip.file('docProps/core.xml', core);
  return zip.generateAsync({ type: 'nodebuffer' });
}

const review = (page) => page.getByRole('dialog', { name: 'Review your changes' });

/** Your name, as given once for comments and Track Changes (remembered in this browser). */
const named = (page, name) => page.addInitScript((n) => localStorage.setItem('intact.author', n), name);

async function edit(page) {
  await page.locator('.docx-page p').first().click();
  await page.keyboard.press('End');
  await page.keyboard.type(' Changed');
}

async function saveCore(page) {
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  const [dl] = await Promise.all([page.waitForEvent('download'), review(page).getByRole('button', { name: 'Download', exact: true }).click()]);
  return (await JSZip.loadAsync(fs.readFileSync(await dl.path()))).file('docProps/core.xml').async('string');
}

test('an edit is recorded: modified now, by you, the next revision', async ({ page }) => {
  await named(page, 'Rita Reviewer');
  await openInEditor(page, await withCore());
  await edit(page);
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  await expect(review(page)).toContainText('The document’s properties record this edit: modified now, last saved by Rita Reviewer, revision 5.');
  await review(page).getByRole('button', { name: 'Keep editing' }).click();
  const before = Date.now();
  const core = await saveCore(page);
  expect(core).toContain('<cp:lastModifiedBy>Rita Reviewer</cp:lastModifiedBy>');
  expect(core).toContain('<cp:revision>5</cp:revision>');
  expect(core).toContain('<dc:creator>Officer Alpha</dc:creator>'); // who created it stays
  expect(core).toContain('<dcterms:created xsi:type="dcterms:W3CDTF">2025-12-01T00:00:00Z</dcterms:created>');
  const modified = Date.parse(/<dcterms:modified xsi:type="dcterms:W3CDTF">([^<]+)</.exec(core)[1]);
  expect(Math.abs(modified - before)).toBeLessThan(60_000);

  // Saved again in the same session: still one revision on from the file as opened.
  expect(await saveCore(page)).toContain('<cp:revision>5</cp:revision>');
});

test('without your name, the last to save it is kept, and the review says so', async ({ page }) => {
  await openInEditor(page, await withCore());
  await edit(page);
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  await expect(review(page)).toContainText('last saved by kept as “Officer Bravo” (give your name with a comment or Track changes to record it), revision 5.');
  await review(page).getByRole('button', { name: 'Keep editing' }).click();
  const core = await saveCore(page);
  expect(core).toContain('<cp:lastModifiedBy>Officer Bravo</cp:lastModifiedBy>');
  expect(core).toContain('<cp:revision>5</cp:revision>');
});

test('nothing changed, nothing recorded: the properties come back byte for byte', async ({ page }) => {
  await named(page, 'Rita Reviewer');
  await openInEditor(page, await withCore());
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  await expect(review(page)).not.toContainText('properties record');
  await review(page).getByRole('button', { name: 'Keep editing' }).click();
  expect(await saveCore(page)).toBe(CORE);
});

test('recording can be switched off, and the choice is remembered', async ({ page }) => {
  await named(page, 'Rita Reviewer');
  const docx = await withCore();
  await openInEditor(page, docx);
  await edit(page);
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  await review(page).getByRole('checkbox', { name: 'Record this edit in the document’s properties' }).uncheck();
  await expect(review(page)).toContainText('The document’s properties stay as they were: they don’t record this edit.');
  await review(page).getByRole('button', { name: 'Keep editing' }).click();
  expect(await saveCore(page)).toBe(CORE);

  await page.getByRole('button', { name: 'Close' }).click();
  await openInEditor(page, docx);
  await edit(page);
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  await expect(review(page).getByRole('checkbox', { name: 'Record this edit in the document’s properties' })).not.toBeChecked();
});

test('with personal information removed, no name is recorded', async ({ page }) => {
  await named(page, 'Rita Reviewer');
  await openInEditor(page, await withCore());
  await edit(page);
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  await review(page).getByRole('checkbox', { name: /Remove it from the downloaded file/ }).check();
  await expect(review(page)).toContainText('The document’s properties record this edit: modified now, revision 5.');
  await review(page).getByRole('button', { name: 'Keep editing' }).click();
  const core = await saveCore(page);
  expect(core).not.toContain('lastModifiedBy');
  expect(core).not.toContain('Officer');
  expect(core).toContain('<cp:revision>5</cp:revision>');
});
