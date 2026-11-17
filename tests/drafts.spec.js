// Drafts: while a document has changes not downloaded, a copy is kept in this browser, so a reload,
// a crash or a closed tab doesn't lose them; the start page offers it back. Nothing leaves the
// device, and the copy goes when the document is downloaded, closed, or discarded.
import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import { buildDocx, openInEditor, download } from './helpers.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CORE = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>Ann</dc:creator><cp:revision>4</cp:revision></cp:coreProperties>';

const BODY = '<w:p><w:r><w:t xml:space="preserve">The fee is due.</w:t></w:r></w:p>';
const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const kept = (page) => page.locator('.statusbar .draft-status');
const drafts = (page) => page.getByRole('region', { name: 'Unsaved changes on this device' });

async function typeAtEnd(page, text) {
  await page.locator('.docx-page p').first().click();
  await page.keyboard.press('End');
  await page.keyboard.type(text);
}

/** How many drafts IndexedDB holds. */
const storedDrafts = (page) => page.evaluate(() => new Promise((resolve) => {
  const req = indexedDB.open('intact');
  req.onsuccess = () => {
    const db = req.result;
    if (!db.objectStoreNames.contains('drafts')) { resolve(0); return; }
    const all = db.transaction('drafts').objectStore('drafts').getAll();
    all.onsuccess = () => resolve(all.result.length);
  };
}));

test('changes survive a reload: the start page offers them back, and they download as edited', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY), 'fees.docx');
  await expect(kept(page)).toHaveCount(0); // nothing changed, nothing kept
  await typeAtEnd(page, ' Pay by Friday.');
  await expect(kept(page)).toHaveText('Changes kept on this device');
  await expect(kept(page)).toHaveAttribute('title', /on this device only/);
  await page.reload(); // the editor's beforeunload prompt is accepted by openInEditor's handler

  await expect(drafts(page)).toContainText('fees.docx');
  await drafts(page).getByRole('button', { name: 'Restore fees.docx' }).click();
  await expect(page.locator('.docx-page')).toContainText('The fee is due. Pay by Friday.');
  await expect(page.locator('.app-header .dirty')).toBeVisible(); // not downloaded yet
  await expect(page.locator('.statusbar [role=status]')).toHaveText(/^Restored your unsaved changes from /);
  await expect(page.locator('.app-header .filename')).toHaveValue('fees');

  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  const review = page.getByRole('dialog', { name: 'Review your changes' });
  await expect(review).toContainText('Restored from the copy kept on this device');
  await expect(review).toContainText('No changes since the restored copy.');
  await review.getByRole('button', { name: 'Keep editing' }).click();
  const { xml } = await download(page);
  expect(xml).toContain('The fee is due. Pay by Friday.');
  // Downloaded: the file is the copy now.
  await expect(kept(page)).toHaveCount(0);
  await expect.poll(() => storedDrafts(page)).toBe(0);
});

test('a restored draft records the edit in the properties once, as the first download would have', async ({ page }) => {
  // The draft holds the edits only: the properties' record of the edit is made by the download.
  const zip = await JSZip.loadAsync(await buildDocx(BODY));
  const ct = await zip.file('[Content_Types].xml').async('string');
  zip.file('[Content_Types].xml', ct.replace('</Types>', '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>'));
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PKG}"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="${PKG}/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`);
  zip.file('docProps/core.xml', CORE);
  await openInEditor(page, await zip.generateAsync({ type: 'nodebuffer' }), 'fees.docx');
  await typeAtEnd(page, ' Pay by Friday.');
  await expect(kept(page)).toBeVisible();
  await page.reload();
  await drafts(page).getByRole('button', { name: 'Restore fees.docx' }).click();
  const { zip: saved } = await download(page); // nothing edited since restoring: it still counts as changed
  const props = await saved.file('docProps/core.xml').async('string');
  expect(props).toContain('<cp:revision>5</cp:revision>');
  expect(props).toContain('<dc:creator>Ann</dc:creator>');
});

test('Discard on the start page deletes the draft; closing the document on purpose does too', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY), 'fees.docx');
  await typeAtEnd(page, ' Draft one.');
  await expect(kept(page)).toBeVisible();
  await page.reload();
  await drafts(page).getByRole('button', { name: 'Discard the changes to fees.docx' }).click();
  await expect(drafts(page)).toHaveCount(0);
  expect(await storedDrafts(page)).toBe(0);

  await page.setInputFiles('input[type=file]', { name: 'fees.docx', mimeType: DOCX, buffer: await buildDocx(BODY) });
  await typeAtEnd(page, ' Draft two.');
  await expect(kept(page)).toBeVisible();
  await page.locator('.app-header').getByRole('button', { name: 'Close' }).click(); // "Leave without downloading?" accepted
  await expect(page.locator('.home')).toBeVisible();
  await expect(drafts(page)).toHaveCount(0);
  await expect.poll(() => storedDrafts(page)).toBe(0);
});

test('after restoring, Track Changes is still on, and the changes tracked before are there to accept or reject', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY), 'fees.docx');
  await page.getByRole('button', { name: 'Track changes' }).click();
  const dialog = page.getByRole('dialog', { name: 'Your name on tracked changes' });
  await dialog.getByRole('textbox').fill('Rita');
  await dialog.getByRole('button', { name: 'Start tracking' }).click();
  await typeAtEnd(page, ' Pay by Friday.');
  await expect(kept(page)).toBeVisible();
  await page.reload();
  await drafts(page).getByRole('button', { name: 'Restore fees.docx' }).click();
  await expect(page.getByRole('button', { name: 'Track changes' })).toHaveAttribute('aria-pressed', 'true');
  await expect(page.locator('.changes-group .badge')).toHaveText('1');
  // Typed away from the earlier change, which reopens as a tracked change kept from the file (in
  // Firefox, End puts the caret at the edge of that read-only text, where typing is refused).
  await page.locator('.docx-page p').first().click();
  await page.keyboard.press('Home');
  await page.keyboard.type('Note: ');
  await expect(page.locator('.docx-page .tc-ins')).toHaveText('Note: ');
  const { xml } = await download(page);
  expect(xml.match(/<w:ins [^>]*w:author="Rita"/g)).toHaveLength(2);
});
