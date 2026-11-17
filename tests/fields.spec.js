// Fields Word works out (tables of contents, cross-references, caption numbers) are kept as they are
// and never updated here. After editing, the review says they may be out of date and offers to have
// Word update them on opening (w:updateFields in settings.xml); when citations were added or removed
// and a bibliography lists them, it says the bibliography needs refreshing.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, W_NS } from './helpers.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const field = (instr, shown) => `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t xml:space="preserve">${shown}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
const p = (inner) => `<w:p>${inner}</w:p>`;
const text = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const ZOTERO = 'ADDIN ZOTERO_ITEM CSL_CITATION {&quot;citationID&quot;:&quot;a1&quot;}';

const THESIS = p(field('TOC \\o "1-3" \\h \\z \\u', 'Introduction 1'))
  + p(`${text('As shown in Figure ')}${field('REF _Ref1 \\h', '1')}${text(', results hold.')}`)
  + p(`${text('Figure ')}<w:fldSimple w:instr=" SEQ Figure \\* ARABIC "><w:r><w:t>1</w:t></w:r></w:fldSimple>`)
  + p(`${text('Prior work ')}${field(ZOTERO, '(Smith, 2020)')}`)
  + p(`${text('Later work ')}${field(ZOTERO, '(Jones, 2021)')}`)
  + p(field('ADDIN ZOTERO_BIBL {&quot;uncited&quot;:[]} CSL_BIBLIOGRAPHY', 'Jones, A. (2021). Smith, J. (2020).'));

const SETTINGS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:settings ${W_NS}><w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat><w:rsids><w:rsidRoot w:val="00000001"/></w:rsids></w:settings>`;
const WITH_SETTINGS = (settings) => ({
  parts: { 'word/settings.xml': settings },
  rels: `<Relationship Id="rIdS" Type="${REL}/settings" Target="settings.xml"/>`,
});

const review = (page) => page.getByRole('dialog', { name: 'Review your changes' });
const notes = (page) => review(page).locator('.notes');

async function openReview(page) {
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  await expect(review(page)).toBeVisible();
}

async function edit(page) {
  await page.getByText(', results hold.').click();
  await page.keyboard.press('End');
  await page.keyboard.type(' Clearly.');
}

test('after editing, the review says which fields Word must update, and how', async ({ page }) => {
  await openInEditor(page, await buildDocx(THESIS, WITH_SETTINGS(SETTINGS)));
  await openReview(page);
  // Nothing edited: nothing to update.
  await expect(notes(page)).toHaveCount(0);
  await review(page).getByRole('button', { name: 'Keep editing' }).click();
  await edit(page);
  await openReview(page);
  await expect(notes(page).locator('li.warn').first()).toHaveText(/^This document has a table of contents, cross-references and figure and table numbers: they may be out of date after your changes. Update them here with Update fields \(F9\) .* select all \(Ctrl\+A\) and press F9/);
  // Citations weren't added or removed: the bibliography is still right.
  await expect(notes(page)).not.toContainText('bibliography');
  // Not asked for: settings.xml stays as it was.
  const [dl] = await Promise.all([page.waitForEvent('download'), review(page).getByRole('button', { name: 'Download' }).click()]);
  const JSZip = (await import('jszip')).default;
  const zip = await JSZip.loadAsync(await (await import('node:fs')).promises.readFile(await dl.path()));
  expect(await zip.file('word/settings.xml').async('string')).toBe(SETTINGS);
});

test('asking Word to update fields on opening writes w:updateFields in schema order', async ({ page }) => {
  await openInEditor(page, await buildDocx(THESIS, WITH_SETTINGS(SETTINGS)));
  await edit(page);
  await openReview(page);
  await review(page).getByRole('checkbox', { name: 'Ask Word to update them when the file is opened' }).check();
  await review(page).getByRole('button', { name: 'Keep editing' }).click();
  const { zip } = await download(page); // the choice is kept for the next download
  expect(await zip.file('word/settings.xml').async('string')).toContain('<w:defaultTabStop w:val="720"/><w:updateFields w:val="true"/><w:compat>');
});

test('a file without settings gets them when Word is asked to update fields', async ({ page }) => {
  await openInEditor(page, await buildDocx(THESIS));
  await edit(page);
  await openReview(page);
  const box = review(page).getByRole('checkbox', { name: 'Ask Word to update them when the file is opened' });
  await box.check();
  await box.uncheck();
  await box.check();
  const [dl] = await Promise.all([page.waitForEvent('download'), review(page).getByRole('button', { name: 'Download' }).click()]);
  const JSZip = (await import('jszip')).default;
  const zip = await JSZip.loadAsync(await (await import('node:fs')).promises.readFile(await dl.path()));
  expect(await zip.file('word/settings.xml').async('string')).toContain('<w:updateFields w:val="true"/>');
  expect(await zip.file('word/_rels/document.xml.rels').async('string')).toContain('settings.xml');
});

test('a file that already asks Word to update fields says so, with nothing to choose', async ({ page }) => {
  const asking = SETTINGS.replace('<w:compat>', '<w:updateFields w:val="true"/><w:compat>');
  await openInEditor(page, await buildDocx(THESIS, WITH_SETTINGS(asking)));
  await edit(page);
  await openReview(page);
  await expect(notes(page).locator('li.info', { hasText: 'table of contents' })).toHaveText('This document has a table of contents, cross-references and figure and table numbers. The file asks Word to update them when it is opened.');
  await expect(review(page).getByRole('checkbox', { name: /update them/ })).toHaveCount(0);
});

test('removing a citation says the bibliography needs refreshing', async ({ page }) => {
  await openInEditor(page, await buildDocx(THESIS));
  await page.evaluate(() => {
    const para = [...document.querySelectorAll('.docx-page p')].find((el) => el.textContent.startsWith('Later work'));
    const range = document.createRange();
    range.selectNodeContents(para);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
  });
  await page.keyboard.press('Delete');
  await expect(page.locator('.docx-page p', { hasText: 'Later work' })).toHaveCount(0);
  await openReview(page);
  await expect(notes(page)).toContainText('Citations were added or removed, but the bibliography doesn’t change until it is refreshed in Word');
});

test('a document without such fields gets no note about them', async ({ page }) => {
  await openInEditor(page, await buildDocx(p(text('Just text.'))));
  await page.getByText('Just text.').click();
  await page.keyboard.press('End');
  await page.keyboard.type(' More.');
  await openReview(page);
  await expect(notes(page)).not.toContainText('Word updates');
  await expect(review(page).getByRole('checkbox', { name: /update them/ })).toHaveCount(0);
});

test('citations and bibliographies say which reference manager they come from, and warn when typed in', async ({ page }) => {
  const endnote = p(`${text('Earlier ')}${field('ADDIN EN.CITE &lt;EndNote&gt;&lt;/EndNote&gt;', '[3]')}`);
  const word = p(`${text('Also ')}${field('CITATION Smi20 \\l 1033', '(Smith 2020)')}`);
  await openInEditor(page, await buildDocx(THESIS + endnote + word));
  const zotero = page.locator('.docx-page .field', { hasText: '(Smith, 2020)' });
  await expect(zotero).toHaveAttribute('title', 'Zotero citation – Zotero updates it in Word; edits made here are lost when it does');
  await expect(page.locator('.docx-page .field', { hasText: 'Jones, A.' })).toHaveAttribute('title', /^Zotero bibliography – /);
  await expect(page.locator('.docx-page .field', { hasText: '[3]' })).toHaveAttribute('title', /^EndNote citation – EndNote updates it in Word/);
  await expect(page.locator('.docx-page .field', { hasText: '(Smith 2020)' })).toHaveAttribute('title', 'Citation – Word updates it from the document’s sources');
  await expect(page.locator('.docx-page .field', { hasText: 'Introduction 1' })).toHaveAttribute('title', 'Table of contents – Word updates it');
  // Typing in one says what will happen to the edit.
  await page.evaluate(() => {
    const t = [...document.querySelectorAll('.docx-page .field')].find((f) => f.textContent === '(Smith, 2020)').firstChild;
    const r = document.createRange();
    r.setStart(t.nodeType === 3 ? t : t.firstChild, 3);
    getSelection().removeAllRanges();
    getSelection().addRange(r);
  });
  await page.keyboard.type('x');
  await expect(page.locator('.statusbar [role=status]')).toHaveText('Zotero citation: Zotero replaces this text when it refreshes in Word – change it there instead');
});
