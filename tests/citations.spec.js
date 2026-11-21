// Citations beyond one source in APA or MLA: several sources in one citation (Word's CITATION field
// with \m), Chicago (author-date), Harvard and IEEE (numbered in the order first cited) styles, the
// style a document's sources part records read back, and sources imported from a reference
// manager's BibTeX or RIS file, read on this device, kept in the document only once cited.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const P = (t) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
const ESSAY = P('Memory matters.') + P('We measured it.') + P('');
const BIB_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/bibliography';
const sourcesPart = (style, selected) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<b:Sources SelectedStyle="${selected}" StyleName="${style}" Version="6" xmlns:b="${BIB_NS}" xmlns="${BIB_NS}">`
  + '<b:Source><b:Tag>Doe19</b:Tag><b:SourceType>JournalArticle</b:SourceType><b:Author><b:Author><b:NameList><b:Person><b:Last>Doe</b:Last><b:First>Ann</b:First></b:Person></b:NameList></b:Author></b:Author>'
  + '<b:Title>On Memory</b:Title><b:JournalName>Mind</b:JournalName><b:Year>2019</b:Year><b:Volume>12</b:Volume><b:Issue>3</b:Issue><b:Pages>45-67</b:Pages></b:Source>'
  + '<b:Source><b:Tag>Smi20</b:Tag><b:SourceType>Book</b:SourceType><b:Author><b:Author><b:NameList><b:Person><b:Last>Smith</b:Last><b:First>Jane</b:First></b:Person><b:Person><b:Last>Lee</b:Last><b:First>Kim</b:First></b:Person></b:NameList></b:Author></b:Author>'
  + '<b:Title>Remembering</b:Title><b:Year>2020</b:Year><b:City>Boston</b:City><b:Publisher>North Press</b:Publisher></b:Source></b:Sources>';
const APA_PART = sourcesPart('APA', '\\APASixthEditionOfficeOnline.xsl');
const para = (page, text) => page.locator('.docx-page > p', { hasText: text });

async function insertMenu(page, item) {
  await page.getByRole('button', { name: /Insert a check box or date/ }).click();
  await page.getByRole('menuitem', { name: item }).click();
}
async function caretAt(page, text) {
  await para(page, text).first().click();
  await page.keyboard.press('End');
}
const dialog = (page) => page.getByRole('dialog', { name: 'Citation' });

test('several sources in one citation, in Chicago author-date: \\m in the field, and the bibliography in Chicago', async ({ page }) => {
  await openInEditor(page, await buildDocx(ESSAY, { parts: { 'customXml/item1.xml': APA_PART } }));
  await caretAt(page, 'Memory matters.');
  await insertMenu(page, /Citation/);
  const cite = dialog(page);
  await cite.getByLabel('Citation style').selectOption('Chicago');
  await expect(cite.getByLabel(/Smith, Lee \(2020\)/)).toBeChecked(); // the last source, suggested
  await cite.getByLabel(/Doe \(2019\)/).check();
  await cite.getByLabel('Page cited').fill('12');
  await expect(cite.locator('.preview')).toHaveText('(Smith and Lee 2020, 12; Doe 2019)');
  await cite.getByRole('button', { name: 'Insert' }).click();
  await expect(para(page, 'Memory')).toHaveText('Memory matters.(Smith and Lee 2020, 12; Doe 2019)');

  await page.locator('.docx-page > p').last().click();
  await insertMenu(page, /Bibliography/);
  await expect(page.locator('.docx-page > h1', { hasText: 'Bibliography' })).toHaveCount(1);
  await expect(para(page, 'On Memory')).toHaveText('Doe, Ann. 2019. “On Memory.” Mind 12 (3): 45–67.');
  await expect(para(page, 'Remembering')).toHaveText('Smith, Jane, and Kim Lee. 2020. Remembering. Boston: North Press.');

  const { zip, xml } = await download(page);
  expect(xml).toContain('<w:instrText xml:space="preserve"> CITATION Smi20 \\p 12 \\l 1033 \\m Doe19 </w:instrText>');
  const sources = await zip.file('customXml/item1.xml').async('string');
  expect(sources).toContain('SelectedStyle="\\CHICAGO.XSL" StyleName="Chicago" Version="16"');
});

test('IEEE numbers sources in the order first cited, in citations and the bibliography', async ({ page }) => {
  await openInEditor(page, await buildDocx(ESSAY, { parts: { 'customXml/item1.xml': APA_PART } }));
  // Smith cited in the second paragraph first, then Doe in the first: Doe becomes [1].
  await caretAt(page, 'We measured it.');
  await insertMenu(page, /Citation/);
  await dialog(page).getByLabel('Citation style').selectOption('IEEE');
  await expect(dialog(page).locator('.preview')).toHaveText('[1]');
  await dialog(page).getByRole('button', { name: 'Insert' }).click();
  await expect(para(page, 'We measured')).toHaveText('We measured it.[1]');
  await caretAt(page, 'Memory matters.');
  await insertMenu(page, /Citation/);
  await dialog(page).getByLabel(/Smith, Lee/).uncheck();
  await dialog(page).getByLabel(/Doe \(2019\)/).check();
  await dialog(page).getByLabel('Page cited').fill('50');
  await dialog(page).getByRole('button', { name: 'Insert' }).click();
  await expect(para(page, 'Memory')).toHaveText('Memory matters.[1, p. 50]');
  await expect(para(page, 'We measured')).toHaveText('We measured it.[2]');

  await page.locator('.docx-page > p').last().click();
  await insertMenu(page, /Bibliography/);
  const entries = page.locator('.docx-page > p', { hasText: /^\[\d\]/ });
  await expect(entries).toHaveCount(2);
  await expect(entries.nth(0)).toContainText('A. Doe, “On Memory,” Mind, vol. 12, no. 3, pp. 45–67, 2019.');
  await expect(entries.nth(1)).toContainText('J. Smith and K. Lee, Remembering. Boston: North Press, 2020.');
  const { zip, xml } = await download(page);
  expect(xml).toContain('<w:t xml:space="preserve">[1]</w:t><w:tab/><w:t xml:space="preserve">A. Doe, “On Memory,” </w:t>'); // the number, a tab, the entry
  expect(await zip.file('customXml/item1.xml').async('string')).toContain('StyleName="IEEE" Version="2006"');
});

test('a document whose sources Word keeps in Harvard: Update (F9) formats its citations in Harvard', async ({ page }) => {
  const cited = '<w:p><w:r><w:t xml:space="preserve">As shown </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> CITATION Doe19 \\l 1033 \\m Smi20 </w:instrText></w:r>'
    + '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>(Doe)</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>';
  await openInEditor(page, await buildDocx(cited + P('End.'), { parts: { 'customXml/item1.xml': sourcesPart('Harvard - Anglia', '\\HarvardAnglia2008OfficeOnline.xsl') } }));
  await page.locator('.docx-page > p').last().click();
  await page.keyboard.press('F9');
  await expect(para(page, 'As shown')).toHaveText('As shown (Doe, 2019; Smith and Lee, 2020)');
  await insertMenu(page, /Citation/);
  await expect(dialog(page).getByLabel('Citation style')).toHaveValue('Harvard');
});

const BIBTEX = String.raw`@article{doe2019,
  author = {Doe, Ann and M{\"u}ller, J{\"o}rg},
  title = {{Memory} under load},
  journal = {Mind},
  year = {2019}, volume = {12}, number = {3}, pages = {45--67},
  doi = {10.1000/mind.12}
}
@book{who2020, author = {{World Health Organization}}, title = {Global Report}, publisher = {WHO Press}, address = {Geneva}, year = {2020}}
@online{nrel, author = {{NREL}}, title = {Efficiency Chart}, url = {https://www.nrel.gov/chart}, urldate = {2024-03-05}, year = {2024}}`;

test('sources imported from a BibTeX file are offered; only those cited go into the document', async ({ page }) => {
  await openInEditor(page, await buildDocx(ESSAY));
  await caretAt(page, 'Memory matters.');
  await insertMenu(page, /Citation/);
  const cite = dialog(page);
  await cite.locator('input[type=file]').setInputFiles({ name: 'library.bib', mimeType: 'application/x-bibtex', buffer: Buffer.from(BIBTEX) });
  await expect(cite.getByRole('status')).toHaveText('3 sources imported from library.bib: tick those to cite. Only sources cited go into the document.');
  await cite.getByLabel('A new source…').uncheck();
  await cite.getByLabel(/Doe, Müller \(2019\) – Memory under load/).check();
  await cite.getByLabel(/World Health Organization \(2020\)/).check();
  await expect(cite.locator('.preview')).toHaveText('(Doe & Müller, 2019; World Health Organization, 2020)');
  await cite.getByRole('button', { name: 'Insert' }).click();
  await expect(para(page, 'Memory')).toHaveText('Memory matters.(Doe & Müller, 2019; World Health Organization, 2020)');

  // The one not cited is still offered; the two cited are now the document's.
  await caretAt(page, 'We measured it.');
  await insertMenu(page, /Citation/);
  await expect(cite.getByRole('group', { name: 'Imported from a file' }).locator('label')).toHaveText([/NREL \(2024\) – Efficiency Chart/]);
  await expect(cite.getByRole('group', { name: 'Cite sources in this document' }).locator('label')).toHaveCount(2);
  // Importing the same file again adds nothing new.
  await cite.locator('input[type=file]').setInputFiles({ name: 'library.bib', mimeType: 'application/x-bibtex', buffer: Buffer.from(BIBTEX) });
  await expect(cite.getByRole('status')).toHaveText('0 sources imported from library.bib (3 already here): tick those to cite. Only sources cited go into the document.');
  await cite.getByRole('button', { name: 'Close' }).click();

  await page.locator('.docx-page > p').last().click();
  await insertMenu(page, /Bibliography/);
  await expect(para(page, 'Memory under load')).toHaveText('Doe, A., & Müller, J. (2019). Memory under load. Mind, 12(3), 45–67. https://doi.org/10.1000/mind.12');
  await expect(para(page, 'Global Report')).toHaveText('World Health Organization. (2020). Global Report. WHO Press.');
  const { zip } = await download(page);
  const sources = await zip.file('customXml/item1.xml').async('string');
  expect(sources).toContain('<b:Person><b:Last>Müller</b:Last><b:First>Jörg</b:First></b:Person>');
  expect(sources).toContain('<b:Corporate>World Health Organization</b:Corporate>');
  expect(sources).toContain('<b:City>Geneva</b:City><b:Publisher>WHO Press</b:Publisher>');
  expect(sources).not.toContain('Efficiency Chart');
});

test('a RIS file from EndNote is read too; a file that is neither says so', async ({ page }) => {
  await openInEditor(page, await buildDocx(ESSAY));
  await caretAt(page, 'Memory matters.');
  await insertMenu(page, /Citation/);
  const cite = dialog(page);
  await cite.locator('input[type=file]').setInputFiles({ name: 'notes.txt', mimeType: 'text/plain', buffer: Buffer.from('Just some notes.') });
  await expect(cite.getByRole('alert')).toHaveText('No sources found in notes.txt: choose a BibTeX (.bib) or RIS (.ris) file');
  const ris = 'TY  - JOUR\r\nAU  - Roe, Bo\r\nTI  - Sleep and recall\r\nJO  - Brain\r\nPY  - 2018///\r\nVL  - 4\r\nSP  - 10\r\nEP  - 19\r\nER  - \r\n';
  await cite.locator('input[type=file]').setInputFiles({ name: 'export.ris', mimeType: 'application/x-research-info-systems', buffer: Buffer.from(ris) });
  await cite.getByLabel('A new source…').uncheck();
  await cite.getByLabel(/Roe \(2018\) – Sleep and recall/).check();
  await cite.getByLabel('Citation style').selectOption('MLA');
  await cite.getByLabel('Page cited').fill('12');
  await cite.getByRole('button', { name: 'Insert' }).click();
  await expect(para(page, 'Memory')).toHaveText('Memory matters.(Roe 12)');
  const { zip } = await download(page);
  const sources = await zip.file('customXml/item1.xml').async('string');
  expect(sources).toContain('<b:Title>Sleep and recall</b:Title><b:Year>2018</b:Year><b:JournalName>Brain</b:JournalName><b:Volume>4</b:Volume><b:Pages>10-19</b:Pages>');
});

test('nothing ticked is refused', async ({ page }) => {
  await openInEditor(page, await buildDocx(ESSAY));
  await caretAt(page, 'Memory matters.');
  await insertMenu(page, /Citation/);
  await dialog(page).getByLabel('A new source…').uncheck();
  await dialog(page).getByRole('button', { name: 'Insert' }).click();
  await expect(dialog(page).getByRole('alert')).toHaveText('Tick a source to cite');
});
