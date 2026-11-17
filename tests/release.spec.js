// The sensitivity label is shown, never changed; "Before you release" in the review lists what a
// file carries beyond the page (comments, tracked changes, hidden text, personal information,
// properties, custom XML, embedded files) and removes what can be removed here.
import fs from 'node:fs';
import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import { buildDocx, openInEditor, W_NS } from './helpers.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const LABEL = '3a7c1d2e-0000-4000-8000-000000000001';

const CORE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Budget memo</dc:title><dc:creator>Officer Alpha</dc:creator><cp:lastModifiedBy>Officer Bravo</cp:lastModifiedBy><cp:revision>4</cp:revision><dcterms:modified xsi:type="dcterms:W3CDTF">2026-01-05T00:00:00Z</dcterms:modified></cp:coreProperties>`;
const APP = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Template>Normal.dotm</Template><Company>Ministry of Examples</Company><Manager>Director Charlie</Manager></Properties>';
const prop = (pid, name, value) => `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="${pid}" name="${name}"><vt:lpwstr>${value}</vt:lpwstr></property>`;
const custom = (...props) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">${props.join('')}</Properties>`;
const LABEL_PROPS = [
  prop(2, `MSIP_Label_${LABEL}_Enabled`, 'true'), prop(3, `MSIP_Label_${LABEL}_SetDate`, '2026-03-02T09:00:00Z'),
  prop(4, `MSIP_Label_${LABEL}_Method`, 'Privileged'), prop(5, `MSIP_Label_${LABEL}_Name`, 'OFFICIAL-SENSITIVE'),
  prop(6, `MSIP_Label_${LABEL}_ContentBits`, '3'),
];

const COMMENTS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments ${W_NS}><w:comment w:id="10" w:author="Ann Analyst" w:initials="AA" w:date="2026-01-02T00:00:00Z"><w:p><w:r><w:t>Check the figure.</w:t></w:r></w:p></w:comment></w:comments>`;

const BODY = '<w:p><w:r><w:t xml:space="preserve">Spending is </w:t></w:r><w:commentRangeStart w:id="10"/><w:r><w:t>up</w:t></w:r><w:commentRangeEnd w:id="10"/><w:r><w:commentReference w:id="10"/></w:r>'
  + '<w:r><w:t xml:space="preserve"> by </w:t></w:r><w:del w:id="11" w:author="Bob Budget" w:date="2026-01-03T00:00:00Z"><w:r><w:delText>12</w:delText></w:r></w:del><w:ins w:id="12" w:author="Bob Budget" w:date="2026-01-03T00:00:00Z"><w:r><w:t>9</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> per cent.</w:t></w:r></w:p>'
  + '<w:p><w:r><w:t xml:space="preserve">Published. </w:t></w:r><w:r><w:rPr><w:vanish/></w:rPr><w:t>Internal note: the minister disagrees.</w:t></w:r></w:p>';

/** A document with properties (and `extra` parts in the package beside them). */
async function withProperties(body, { customXml = custom(...LABEL_PROPS, prop(7, 'Case number', 'FOI-2026-041')), extra = {}, rels = '', parts = {} } = {}) {
  const zip = await JSZip.loadAsync(await buildDocx(body, { rels, parts }));
  let ct = await zip.file('[Content_Types].xml').async('string');
  ct = ct.replace('</Types>', '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
    + '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
    + '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/></Types>');
  zip.file('[Content_Types].xml', ct);
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PKG}">`
    + `<Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/>`
    + `<Relationship Id="rId2" Type="${PKG}/metadata/core-properties" Target="docProps/core.xml"/>`
    + `<Relationship Id="rId3" Type="${REL}/extended-properties" Target="docProps/app.xml"/>`
    + (customXml ? `<Relationship Id="rId4" Type="${REL}/custom-properties" Target="docProps/custom.xml"/>` : '')
    + '</Relationships>');
  zip.file('docProps/core.xml', CORE);
  zip.file('docProps/app.xml', APP);
  if (customXml) zip.file('docProps/custom.xml', customXml);
  for (const [path, content] of Object.entries(extra)) zip.file(path, content);
  return zip.generateAsync({ type: 'nodebuffer' });
}

const FULL = () => withProperties(BODY, {
  rels: `<Relationship Id="rIdC" Type="${REL}/comments" Target="comments.xml"/>`,
  parts: { 'word/comments.xml': COMMENTS },
  extra: { 'customXml/item1.xml': '<root><library>Cabinet papers</library></root>', 'word/embeddings/Microsoft_Excel_Worksheet.xlsx': 'PK' },
});

const chip = (page) => page.getByRole('note');
const review = (page) => page.getByRole('dialog', { name: 'Review your changes' });
const releaseSection = (page) => review(page).getByRole('region', { name: 'Before you release' });

async function openReview(page) {
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  await expect(review(page)).toBeVisible();
}

async function downloadFromReview(page) {
  const [dl] = await Promise.all([page.waitForEvent('download'), review(page).getByRole('button', { name: 'Download', exact: true }).click()]);
  return JSZip.loadAsync(fs.readFileSync(await dl.path()));
}

// ------------------------------------------------------------------ sensitivity label

test('the sensitivity label is shown beside the name and kept as it is', async ({ page }) => {
  const docx = await FULL();
  await openInEditor(page, docx);
  await expect(chip(page)).toHaveText('OFFICIAL-SENSITIVE');
  const title = await chip(page).getAttribute('title');
  expect(title).toContain('Sensitivity label: OFFICIAL-SENSITIVE – chosen by a person, on');
  expect(title).toContain('In Office it marks the document with a header and footer.');
  expect(title).toContain('Intact keeps the label as it is');
  await openReview(page);
  await expect(review(page)).toContainText('Sensitivity label “OFFICIAL-SENSITIVE” stays on the file, unchanged.');
  const zip = await downloadFromReview(page);
  const original = await JSZip.loadAsync(docx);
  expect(await zip.file('docProps/custom.xml').async('string')).toBe(await original.file('docProps/custom.xml').async('string'));
});

test('a label recorded only in LabelInfo.xml is shown without a name', async ({ page }) => {
  await openInEditor(page, await withProperties('<w:p><w:r><w:t>Text.</w:t></w:r></w:p>', {
    customXml: null,
    extra: { 'docMetadata/LabelInfo.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><clbl:labelList xmlns:clbl="http://schemas.microsoft.com/office/2020/mipLabelMetadata"><clbl:label id="{${LABEL}}" enabled="1" method="Standard" siteId="{00000000-0000-0000-0000-000000000002}" removed="0"/></clbl:labelList>` },
  }));
  await expect(chip(page)).toHaveText('Sensitivity label');
  expect(await chip(page).getAttribute('title')).toContain('set (its name isn’t stored in the file) – applied by default or automatically');
});

test('a removed or disabled label isn’t shown; another tool’s classification is', async ({ page }) => {
  await openInEditor(page, await withProperties('<w:p><w:r><w:t>Text.</w:t></w:r></w:p>', {
    customXml: custom(prop(2, `MSIP_Label_${LABEL}_Enabled`, 'false'), prop(3, `MSIP_Label_${LABEL}_Name`, 'SECRET'), prop(4, 'PM_SecurityClassification', 'PROTECTED')),
  }));
  await expect(chip(page)).toHaveText('PROTECTED');
  expect(await chip(page).getAttribute('title')).toContain('Classification (PM_SecurityClassification): PROTECTED');
});

// ------------------------------------------------------------------ before you release

test('Before you release lists what the file carries beyond the page', async ({ page }) => {
  await openInEditor(page, await FULL());
  await openReview(page);
  const section = releaseSection(page);
  await expect(section).toContainText('1 comment by “Ann Analyst”. Comments go with the file');
  await expect(section).toContainText('2 tracked changes by “Bob Budget”. Until they are accepted or rejected, the file keeps the earlier wording, deleted text included');
  await expect(section).toContainText('Hidden text: 1 passage, such as “Internal note: the minister disagrees.” – Word doesn’t show or print it');
  await expect(section).toContainText('Personal information: Author “Officer Alpha”, Last saved by “Officer Bravo”, Company “Ministry of Examples”, Manager “Director Charlie”; names on comments and tracked changes (“Ann Analyst” and “Bob Budget”).');
  await expect(section).toContainText('Document properties: Title “Budget memo”; custom: Case number “FOI-2026-041”.');
  await expect(section).not.toContainText('Normal.dotm');
  await expect(section).not.toContainText('MSIP_Label');
  await expect(section).toContainText('1 custom XML data part');
  await expect(section).toContainText('1 embedded file');
});

test('comments, tracked changes and hidden text can be removed from the review', async ({ page }) => {
  await openInEditor(page, await FULL());
  await openReview(page);
  const section = releaseSection(page);
  await section.getByRole('button', { name: 'Remove all comments' }).click();
  await expect(section).not.toContainText('Ann Analyst”. Comments');
  await section.getByRole('button', { name: 'Accept all changes' }).click();
  await expect(section).not.toContainText('by “Bob Budget”');
  await section.getByRole('button', { name: 'Remove hidden text' }).click();
  await expect(section).not.toContainText('Hidden text');
  const zip = await downloadFromReview(page);
  const xml = await zip.file('word/document.xml').async('string');
  expect(xml).not.toMatch(/comment(RangeStart|Reference)/);
  expect(xml).not.toMatch(/<w:(ins|del) /);
  expect(xml).toContain('9');
  expect(xml).not.toContain('<w:delText>12');
  expect(xml).not.toContain('minister');
  expect(xml).toContain('Published.');
  expect(await zip.file('word/comments.xml').async('string')).not.toContain('Check the figure');
  // Undo brings them back, as for any edit.
});

test('personal information is removed from the download, and comes back when unticked', async ({ page }) => {
  await openInEditor(page, await FULL());
  await openReview(page);
  const box = releaseSection(page).getByRole('checkbox', { name: /Remove it from the downloaded file/ });
  await box.check();
  await expect(releaseSection(page)).toContainText('Personal information is removed from the downloaded file.');
  let zip = await downloadFromReview(page);
  const core = await zip.file('docProps/core.xml').async('string');
  expect(core).not.toContain('Officer');
  expect(core).toContain('Budget memo'); // not personal: kept
  const app = await zip.file('docProps/app.xml').async('string');
  expect(app).not.toContain('Ministry');
  expect(app).not.toContain('Director');
  expect(await zip.file('word/comments.xml').async('string')).toContain('w:author="Author" w:initials="A"');
  const xml = await zip.file('word/document.xml').async('string');
  expect(xml).not.toContain('Bob Budget');
  expect(xml).toContain('w:author="Author"');
  expect(await zip.file('word/settings.xml').async('string')).toContain('<w:removePersonalInformation/>');
  expect(await zip.file('docProps/custom.xml').async('string')).toContain('OFFICIAL-SENSITIVE'); // the label stays

  await openReview(page);
  await releaseSection(page).getByRole('checkbox', { name: /Remove it from the downloaded file/ }).uncheck();
  zip = await downloadFromReview(page);
  expect(await zip.file('docProps/core.xml').async('string')).toBe(CORE);
  expect(await zip.file('docProps/app.xml').async('string')).toBe(APP);
  expect(await zip.file('word/comments.xml').async('string')).toBe(COMMENTS);
  expect(await zip.file('word/document.xml').async('string')).toContain('Bob Budget');
});

test('a document with nothing hidden says so', async ({ page }) => {
  await openInEditor(page, await buildDocx('<w:p><w:r><w:t>Plain text.</w:t></w:r></w:p>'));
  await openReview(page);
  await expect(releaseSection(page)).toContainText('Nothing hidden: no comments, tracked changes, hidden text or personal information.');
  await expect(page.getByRole('note')).toHaveCount(0);
});

test('a restricted document lists what it carries but removes nothing until Edit anyway', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY, {
    rels: `<Relationship Id="rIdC" Type="${REL}/comments" Target="comments.xml"/><Relationship Id="rIdS" Type="${REL}/settings" Target="settings.xml"/>`,
    parts: { 'word/comments.xml': COMMENTS, 'word/settings.xml': `<w:settings ${W_NS}><w:documentProtection w:edit="readOnly" w:enforcement="1"/></w:settings>` },
  }));
  await openReview(page);
  const section = releaseSection(page);
  await expect(section.getByRole('button', { name: 'Remove all comments' })).toBeDisabled();
  await expect(section.getByRole('button', { name: 'Accept all changes' })).toBeDisabled();
  await expect(section.getByRole('checkbox')).toBeDisabled();
  await expect(section).toContainText('The document is restricted: choose Edit anyway to remove these here.');
});
