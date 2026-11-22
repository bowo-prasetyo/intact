// Document properties (Insert > Document property…), as Word's File > Info > Properties and Advanced
// Properties > Custom: the title, subject and the like in docProps/core.xml, the document's own custom
// properties in docProps/custom.xml (added with its relationship and content type when the file has
// none), and DOCPROPERTY fields showing them, which change once a property does, and with F9 (in
// headers and footers too). Ctrl+Z puts the properties back.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const P = (t) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
const CUSTOM = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Version"><vt:lpwstr>1.0</vt:lpwstr></property><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="MSIP_Label_abc_Enabled"><vt:lpwstr>true</vt:lpwstr></property></Properties>`;
const FIELD = (name, shown) => `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> DOCPROPERTY  ${name}  \\* MERGEFORMAT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${shown}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;

async function insertMenu(page, item) {
  await page.getByRole('button', { name: /Insert a check box or date/ }).click();
  await page.getByRole('menuitem', { name: item }).click();
}

test('a custom property added, shown as a field at the caret, and written as Word writes one', async ({ page }) => {
  await openInEditor(page, await buildDocx(P('Manual for the pump, version ')));
  await page.locator('.docx-page p').first().click();
  await page.keyboard.press('End');
  await insertMenu(page, /^Document property…/);
  const dialog = page.getByRole('dialog', { name: 'Document properties' });
  await dialog.getByLabel('Title').fill('Pump manual');
  await dialog.getByRole('button', { name: 'Add a property' }).click();
  await dialog.getByLabel('Name of custom property 1').fill('Version');
  await dialog.getByLabel('Value of Version').fill('2.4');
  await dialog.getByLabel('Insert as a field').selectOption('Version');
  await dialog.getByRole('button', { name: 'Insert field' }).click();
  await expect(page.locator('.docx-page p').first()).toHaveText('Manual for the pump, version 2.4');

  const { zip, xml } = await download(page);
  expect(xml).toContain('<w:instrText xml:space="preserve"> DOCPROPERTY  Version  \\* MERGEFORMAT </w:instrText>');
  const custom = await zip.file('docProps/custom.xml').async('string');
  expect(custom).toMatch(/<property fmtid="\{D5CDD505-2E9C-101B-9397-08002B2CF9AE\}" pid="2" name="Version"><vt:lpwstr>2\.4<\/vt:lpwstr><\/property>/);
  expect(await zip.file('_rels/.rels').async('string')).toContain('Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"');
  expect(await zip.file('[Content_Types].xml').async('string')).toContain('<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>');
  // The core part, which the file didn't have, is added for the title.
  expect(await zip.file('docProps/core.xml').async('string')).toContain('<dc:title>Pump manual</dc:title>');
});

test('changing a property brings its fields up to date, footer too; the label\'s properties are kept; Ctrl+Z puts it back', async ({ page }) => {
  const footer = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p>${FIELD('Version', '1.0')}</w:p></w:ftr>`;
  const docx = await buildDocx(`<w:p><w:r><w:t xml:space="preserve">Version </w:t></w:r>${FIELD('Version', '1.0')}</w:p>`, {
    parts: { 'docProps/custom.xml': CUSTOM, 'word/footer1.xml': footer },
    rels: '<Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>',
    sectPr: '<w:footerReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:type="default" r:id="rIdF"/>',
  });
  // The package's relationship to the custom properties, as Word writes it.
  const JSZip = (await import('jszip')).default;
  const zip = await JSZip.loadAsync(docx);
  zip.file('_rels/.rels', (await zip.file('_rels/.rels').async('string')).replace('</Relationships>', '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/></Relationships>'));
  zip.file('[Content_Types].xml', (await zip.file('[Content_Types].xml').async('string')).replace('</Types>', '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>'));
  await openInEditor(page, await zip.generateAsync({ type: 'nodebuffer' }));

  await page.locator('.docx-page > p').first().click();
  await insertMenu(page, /^Document property…/);
  const dialog = page.getByRole('dialog', { name: 'Document properties' });
  // The label's own properties aren't offered.
  await expect(dialog.getByLabel('Value of Version')).toHaveValue('1.0');
  await expect(dialog.getByLabel(/MSIP/)).toHaveCount(0);
  await dialog.getByLabel('Value of Version').fill('1.1');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(page.locator('.docx-page > p').first()).toHaveText('Version 1.1');
  await expect(page.locator('.docx-hf-band .docx-hf-part').last()).toContainText('1.1');

  let saved = await download(page);
  const custom = await saved.zip.file('docProps/custom.xml').async('string');
  expect(custom).toContain('name="Version"><vt:lpwstr>1.1</vt:lpwstr>');
  expect(custom).toContain('name="MSIP_Label_abc_Enabled"><vt:lpwstr>true</vt:lpwstr>');
  expect(await saved.zip.file('word/footer1.xml').async('string')).toContain('<w:t xml:space="preserve">1.1</w:t>');

  await page.locator('.docx-page > p').first().click();
  await page.keyboard.press('ControlOrMeta+z');
  await expect(page.locator('.docx-page > p').first()).toHaveText('Version 1.0');
  saved = await download(page);
  expect(await saved.zip.file('docProps/custom.xml').async('string')).toContain('name="Version"><vt:lpwstr>1.0</vt:lpwstr>');
});
