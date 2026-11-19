// A watermark (DRAFT, CONFIDENTIAL…), as Word's Design > Watermark puts one: WordArt in every header,
// behind the text, written as Word writes it (a VML text path, PowerPlusWaterMarkObject). Set and
// removed in Page setup; a document without a header gets one; Word's own watermark shows and can
// be changed or taken off. Drawn on each page here, and printed on every page.
import { test, expect } from '@playwright/test';
import fs from 'node:fs';
import JSZip from 'jszip';
import { buildDocx, openInEditor, downloadFile } from './helpers.js';

const BODY = '<w:p><w:r><w:t>Minutes of the meeting</w:t></w:r></w:p>';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const HDR_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w10="urn:schemas-microsoft-com:office:word"';
// A header as Word writes it with its CONFIDENTIAL watermark (shortened: no shapetype).
const WORD_HEADER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${HDR_NS}><w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr>`
  + '<w:r><w:rPr><w:noProof/></w:rPr><w:pict><v:shape id="PowerPlusWaterMarkObject1" o:spid="_x0000_s2050" type="#_x0000_t136" style="position:absolute;margin-left:0;margin-top:0;width:527.85pt;height:131.95pt;rotation:315;z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" o:allowincell="f" fillcolor="silver" stroked="f"><v:fill opacity=".5"/><v:textpath style="font-family:&quot;Calibri&quot;;font-size:1pt" string="CONFIDENTIAL"/></v:shape></w:pict></w:r>'
  + '<w:r><w:t>Acme Ltd</w:t></w:r></w:p></w:hdr>';
const WITH_HEADER = {
  parts: { 'word/header1.xml': WORD_HEADER },
  rels: `<Relationship Id="rIdH1" Type="${REL}/header" Target="header1.xml"/>`,
  sectPr: '<w:headerReference w:type="default" r:id="rIdH1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>',
};

const status = (page) => page.locator('.statusbar [role=status]');

async function setWatermark(page, text, { custom = '', layout = '' } = {}) {
  await page.getByRole('button', { name: 'Page setup' }).click();
  const dialog = page.getByRole('dialog', { name: 'Page setup' });
  await dialog.getByRole('combobox', { name: 'Watermark text' }).selectOption(text);
  if (custom) await dialog.getByRole('textbox', { name: 'Watermark’s own text' }).fill(custom);
  if (layout) await dialog.getByRole('combobox', { name: 'Watermark layout' }).selectOption(layout);
  await dialog.getByRole('button', { name: 'OK' }).click();
}

async function files(page) {
  const zip = await JSZip.loadAsync(fs.readFileSync(await downloadFile(page)));
  const read = (p) => zip.file(p)?.async('string');
  return { zip, read };
}

test('a watermark in a document without a header: a header is made for it, as Word writes one', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await setWatermark(page, 'DRAFT');
  await expect(status(page)).toHaveText('Watermark: DRAFT');
  await expect(page.locator('.watermark-mark').first()).toHaveText('DRAFT');
  await expect(page.locator('.watermark-mark').first()).toHaveClass(/diagonal/);

  const { read } = await files(page);
  const header = await read('word/header1.xml');
  expect(header).toContain('<w:pStyle w:val="Header"/>');
  expect(header).toMatch(/<v:shape id="PowerPlusWaterMarkObject\d+"[^>]*rotation:315;[^>]*fillcolor="silver" stroked="f">/);
  expect(header).toContain('string="DRAFT"');
  expect(header).toContain('<v:shapetype id="_x0000_t136"');
  expect(header).toMatch(/^<\?xml[^>]*\?>\s*<w:hdr [^>]*xmlns:v="urn:schemas-microsoft-com:vml"/);
  const doc = await read('word/document.xml');
  expect(doc).toMatch(/<w:headerReference w:type="default" r:id="rIdHeader1"\/>/);
  expect(await read('word/_rels/document.xml.rels')).toContain(`Id="rIdHeader1" Type="${REL}/header" Target="header1.xml"`);
  expect(await read('[Content_Types].xml')).toContain('<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>');
});

test('Undo takes a new watermark away, and the header made for it', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await setWatermark(page, 'SAMPLE');
  await expect(page.locator('.watermark-mark').first()).toHaveText('SAMPLE');
  await page.locator('.docx-page').focus();
  await page.keyboard.press('Control+z');
  await expect(page.locator('.watermark-mark')).toHaveCount(0);
  const { read } = await files(page);
  expect(await read('word/header1.xml')).toBeUndefined();
  expect(await read('word/document.xml')).not.toContain('headerReference');
});

test('Word’s own watermark shows, and Page setup changes it, keeping the rest of the header', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY, WITH_HEADER));
  await expect(page.locator('.watermark-mark').first()).toHaveText('CONFIDENTIAL');
  await expect(page.locator('.docx-hf-part')).toHaveText('Acme Ltd');
  await page.getByRole('button', { name: 'Page setup' }).click();
  const dialog = page.getByRole('dialog', { name: 'Page setup' });
  await expect(dialog.getByRole('combobox', { name: 'Watermark text' })).toHaveValue('CONFIDENTIAL');
  await expect(dialog.getByRole('combobox', { name: 'Watermark layout' })).toHaveValue('diagonal');
  await dialog.getByRole('button', { name: 'Cancel' }).click();

  await setWatermark(page, 'custom', { custom: 'Internal use only', layout: 'horizontal' });
  await expect(page.locator('.watermark-mark').first()).toHaveText('Internal use only');
  await expect(page.locator('.watermark-mark').first()).not.toHaveClass(/diagonal/);
  const { read } = await files(page);
  const header = await read('word/header1.xml');
  expect(header).toContain('string="Internal use only"');
  expect(header).not.toContain('string="CONFIDENTIAL"');
  expect(header).not.toContain('rotation:315');
  expect(header).toMatch(/<w:t[^>]*>Acme Ltd<\/w:t>/);
});

test('removing Word’s watermark leaves the header’s text as it was', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY, WITH_HEADER));
  await setWatermark(page, '');
  await expect(status(page)).toHaveText('Watermark removed');
  await expect(page.locator('.watermark-mark')).toHaveCount(0);
  const { read } = await files(page);
  const header = await read('word/header1.xml');
  expect(header).not.toContain('PowerPlusWaterMarkObject');
  expect(header).toMatch(/<w:t[^>]*>Acme Ltd<\/w:t>/);
});

test('the watermark prints on every page', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await setWatermark(page, 'DRAFT');
  await page.emulateMedia({ media: 'print' });
  const after = await page.locator('.docx-page').evaluate((el) => {
    const s = getComputedStyle(el, '::after');
    return { content: s.content, position: s.position };
  });
  expect(after).toEqual({ content: '"DRAFT"', position: 'fixed' });
});
