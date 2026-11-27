// The style menu offers the document's own paragraph styles after Normal, Title and Heading 1–3, as
// Word's style gallery does: headings 4–9 the file defines, styles in the gallery that aren't
// hidden, and styles the text uses. Applying one writes only the paragraph's style.
import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import { buildDocx, openInEditor, download, W_NS } from './helpers.js';

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${W_NS}>
<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading4"><w:name w:val="heading 4"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="9"/><w:semiHidden/><w:unhideWhenUsed/><w:qFormat/><w:rPr><w:i/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Caption"><w:name w:val="caption"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="35"/><w:semiHidden/><w:unhideWhenUsed/><w:qFormat/><w:rPr><w:i/><w:sz w:val="18"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="29"/><w:qFormat/><w:pPr><w:ind w:left="864" w:right="864"/></w:pPr><w:rPr><w:i/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="BlockText"><w:name w:val="Block Text"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/></w:style>
<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style>
<w:style w:type="paragraph" w:styleId="Secret"><w:name w:val="Secret"/><w:basedOn w:val="Normal"/><w:semiHidden/><w:qFormat/></w:style>
<w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:b/></w:rPr></w:style>
</w:styles>`;

const BODY = '<w:p><w:r><w:t xml:space="preserve">Plain paragraph</w:t></w:r></w:p>'
  + '<w:p><w:pPr><w:pStyle w:val="BlockText"/></w:pPr><w:r><w:t>In Block Text</w:t></w:r></w:p>'
  + '<w:p><w:pPr><w:pStyle w:val="Ghost"/></w:pPr><w:r><w:t>In a style the file doesn’t define</w:t></w:r></w:p>';

async function docWithStyles(body = BODY) {
  const zip = await JSZip.loadAsync(await buildDocx(body));
  zip.file('word/styles.xml', STYLES);
  return zip.generateAsync({ type: 'nodebuffer' });
}

const menu = (page) => page.getByRole('combobox', { name: 'Paragraph style' });

test('the menu offers the document’s own styles: headings, gallery styles and styles in use', async ({ page }) => {
  await openInEditor(page, await docWithStyles());
  const own = menu(page).locator('optgroup[label="In this document"] option');
  // Heading 4 first; then Word's order: Quote (29), Caption (35), Block Text (in use). Not the
  // hidden or unused ones (Footnote Text, Secret).
  await expect(own).toHaveText(['Heading 4', 'Quote', 'Caption', 'Block Text']);
  await page.getByText('In Block Text').click();
  await expect(menu(page)).toHaveValue('style:BlockText');
  // A style the menu doesn't offer is named, not "Other style".
  await page.getByText('In a style the file doesn’t define').click();
  await expect(menu(page).locator('option:checked')).toHaveText('Ghost');
  await expect(menu(page).locator('option:checked')).toBeDisabled();
});

test('applying one of the document’s styles writes only the paragraph’s style', async ({ page }) => {
  const buffer = await docWithStyles();
  await openInEditor(page, buffer);
  await page.getByText('Plain paragraph').click();
  await menu(page).selectOption({ label: 'Quote' });
  const quoted = page.locator('.docx-page p[data-style="Quote"]');
  await expect(quoted).toHaveText('Plain paragraph');
  // Shown in the style: Quote's indentation.
  // 57.6px, to within WebKit's layout precision (it gives 57.599998px).
  expect(parseFloat(await quoted.evaluate((el) => getComputedStyle(el).paddingLeft))).toBeCloseTo(57.6, 3);
  await page.getByText('In Block Text').click();
  await menu(page).selectOption({ label: 'Heading 4' });
  await expect(page.locator('.docx-page h4[data-style="Heading4"]')).toHaveText('In Block Text');
  await expect(menu(page)).toHaveValue('style:Heading4');

  const { zip, xml } = await download(page);
  expect(xml).toContain('<w:p><w:pPr><w:pStyle w:val="Quote"/></w:pPr><w:r><w:t xml:space="preserve">Plain paragraph</w:t></w:r></w:p>');
  expect(xml).toContain('<w:p><w:pPr><w:pStyle w:val="Heading4"/></w:pPr><w:r><w:t xml:space="preserve">In Block Text</w:t></w:r></w:p>');
  expect(await zip.file('word/styles.xml').async('string')).toBe(STYLES); // nothing added
});

test('Undo takes a style back, and Normal still clears it', async ({ page }) => {
  await openInEditor(page, await docWithStyles());
  await page.getByText('Plain paragraph').click();
  await menu(page).selectOption({ label: 'Caption' });
  await expect(page.locator('.docx-page p[data-style="Caption"]')).toHaveText('Plain paragraph');
  await page.keyboard.press('ControlOrMeta+z');
  await expect(page.locator('.docx-page p[data-style="Caption"]')).toHaveCount(0);
  await page.getByText('In Block Text').click();
  await menu(page).selectOption({ label: 'Normal' });
  await expect(page.getByText('In Block Text')).not.toHaveAttribute('data-style');
});
