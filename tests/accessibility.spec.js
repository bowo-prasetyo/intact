// The review checks accessibility, as Word's Accessibility Checker does: pictures without alt text,
// tables without a header row, skipped heading levels, the document's language, hard-to-read
// contrast, unclear link text, blank paragraphs for spacing, floating pictures and a title. "Show"
// goes to each place; a picture's alt text and the title can be added from there.
import fs from 'node:fs';
import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import { buildDocx, openInEditor, W_NS } from './helpers.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';

const picture = (id, docPr = '', wrap = 'inline') => {
  const graphic = `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rId4"/></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm></pic:spPr></pic:pic></a:graphicData></a:graphic>`;
  const docPrXml = `<wp:docPr id="${id}" name="Picture ${id}"${docPr}`;
  return wrap === 'inline'
    ? `<w:r><w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/>${docPrXml}${graphic}</wp:inline></w:drawing></w:r>`
    : `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="5" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="952500" cy="952500"/><wp:wrapSquare wrapText="bothSides"/>${docPrXml}${graphic}</wp:anchor></w:drawing></w:r>`;
};
const DECORATIVE = '><a:extLst><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}"><adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/></a:ext></a:extLst></wp:docPr>';
const table = (header) => `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>`
  + `<w:tr>${header ? '<w:trPr><w:tblHeader/></w:trPr>' : ''}<w:tc><w:p><w:r><w:t>Region</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
const heading = (level, text) => `<w:p><w:pPr><w:pStyle w:val="Heading${level}"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;

const BODY = heading(1, 'Annual report')
  + `<w:p>${picture(1, '/>')}${picture(2, ' descr="Bar chart of spending by region"/>')}${picture(3, DECORATIVE)}</w:p>`
  + `<w:p>${picture(4, ' descr="Logo"/>', 'anchor')}<w:r><w:t>Next to the logo.</w:t></w:r></w:p>`
  + table(false) + table(true)
  + heading(3, 'Details')
  + '<w:p/><w:p/><w:p/>'
  + '<w:p><w:r><w:rPr><w:color w:val="C8C8C8"/></w:rPr><w:t>Faint grey footnote text.</w:t></w:r><w:r><w:rPr><w:color w:val="1F3864"/></w:rPr><w:t> Dark blue is fine.</w:t></w:r></w:p>'
  + '<w:p><w:r><w:t xml:space="preserve">For more, </w:t></w:r><w:hyperlink r:id="rId3"><w:r><w:t>click here</w:t></w:r></w:hyperlink></w:p>';

/** The heading styles (Word's built-in names), and the document's language when `lang` is given. */
async function document(body, { lang = null, title = undefined } = {}) {
  const zip = await JSZip.loadAsync(await buildDocx(body));
  let styles = await zip.file('word/styles.xml').async('string');
  styles = styles.replace('</w:styles>', [2, 3].map((n) => `<w:style w:type="paragraph" w:styleId="Heading${n}"><w:name w:val="heading ${n}"/><w:basedOn w:val="Normal"/></w:style>`).join('') + '</w:styles>');
  if (lang) styles = styles.replace('<w:sz w:val="22"/>', `<w:sz w:val="22"/><w:lang w:val="${lang}"/>`);
  zip.file('word/styles.xml', styles);
  if (title !== undefined) {
    const ct = await zip.file('[Content_Types].xml').async('string');
    zip.file('[Content_Types].xml', ct.replace('</Types>', '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>'));
    zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PKG}"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="${PKG}/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`);
    zip.file('docProps/core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">${title ? `<dc:title>${title}</dc:title>` : ''}<dc:creator>Officer Alpha</dc:creator></cp:coreProperties>`);
  }
  return zip.generateAsync({ type: 'nodebuffer' });
}

const review = (page) => page.getByRole('dialog', { name: 'Review your changes' });
const section = (page) => review(page).getByRole('region', { name: 'Accessibility' });
const item = (page, title) => section(page).getByRole('listitem').filter({ hasText: title });
const status = (page) => page.locator('.statusbar [role=status]');

async function openReview(page) {
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  await expect(section(page)).toBeVisible();
}

test('the review lists accessibility issues, as Word’s checker does', async ({ page }) => {
  await openInEditor(page, await document(BODY, { title: '' }));
  await openReview(page);
  await expect(item(page, 'Missing alternative text')).toContainText('Error: Missing alternative text. 1 picture without alt text');
  await expect(item(page, 'No header row')).toContainText('Error: No header row. 1 table without a header row');
  await expect(item(page, 'Skipped heading level')).toContainText('Warning: Skipped heading level. 1 heading skips a level');
  await expect(item(page, 'No document language')).toBeVisible();
  await expect(item(page, 'Hard-to-read text contrast')).toContainText('1 passage of text in a colour too light (or dark) for its background, such as “Faint grey footnote text.”');
  await expect(item(page, 'Unclear link text')).toContainText('such as “click here”');
  await expect(item(page, 'Repeated blank paragraphs')).toContainText('1 place where empty paragraphs make space');
  await expect(item(page, 'Pictures not in line with text')).toContainText('1 picture with text wrapping around it');
  await expect(item(page, 'No title')).toBeVisible();
});

test('Show goes to a picture without alt text and opens its alt text; added, the issue goes', async ({ page }) => {
  await openInEditor(page, await document(BODY));
  await openReview(page);
  await item(page, 'Missing alternative text').getByRole('button', { name: 'Show' }).click();
  await expect(review(page)).toHaveCount(0);
  await expect(status(page)).toHaveText('Missing alternative text');
  const alt = page.getByRole('textbox', { name: /Alt text/ });
  await expect(alt).toBeFocused();
  await alt.fill('Organisation chart');
  await alt.press('Tab');
  await openReview(page);
  await expect(item(page, 'Missing alternative text')).toHaveCount(0);
});

test('Show goes to each place in turn', async ({ page }) => {
  await openInEditor(page, await document(`${heading(1, 'One')}${heading(3, 'Three')}${heading(1, 'Again')}${heading(3, 'Three again')}`));
  await openReview(page);
  await item(page, 'Skipped heading level').getByRole('button', { name: 'Show (2)' }).click();
  await expect(status(page)).toHaveText('Skipped heading level (1 of 2)');
  const caretIn = () => page.evaluate(() => {
    const n = getSelection().anchorNode;
    return (n?.nodeType === 1 ? n : n?.parentElement)?.closest('h1, h2, h3, p')?.textContent;
  });
  expect(await caretIn()).toBe('Three');
  await openReview(page);
  await item(page, 'Skipped heading level').getByRole('button', { name: 'Show (2)' }).click();
  await expect(status(page)).toHaveText('Skipped heading level (2 of 2)');
  expect(await caretIn()).toBe('Three again');
});

test('a title can be added from the review, and is saved in the properties', async ({ page }) => {
  await openInEditor(page, await document('<w:p><w:r><w:t>Text.</w:t></w:r></w:p>', { lang: 'en-GB', title: '' }));
  await openReview(page);
  await item(page, 'No title').getByRole('textbox', { name: 'Document title' }).fill('Annual report 2026');
  await item(page, 'No title').getByRole('button', { name: 'Add title' }).click();
  await expect(section(page)).toContainText('No accessibility issues found');
  const [dl] = await Promise.all([page.waitForEvent('download'), review(page).getByRole('button', { name: 'Download', exact: true }).click()]);
  const core = await (await JSZip.loadAsync(fs.readFileSync(await dl.path()))).file('docProps/core.xml').async('string');
  expect(core).toContain('<dc:title>Annual report 2026</dc:title>');
  expect(core).toContain('<dc:creator>Officer Alpha</dc:creator>');
});

test('an accessible document says so', async ({ page }) => {
  const body = heading(1, 'Report') + heading(2, 'Summary') + `<w:p>${picture(1, ' descr="Map of the regions"/>')}</w:p>${table(true)}`
    + '<w:p><w:r><w:t xml:space="preserve">See </w:t></w:r><w:hyperlink r:id="rId3"><w:r><w:t>the full budget tables</w:t></w:r></w:hyperlink></w:p>';
  await openInEditor(page, await document(body, { lang: 'en-GB', title: 'Report' }));
  await openReview(page);
  await expect(section(page)).toContainText('No accessibility issues found: pictures have alt text, tables have header rows, headings are in order.');
});
