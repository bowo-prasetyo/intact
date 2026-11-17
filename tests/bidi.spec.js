// Right-to-left text (Arabic, Hebrew, Persian, Urdu): paragraphs (w:bidi), runs (w:rtl) and tables
// (w:bidiVisual) are shown right to left, alignment is logical as in Word, and direction wrappers
// (w:dir, w:bdo), smart tags and custom XML are kept around their text instead of dropped.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, documentXmlOf } from './helpers.js';

const ARABIC = 'مرحبا بالعالم';
const RTL_P = (jc = '', text = ARABIC) => `<w:p><w:pPr><w:bidi/>${jc ? `<w:jc w:val="${jc}"/>` : ''}</w:pPr><w:r><w:rPr><w:rtl/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;

test('a right-to-left paragraph is shown right to left, aligned at the right', async ({ page }) => {
  await openInEditor(page, await buildDocx(`${RTL_P()}<w:p><w:r><w:t>English.</w:t></w:r></w:p>`));
  const p = page.locator('.docx-page p').first();
  await expect(p).toHaveAttribute('dir', 'rtl');
  expect(await p.evaluate((el) => [getComputedStyle(el).direction, getComputedStyle(el).textAlign])).toEqual(['rtl', 'start']);
  await expect(p.locator('span[dir="rtl"]')).toHaveText(ARABIC);
  await expect(page.locator('.docx-page p').nth(1)).not.toHaveAttribute('dir', 'rtl');
});

test('Word’s left and right are logical in a right-to-left paragraph, and come back as they were', async ({ page }) => {
  const docx = await buildDocx(`${RTL_P('left')}${RTL_P('right', 'שלום עולם')}${RTL_P('center', 'سلام')}`);
  await openInEditor(page, docx);
  const align = (i) => page.locator('.docx-page p').nth(i).evaluate((el) => {
    const r = el.getBoundingClientRect();
    const range = document.createRange();
    range.selectNodeContents(el);
    const t = range.getBoundingClientRect();
    return t.right > r.right - 2 - parseFloat(getComputedStyle(el).paddingRight) ? 'right' : t.left < r.left + 2 + parseFloat(getComputedStyle(el).paddingLeft) ? 'left' : 'center';
  });
  expect(await align(0)).toBe('right'); // "left" = start = right
  expect(await align(1)).toBe('left'); // "right" = end = left
  const { xml } = await download(page);
  const original = await documentXmlOf(docx);
  for (const jc of ['left', 'right', 'center']) expect(xml.includes(`<w:jc w:val="${jc}"/>`)).toBe(original.includes(`<w:jc w:val="${jc}"/>`));
});

test('the alignment buttons align as shown in a right-to-left paragraph', async ({ page }) => {
  await openInEditor(page, await buildDocx(RTL_P()));
  await page.locator('.docx-page p').first().click();
  await expect(page.getByTitle('Align right (Ctrl+R)')).toHaveClass(/on/); // the start of a right-to-left line
  await page.getByTitle('Align left (Ctrl+L)').click();
  await expect(page.getByTitle('Align left (Ctrl+L)')).toHaveClass(/on/);
  const { xml } = await download(page);
  expect(xml).toContain('<w:bidi/><w:jc w:val="right"/>'); // the end, at the left
});

test('a right-to-left table keeps its first column at the right', async ({ page }) => {
  const tbl = '<w:tbl><w:tblPr><w:bidiVisual/><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>'
    + '<w:tr><w:tc><w:p><w:r><w:t>First</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Second</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
  await openInEditor(page, await buildDocx(tbl));
  await expect(page.locator('table.docx-table')).toHaveAttribute('dir', 'rtl');
  const [first, second] = await page.locator('table.docx-table td').evaluateAll((tds) => tds.map((td) => td.getBoundingClientRect().left));
  expect(first).toBeGreaterThan(second);
});

test('direction wrappers, smart tags and custom XML are kept around their text, also when it is edited', async ({ page }) => {
  const body = `<w:p><w:pPr><w:bidi/></w:pPr><w:dir w:val="rtl"><w:r><w:t>${ARABIC}</w:t></w:r></w:dir><w:r><w:t xml:space="preserve"> </w:t></w:r><w:bdo w:val="ltr"><w:r><w:t>ABC</w:t></w:r></w:bdo></w:p>`
    + '<w:p><w:smartTag w:uri="urn:schemas-microsoft-com:office:smarttags" w:element="place"><w:smartTagPr><w:attr w:name="type" w:val="city"/></w:smartTagPr><w:r><w:t>Jakarta</w:t></w:r></w:smartTag><w:r><w:t xml:space="preserve"> is </w:t></w:r>'
    + '<w:customXml w:uri="urn:example" w:element="term"><w:customXmlPr><w:attr w:name="id" w:val="7"/></w:customXmlPr><w:r><w:t>defined</w:t></w:r></w:customXml><w:r><w:t>.</w:t></w:r></w:p>';
  await openInEditor(page, await buildDocx(body));
  await expect(page.locator('.docx-wrap[dir="ltr"]')).toHaveCSS('unicode-bidi', 'bidi-override');
  // Typing inside a wrapper keeps the text in it.
  await page.locator('.docx-wrap', { hasText: 'Jakarta' }).click();
  await page.evaluate(() => {
    const t = [...document.querySelectorAll('.docx-wrap')].find((w) => w.textContent === 'Jakarta').firstChild;
    const node = t.nodeType === 3 ? t : t.firstChild;
    getSelection().collapse(node, node.data.length);
  });
  await page.keyboard.type(' Raya');
  const { xml } = await download(page);
  expect(xml).toContain(`<w:dir w:val="rtl"><w:r><w:t xml:space="preserve">${ARABIC}</w:t></w:r></w:dir>`);
  expect(xml).toContain('<w:bdo w:val="ltr"><w:r><w:t xml:space="preserve">ABC</w:t></w:r></w:bdo>');
  expect(xml).toMatch(/<w:smartTag w:uri="urn:schemas-microsoft-com:office:smarttags" w:element="place"><w:smartTagPr><w:attr w:name="type" w:val="city"\/><\/w:smartTagPr><w:r><w:t xml:space="preserve">Jakarta Raya<\/w:t><\/w:r><\/w:smartTag>/);
  expect(xml).toContain('<w:customXml w:uri="urn:example" w:element="term"><w:customXmlPr><w:attr w:name="id" w:val="7"/></w:customXmlPr><w:r><w:t xml:space="preserve">defined</w:t></w:r></w:customXml>');
});
