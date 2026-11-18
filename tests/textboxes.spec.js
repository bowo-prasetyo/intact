// Text boxes and shapes with text are shown where Word puts them, with their fill and outline, and
// their text is edited in place: typing, formatting, Undo and Track Changes work in them as in the
// text. Only the text is written back, into the shape and into its copy for older readers
// (mc:Fallback); a text box left alone comes back as it was.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, documentXmlOf } from './helpers.js';

const NS = 'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml"';
const text = (t) => `<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>${t}</w:t></w:r></w:p>`;
/** A floating text box as Word 2016 writes it: the shape, and a VML copy for older readers. */
export const textbox = (inner = text('A pull quote'), { align = 'right', wrap = '<wp:wrapSquare wrapText="bothSides"/>' } = {}) => `<w:r><mc:AlternateContent ${NS}><mc:Choice Requires="wps"><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:align>${align}</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="2286000" cy="914400"/><wp:effectExtent l="0" t="0" r="0" b="0"/>${wrap}<wp:docPr id="5" name="Text Box 5"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2286000" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="DEEAF6"/></a:solidFill><a:ln w="19050"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:ln></wps:spPr><wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t"/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape id="Text Box 5" o:spid="_x0000_s1026" type="#_x0000_t202" style="position:absolute;margin-left:0;margin-top:0;width:180pt;height:72pt;z-index:251659264;mso-position-horizontal:${align}" fillcolor="#deeaf6" xmlns:o="urn:schemas-microsoft-com:office:office"><v:textbox><w:txbxContent>${inner}</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>`;
const BODY = `<w:p>${textbox()}<w:r><w:t>Text beside the box.</w:t></w:r></w:p><w:p><w:r><w:t>After.</w:t></w:r></w:p>`;

const box = (page) => page.locator('.docx-page .textbox');
const content = (page) => box(page).locator('.textbox-content');

/** The run holding the text box in a saved document.xml, as a DOM, and its two texts. */
function texts(page, xml) {
  return page.evaluate((x) => {
    const doc = new DOMParser().parseFromString(x, 'application/xml');
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    return [...doc.getElementsByTagNameNS(W, 'txbxContent')].map((t) => t.textContent);
  }, xml);
}

test('a text box is shown where Word puts it, with its fill and outline', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await expect(box(page)).toHaveCount(1);
  await expect(content(page)).toHaveText('A pull quote');
  await expect(content(page)).toHaveAttribute('contenteditable', 'true');
  const look = await box(page).evaluate((el) => {
    const cs = getComputedStyle(el);
    return { float: cs.float, bg: cs.backgroundColor, border: cs.borderTopColor, width: Math.round(el.getBoundingClientRect().width) };
  });
  expect(look).toEqual({ float: 'right', bg: 'rgb(222, 234, 246)', border: 'rgb(68, 114, 196)', width: 240 });
  // Not shown as preserved content any more.
  await expect(page.locator('.docx-page [title^="Shape or text box"]')).toHaveCount(0);
});

test('typing in a text box changes its text, in the shape and in its copy for older readers', async ({ page }) => {
  const docx = await buildDocx(BODY);
  await openInEditor(page, docx);
  await content(page).getByText('A pull quote').click();
  await page.keyboard.press('End');
  await page.keyboard.type(' that sings');
  await page.keyboard.press('ControlOrMeta+b');
  await page.keyboard.type('!');
  await expect(content(page)).toHaveText('A pull quote that sings!');
  const { xml } = await download(page);
  expect(await texts(page, xml)).toEqual(['A pull quote that sings!', 'A pull quote that sings!']);
  // The shape itself is kept: fill, outline, position and size.
  expect(xml).toContain('<a:srgbClr val="DEEAF6"/>');
  expect(xml).toContain('<wp:align>right</wp:align>');
  expect(xml).toContain('style="position:absolute;margin-left:0;margin-top:0;width:180pt;height:72pt;z-index:251659264;mso-position-horizontal:right"');
  expect(xml).toMatch(/<w:rPr><w:b\/><w:i\/><\/w:rPr><w:t( xml:space="preserve")?>!<\/w:t>/);
  expect(xml).toContain('Text beside the box.');
});

test('a text box left alone comes back exactly as it was', async ({ page }) => {
  const docx = await buildDocx(BODY);
  await openInEditor(page, docx);
  await page.locator('.docx-page > p').nth(1).click();
  await page.keyboard.press('End');
  await page.keyboard.type(' More.');
  const { xml } = await download(page);
  const runOf = (x) => page.evaluate((s) => {
    const doc = new DOMParser().parseFromString(s, 'application/xml');
    return new XMLSerializer().serializeToString(doc.getElementsByTagNameNS('http://schemas.openxmlformats.org/markup-compatibility/2006', 'AlternateContent')[0]);
  }, x);
  const original = await documentXmlOf(docx);
  expect(await runOf(xml)).toBe(await runOf(original));
  expect(xml).toContain('After. More.');
});

test('Undo takes back typing in a text box, and the box stays in its paragraph', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await content(page).getByText('A pull quote').click();
  await page.keyboard.press('End');
  await page.keyboard.type(' again');
  await page.keyboard.press('Enter');
  await page.keyboard.type('Second line');
  await expect(content(page).locator('p')).toHaveText(['A pull quote again', 'Second line']);
  await page.keyboard.press('ControlOrMeta+z');
  await page.keyboard.press('ControlOrMeta+z');
  await expect(content(page).locator('p')).toHaveText(['A pull quote']);
  // The text box is still inside the paragraph it is anchored in (not moved out by the HTML parser).
  await expect(page.locator('.docx-page > p').first().locator('.textbox')).toHaveCount(1);
  await page.keyboard.press('ControlOrMeta+y');
  await expect(content(page).locator('p').first()).toHaveText('A pull quote again');
});

test('typing in a text box while tracking is a tracked insertion in it and in its copy', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await page.getByRole('button', { name: 'Track changes' }).click();
  const dialog = page.getByRole('dialog', { name: 'Your name on tracked changes' });
  await dialog.getByRole('textbox').fill('Rita');
  await dialog.getByRole('button', { name: 'Start tracking' }).click();
  await content(page).getByText('A pull quote').click();
  await page.keyboard.press('End');
  await page.keyboard.type(' now');
  const { xml } = await download(page);
  const ins = await page.evaluate((x) => {
    const doc = new DOMParser().parseFromString(x, 'application/xml');
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    return [...doc.getElementsByTagNameNS(W, 'ins')].map((i) => ({ id: i.getAttributeNS(W, 'id'), author: i.getAttributeNS(W, 'author'), text: i.textContent }));
  }, xml);
  expect(ins.map((i) => [i.author, i.text])).toEqual([['Rita', ' now'], ['Rita', ' now']]);
  expect(ins[0].id).not.toBe(ins[1].id); // each has its own revision id
});

test('a legacy VML text box is edited the same way', async ({ page }) => {
  const vml = '<w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml" style="width:150pt;height:40pt" fillcolor="#fff2cc" strokecolor="#bf9000"><v:textbox><w:txbxContent><w:p><w:r><w:t>Old box</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r>';
  await openInEditor(page, await buildDocx(`<w:p>${vml}<w:r><w:t> after</w:t></w:r></w:p>`));
  await expect(content(page)).toHaveText('Old box');
  expect(await box(page).evaluate((el) => getComputedStyle(el).backgroundColor)).toBe('rgb(255, 242, 204)');
  await content(page).click();
  await page.keyboard.press('End');
  await page.keyboard.type(' renewed');
  const { xml } = await download(page);
  expect(await texts(page, xml)).toEqual(['Old box renewed']);
  expect(xml).toContain('fillcolor="#fff2cc"');
});

test('tables, pictures and comments can’t be added in a text box; the status bar says so', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await content(page).getByText('A pull quote').click();
  await page.keyboard.press('ControlOrMeta+Enter');
  await expect(page.locator('.statusbar [role=status]')).toHaveText('A page break can’t go in a text box here: add it in Word');
});

test('a text box in one run with text (as LibreOffice writes it) is shown and edited; the text beside it too', async ({ page }) => {
  // LibreOffice puts a run's pictures and text boxes in one w:r with the text around them; Word
  // gives each its own run. Either way it is drawn, and the text stays text.
  const shared = textbox().replace(/^<w:r>/, '').replace(/<\/w:r>$/, '');
  await openInEditor(page, await buildDocx(`<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Before </w:t>${shared}<w:t xml:space="preserve"> after.</w:t></w:r></w:p>`));
  await expect(content(page)).toHaveText('A pull quote');
  await expect(page.locator('.docx-page p strong').first()).toHaveText('Before ');
  await content(page).locator('p').click();
  await page.keyboard.press('End');
  await page.keyboard.type(' here');
  await page.locator('.docx-page > p strong', { hasText: 'after.' }).click();
  await page.keyboard.press('End');
  await page.keyboard.type(' More');
  const { xml } = await download(page);
  expect(await texts(page, xml)).toEqual(['A pull quote here', 'A pull quote here']);
  // A run each, with the run's properties: the text, the box, the text.
  expect(xml).toContain('<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Before </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><mc:AlternateContent');
  expect(xml).toMatch(/<\/mc:AlternateContent><\/w:r><w:r><w:rPr><w:b\/><\/w:rPr><w:t xml:space="preserve"> after\. More<\/w:t><\/w:r>/);
});

test('untouched, a run shared by a text box and text comes back byte for byte', async ({ page }) => {
  const shared = textbox().replace(/^<w:r>/, '').replace(/<\/w:r>$/, '');
  const docx = await buildDocx(`<w:p><w:r><w:t xml:space="preserve">Before </w:t>${shared}</w:r></w:p><w:p><w:r><w:t>Other paragraph.</w:t></w:r></w:p>`);
  await openInEditor(page, docx);
  await page.locator('.docx-page > p', { hasText: 'Other paragraph.' }).click();
  await page.keyboard.press('End');
  await page.keyboard.type('!');
  await page.keyboard.press('ControlOrMeta+z');
  const { xml } = await download(page);
  expect(xml).toBe(await documentXmlOf(docx));
});
