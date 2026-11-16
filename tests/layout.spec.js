// The page shows a designed document as Word lays it out: sections in columns, paragraph borders and
// shading, drop caps, character spacing, raised and lowered text, theme colours, and the fonts
// embedded in the file. Fonts neither embedded nor on this device are named in the status bar. None
// of this changes the file: it is all read from it and kept as it is.
import fs from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, canonicalize, documentXmlOf, W_NS } from './helpers.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const SECT = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>';
const para = (t, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
const LOREM = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.';

test('a section set in two columns is shown in two columns, and saved as it was', async ({ page }) => {
  // One column, then a continuous section break into two columns (with a line between), then one column.
  const body = para('Masthead in one column')
    + `<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/>${SECT}<w:cols w:space="708"/></w:sectPr></w:pPr></w:p>`
    + para(LOREM) + para(LOREM)
    + `<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/>${SECT}<w:cols w:num="2" w:space="720" w:sep="1"/></w:sectPr></w:pPr></w:p>`
    + para('Back to one column');
  const docx = await buildDocx(body);
  await openInEditor(page, docx);
  const cols = page.locator('.docx-page .docx-columns');
  await expect(cols).toHaveCount(1);
  expect(await cols.evaluate((el) => {
    const cs = getComputedStyle(el);
    return [cs.columnCount, cs.columnGap, cs.columnRuleStyle];
  })).toEqual(['2', '48px', 'solid']);
  await expect(cols.locator('p')).toHaveCount(3); // the two paragraphs and the one ending the section
  // The two columns sit side by side.
  const [a, b] = await cols.locator('p').evaluateAll((ps) => ps.slice(0, 2).map((p) => p.getClientRects()[0].left));
  expect(Math.abs(a - b)).toBeGreaterThan(100);
  await expect(page.locator('.docx-page > p').last()).toHaveText('Back to one column');

  // Untouched, the document comes back as it was; edited, only the text changes.
  let { xml } = await download(page);
  expect(await canonicalize(page, xml)).toEqual(await canonicalize(page, await documentXmlOf(docx)));
  await cols.locator('p').first().click({ position: { x: 2, y: 5 } });
  await page.keyboard.press('Home');
  await page.keyboard.type('New: ');
  ({ xml } = await download(page));
  expect(xml).toContain(`New: ${LOREM}`);
  expect(xml).toContain('<w:cols w:num="2" w:space="720" w:sep="1"/>');
  expect(await page.evaluate((x) => {
    const doc = new DOMParser().parseFromString(x, 'application/xml');
    return doc.getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'body')[0].children.length;
  }, xml)).toBe(7); // six paragraphs and the final section: nothing wrapped
});

test('a document in two columns throughout is shown in two columns', async ({ page }) => {
  await openInEditor(page, await buildDocx(para(LOREM) + para(LOREM), { sectPr: '<w:cols w:num="2" w:space="708"/>' }));
  expect(await page.locator('.docx-page > .docx-columns').evaluate((el) => getComputedStyle(el).columnCount)).toBe('2');
});

test('paragraph borders and shading, character spacing, raised text and theme colours show as in Word', async ({ page }) => {
  const theme = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Studio"><a:themeElements><a:clrScheme name="Studio"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F2A44"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="FF0000"/></a:accent1><a:accent2><a:srgbClr val="00AA00"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`;
  const body = '<w:p><w:pPr><w:pBdr><w:top w:val="double" w:sz="6" w:space="1" w:color="2E75B6"/><w:bottom w:val="single" w:sz="12" w:space="4" w:color="C55A11"/></w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="FBE4D5"/></w:pPr><w:r><w:t>Banded tagline</w:t></w:r></w:p>'
    + '<w:p><w:r><w:rPr><w:spacing w:val="60"/><w:color w:val="999999" w:themeColor="accent1" w:themeShade="80"/></w:rPr><w:t>SPACED</w:t></w:r><w:r><w:rPr><w:position w:val="6"/><w:color w:val="00AA00" w:themeColor="accent2"/></w:rPr><w:t>up</w:t></w:r></w:p>';
  const docx = await buildDocx(body, {
    parts: { 'word/theme/theme1.xml': theme },
    rels: `<Relationship Id="rIdT" Type="${REL}/theme" Target="theme/theme1.xml"/>`,
  });
  await openInEditor(page, docx);
  const banded = await page.locator('.docx-page > p').first().evaluate((el) => {
    const cs = getComputedStyle(el);
    return {
      top: [cs.borderTopStyle, cs.borderTopColor], bottom: [cs.borderBottomStyle, cs.borderBottomWidth, cs.borderBottomColor],
      paddingBottom: cs.paddingBottom, background: cs.backgroundColor,
    };
  });
  expect(banded).toEqual({
    top: ['double', 'rgb(46, 117, 182)'], bottom: ['solid', '2px', 'rgb(197, 90, 17)'], paddingBottom: '5.33333px', background: 'rgb(251, 228, 213)',
  });
  const spaced = page.locator('.docx-page span', { hasText: 'SPACED' });
  expect(await spaced.evaluate((el) => [getComputedStyle(el).letterSpacing, getComputedStyle(el).color])).toEqual(['4px', 'rgb(128, 0, 0)']);
  expect(await page.locator('.docx-page span', { hasText: 'up' }).evaluate((el) => getComputedStyle(el).verticalAlign)).toBe('4px');

  const { xml } = await download(page);
  expect(await canonicalize(page, xml)).toEqual(await canonicalize(page, await documentXmlOf(docx)));
});

test('a drop cap is shown beside the lines that follow it', async ({ page }) => {
  const body = '<w:p><w:pPr><w:framePr w:dropCap="drop" w:lines="3" w:wrap="around" w:vAnchor="text" w:hAnchor="text"/><w:spacing w:after="0" w:line="827" w:lineRule="exact"/></w:pPr><w:r><w:rPr><w:position w:val="-8"/><w:sz w:val="110"/></w:rPr><w:t>O</w:t></w:r></w:p>'
    + para(`nce upon a deadline. ${LOREM} ${LOREM}`);
  const docx = await buildDocx(body);
  await openInEditor(page, docx);
  const cap = page.locator('.docx-page > p.drop-cap');
  await expect(cap).toHaveText('O');
  expect(await cap.evaluate((el) => getComputedStyle(el).float)).toBe('left');
  // The text after it starts to its right, on the same line as the top of the letter.
  const [capBox, textLine] = await page.locator('.docx-page > p').evaluateAll((ps) => {
    const first = document.createRange();
    first.setStart(ps[1].firstChild.firstChild || ps[1].firstChild, 0);
    first.setEnd(ps[1].firstChild.firstChild || ps[1].firstChild, 1);
    return [ps[0].getBoundingClientRect(), first.getBoundingClientRect()].map((r) => ({ top: r.top, left: r.left, right: r.right }));
  });
  expect(Math.abs(textLine.top - capBox.top)).toBeLessThan(10);
  await page.locator('.docx-page > p').nth(1).click();
  await page.keyboard.press('End');
  await page.keyboard.type(' The end.');
  const { xml } = await download(page);
  expect(xml).toContain('<w:framePr w:dropCap="drop" w:lines="3" w:wrap="around" w:vAnchor="text" w:hAnchor="text"/>');
  expect(xml).toContain('The end.');
  expect(textLine.left).toBeGreaterThan(capBox.left + 20);
});

test('a font that is not on this device is named in the status bar', async ({ page }) => {
  await openInEditor(page, await buildDocx('<w:p><w:r><w:rPr><w:rFonts w:ascii="Nowhere Display" w:hAnsi="Nowhere Display"/></w:rPr><w:t>Masthead</w:t></w:r></w:p>'));
  const note = page.locator('.statusbar .fonts-missing');
  await expect(note).toHaveText('Font shown with a substitute: Nowhere Display');
  await expect(note).toHaveAttribute('title', /^Not on this device, so shown in another font: Nowhere Display\. The document keeps them/);
});

/** A TrueType font on this machine, if any (Playwright's dependencies install some on Linux). */
function anyFont() {
  const dirs = ['/usr/share/fonts', '/Library/Fonts', '/System/Library/Fonts', 'C:\\Windows\\Fonts'];
  const walk = (dir, depth = 0) => {
    let entries = [];
    try { entries = fs.readdirSync(dir, { withFileTypes: true }); } catch { return null; }
    for (const e of entries) if (e.isFile() && /\.ttf$/i.test(e.name)) return path.join(dir, e.name);
    for (const e of entries) {
      if (e.isDirectory() && depth < 3) {
        const found = walk(path.join(dir, e.name), depth + 1);
        if (found) return found;
      }
    }
    return null;
  };
  for (const d of dirs) {
    const found = walk(d);
    if (found) return found;
  }
  return null;
}

test('a font embedded in the document is shown with it, as Word shows it', async ({ page }) => {
  const file = anyFont();
  test.skip(!file, 'no TrueType font on this machine to embed');
  // Embedded as Word embeds it: obfuscated with the font key (ECMA-376 Part 1, 17.8.1).
  const key = '{0F1E2D3C-4B5A-6978-8796-A5B4C3D2E1F0}';
  const hex = key.replace(/[^0-9A-F]/gi, '');
  const keyBytes = [];
  for (let i = 30; i >= 0; i -= 2) keyBytes.push(parseInt(hex.slice(i, i + 2), 16));
  const font = new Uint8Array(fs.readFileSync(file));
  for (let i = 0; i < 32; i++) font[i] ^= keyBytes[i % 16];
  const docx = await buildDocx('<w:p><w:r><w:rPr><w:rFonts w:ascii="Intact Embedded Test" w:hAnsi="Intact Embedded Test"/></w:rPr><w:t>Embedded face</w:t></w:r></w:p>', {
    parts: {
      'word/fontTable.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:fonts ${W_NS}><w:font w:name="Intact Embedded Test"><w:embedRegular r:id="rIdF" w:fontKey="${key}"/></w:font></w:fonts>`,
      'word/_rels/fontTable.xml.rels': `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdF" Type="${REL}/font" Target="fonts/font1.odttf"/></Relationships>`,
    },
    rels: `<Relationship Id="rIdFT" Type="${REL}/fontTable" Target="fontTable.xml"/>`,
  });
  const zip = await JSZip.loadAsync(docx);
  zip.file('word/fonts/font1.odttf', font);
  await openInEditor(page, await zip.generateAsync({ type: 'nodebuffer' }));
  expect(await page.evaluate(() => [...document.fonts].filter((f) => f.family.replace(/"/g, '') === 'Intact Embedded Test').map((f) => f.status))).toEqual(['loaded']);
  await expect(page.locator('.docx-page p')).toHaveText('Embedded face');
  await page.waitForTimeout(100);
  await expect(page.locator('.statusbar .fonts-missing')).toHaveCount(0);
  const { zip: saved } = await download(page);
  expect(saved.file('word/fonts/font1.odttf')).toBeTruthy(); // kept in the file
});
