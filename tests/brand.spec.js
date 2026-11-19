// A business's own look: a logo in the header or footer (a letterhead), a brand colour typed by
// its code for text (remembered for the next document), and shading and borders for table cells
// (a quote's or invoice's heading row and total).
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, W_NS } from './helpers.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const part = (root, inner) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${root} ${W_NS}>${inner}</w:${root}>`;
const HEADED = {
  parts: {
    'word/header1.xml': part('hdr', '<w:p><w:r><w:t>Acme Studio</w:t></w:r></w:p>'),
    'word/footer1.xml': part('ftr', '<w:p><w:r><w:t>1 High Street</w:t></w:r></w:p>'),
  },
  rels: `<Relationship Id="rIdH1" Type="${REL}/header" Target="header1.xml"/><Relationship Id="rIdF1" Type="${REL}/footer" Target="footer1.xml"/>`,
  sectPr: '<w:headerReference w:type="default" r:id="rIdH1"/><w:footerReference w:type="default" r:id="rIdF1"/>',
};
const BODY = '<w:p><w:r><w:t>Dear client, here is our quote.</w:t></w:r></w:p>';
const status = (page) => page.locator('.statusbar [role=status]');
const header = (page) => page.locator('.docx-hf-band.header .docx-hf-part:visible');
const footer = (page) => page.locator('.docx-hf-band.footer .docx-hf-part:visible');

/** A PNG of `w` × `h` pixels, drawn in the page. */
async function png(page, w, h) {
  const base64 = await page.evaluate(([cw, ch]) => {
    const c = document.createElement('canvas');
    c.width = cw;
    c.height = ch;
    const g = c.getContext('2d');
    g.fillStyle = '#0e6467';
    g.fillRect(0, 0, cw, ch);
    return c.toDataURL('image/png').split(',')[1];
  }, [w, h]);
  return Buffer.from(base64, 'base64');
}

async function insertPicture(page, files) {
  const [chooser] = await Promise.all([page.waitForEvent('filechooser'), page.getByTitle('Insert picture').click()]);
  await chooser.setFiles(files);
}

async function selectText(page, text) {
  await page.evaluate((t) => {
    const walker = document.createTreeWalker(document.querySelector('.docx-page'), NodeFilter.SHOW_TEXT);
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      const i = n.data.indexOf(t);
      if (i < 0) continue;
      const r = document.createRange();
      r.setStart(n, i);
      r.setEnd(n, i + t.length);
      n.parentElement.closest('[contenteditable="true"]').focus();
      getSelection().removeAllRanges();
      getSelection().addRange(r);
      return;
    }
  }, text);
}

// ------------------------------------------------------------------ a logo in the header or footer

test('a logo goes in the header, saved with the header’s own relationship; the text is untouched', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY, HEADED));
  await header(page).getByText('Acme Studio').click();
  await page.keyboard.press('Home');
  await insertPicture(page, { name: 'logo.png', mimeType: 'image/png', buffer: await png(page, 60, 30) });
  await expect(status(page)).toHaveText('Picture inserted');
  const img = header(page).locator('img');
  await expect(img).toHaveCount(1);
  expect(await img.evaluate((i) => i.naturalWidth)).toBe(60); // drawn from the header's relationship

  const { zip, xml } = await download(page);
  const hdr = await zip.file('word/header1.xml').async('string');
  expect(hdr).toContain('<a:blip r:embed="rIdImage1"/>');
  expect(hdr).toMatch(/<\/w:drawing><\/w:r><w:r><w:t( xml:space="preserve")?>Acme Studio<\/w:t>/);
  // The fixture's text already has media/image1.png: the logo is a file of its own.
  expect(await zip.file('word/_rels/header1.xml.rels').async('string')).toContain(`<Relationship Id="rIdImage1" Type="${REL}/image" Target="media/image2.png"/>`);
  expect(zip.file('word/media/image2.png')).toBeTruthy();
  expect(await zip.file('[Content_Types].xml').async('string')).toContain('<Default Extension="png" ContentType="image/png"/>');
  // The text and its relationships aren't changed: the picture is the header's.
  expect(xml).not.toContain('rIdImage');
  expect(await zip.file('word/_rels/document.xml.rels').async('string')).not.toContain('rIdImage');
});

test('a logo in the footer and one in the text each go to their own part; Undo takes the footer’s out', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY, HEADED));
  await page.locator('.docx-page > p').first().click();
  await page.keyboard.press('End');
  await insertPicture(page, { name: 'a.png', mimeType: 'image/png', buffer: await png(page, 10, 10) });
  await footer(page).getByText('1 High Street').click();
  await page.keyboard.press('End');
  await insertPicture(page, { name: 'b.png', mimeType: 'image/png', buffer: await png(page, 20, 10) });
  await expect(footer(page).locator('img')).toHaveCount(1);
  let { zip } = await download(page);
  const docRels = await zip.file('word/_rels/document.xml.rels').async('string');
  const ftrRels = await zip.file('word/_rels/footer1.xml.rels').async('string');
  const target = (rels) => rels.match(/Id="rIdImage1" Type="[^"]+\/image" Target="(media\/image\d+\.png)"/)?.[1];
  expect(target(docRels)).toBeTruthy();
  expect(target(ftrRels)).toBeTruthy();
  expect(target(ftrRels)).not.toBe(target(docRels)); // a file of its own
  const footerImage = `word/${target(ftrRels)}`;
  expect(await zip.file('word/footer1.xml').async('string')).toContain('r:embed="rIdImage1"');

  await footer(page).getByText('1 High Street').click();
  await page.keyboard.press('ControlOrMeta+z');
  await expect(footer(page).locator('img')).toHaveCount(0);
  ({ zip } = await download(page));
  expect(zip.file(footerImage)).toBeNull();
  expect(zip.file('word/_rels/footer1.xml.rels')).toBeNull();
  expect(await zip.file('word/footer1.xml').async('string')).not.toContain('<w:drawing');
});

// ------------------------------------------------------------------ more colours

test('a brand colour typed by its code colours the text, and is offered again in the next document', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await selectText(page, 'our quote');
  await page.getByRole('button', { name: 'Font colour options' }).click();
  const palette = page.getByRole('dialog', { name: 'Font colour' });
  const code = palette.getByLabel('Colour code, such as #0E6467');
  await code.fill('0e6');
  await expect(palette.getByRole('button', { name: 'Apply' })).toBeEnabled(); // short codes are fine
  await code.fill('#zz');
  await expect(palette.getByRole('button', { name: 'Apply' })).toBeDisabled();
  await code.fill('#0e6467');
  await code.press('Enter');
  await expect(palette).toHaveCount(0);
  const { xml } = await download(page);
  expect(xml).toMatch(/<w:color w:val="0E6467"\/><\/w:rPr><w:t( xml:space="preserve")?>our quote<\/w:t>/);

  // Remembered in this browser, for the next document.
  await openInEditor(page, await buildDocx(BODY));
  await selectText(page, 'Dear');
  await page.getByRole('button', { name: 'Font colour options' }).click();
  await page.getByRole('dialog', { name: 'Font colour' }).getByRole('button', { name: 'Your colour #0E6467' }).click();
  const second = await download(page);
  expect(second.xml).toMatch(/<w:color w:val="0E6467"\/><\/w:rPr><w:t( xml:space="preserve")?>Dear<\/w:t>/);
});

// ------------------------------------------------------------------ cell shading and borders

const TABLE = '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>'
  + ['Item|Price', 'Design|500', 'Total|500'].map((row) => `<w:tr>${row.split('|').map((t) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`).join('')}</w:tr>`).join('')
  + '</w:tbl><w:p/>';

async function cellMenu(page) {
  await page.getByTitle('Table layout: insert or delete rows and columns').click();
  await page.getByRole('menuitem', { name: 'Shading and borders…' }).click();
  return page.getByRole('dialog', { name: 'Shading and borders' });
}

const cellsOf = (xml) => [...xml.replaceAll(' xml:space="preserve"', '').matchAll(/<w:tc>(.*?)<\/w:tc>/g)].map((m) => m[1]);

test('shading and a bottom border on the heading row, written in the cells as Word writes them', async ({ page }) => {
  await openInEditor(page, await buildDocx(TABLE));
  // The heading row: from "Item" to "Price".
  await page.evaluate(() => {
    const [a, b] = [...document.querySelectorAll('.docx-page td')];
    const r = document.createRange();
    r.setStart(a.querySelector('p').firstChild.firstChild ?? a.querySelector('p').firstChild, 0);
    r.setEnd(b.querySelector('p').firstChild.firstChild ?? b.querySelector('p').firstChild, 2);
    a.closest('[contenteditable="true"]').focus();
    getSelection().removeAllRanges();
    getSelection().addRange(r);
  });
  let menu = await cellMenu(page);
  await menu.getByLabel('Colour code, such as #0E6467').fill('#DDEEEE');
  await menu.getByRole('button', { name: 'Apply' }).click();
  await expect(status(page)).toHaveText('Shading #DDEEEE on the selected cells');
  const cells = page.locator('.docx-page td');
  await expect(cells.nth(0)).toHaveCSS('background-color', 'rgb(221, 238, 238)');
  await expect(cells.nth(1)).toHaveCSS('background-color', 'rgb(221, 238, 238)');
  await expect(cells.nth(2)).not.toHaveCSS('background-color', 'rgb(221, 238, 238)');

  menu = await cellMenu(page);
  await menu.getByRole('button', { name: 'Bottom border' }).click();
  await expect(status(page)).toHaveText('Bottom border on the selected cells');
  const ownBottom = (i) => cells.nth(i).evaluate((td) => td.style.borderBottom); // over the editor's light gridlines
  await expect.poll(() => ownBottom(0)).toBe('0.5pt solid');

  const { xml } = await download(page);
  const [c1, c2, c3] = cellsOf(xml);
  for (const c of [c1, c2]) {
    expect(c).toContain('<w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:tcBorders><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="DDEEEE"/></w:tcPr>');
  }
  expect(c3).toBe('<w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Design</w:t></w:r></w:p>');

  // Undo takes them off again, one at a time.
  await cells.nth(4).click();
  await page.keyboard.press('ControlOrMeta+z');
  await expect.poll(() => ownBottom(0)).toBe('');
  await page.keyboard.press('ControlOrMeta+z');
  await expect(cells.nth(0)).not.toHaveCSS('background-color', 'rgb(221, 238, 238)');
});

test('outside borders go round the selection only; No border and No colour take them off', async ({ page }) => {
  await openInEditor(page, await buildDocx(TABLE));
  await page.locator('.docx-page td').nth(5).click(); // the total's price
  let menu = await cellMenu(page);
  await menu.getByRole('button', { name: 'All borders' }).click();
  menu = await cellMenu(page);
  await menu.locator('.swatch').first().click();
  let { xml } = await download(page);
  let total = cellsOf(xml)[5];
  expect(total).toContain('<w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tcBorders>');
  expect(total).toMatch(/<w:shd w:val="clear" w:color="auto" w:fill="[0-9A-F]{6}"\/>/);
  expect(cellsOf(xml)[4]).not.toContain('tcBorders');

  await page.locator('.docx-page td').nth(5).click(); // Download leaves the caret at the top
  menu = await cellMenu(page);
  await menu.getByRole('button', { name: 'No border' }).click();
  menu = await cellMenu(page);
  await menu.getByRole('button', { name: 'No colour' }).click();
  ({ xml } = await download(page));
  total = cellsOf(xml)[5];
  expect(total).toContain('<w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>');
  expect(total).not.toContain('w:shd');
});

test('shading while tracking is a tracked change of the cell’s formatting, which can be rejected', async ({ page }) => {
  await openInEditor(page, await buildDocx(TABLE));
  await page.getByRole('button', { name: 'Track changes' }).click();
  const name = page.getByRole('dialog', { name: 'Your name on tracked changes' });
  await name.getByRole('textbox').fill('Rita');
  await name.getByRole('button', { name: 'Start tracking' }).click();
  await page.locator('.docx-page td').first().click();
  const menu = await cellMenu(page);
  await menu.getByLabel('Colour code, such as #0E6467').fill('FFF2CC');
  await menu.getByRole('button', { name: 'Apply' }).click();
  await expect(status(page)).toHaveText('Shading #FFF2CC on the cell (tracked)');
  const { xml } = await download(page);
  expect(cellsOf(xml)[0]).toMatch(/<w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"\/><w:tcPrChange w:id="\d+" w:author="Rita" w:date="[^"]+"><w:tcPr><w:tcW w:w="3000" w:type="dxa"\/><\/w:tcPr><\/w:tcPrChange>/);

  await page.locator('.docx-page td').first().click();
  await page.getByTitle(/^Reject the change/).click();
  await expect(page.locator('.docx-page td').first()).not.toHaveCSS('background-color', 'rgb(255, 242, 204)');
  const after = await download(page);
  expect(cellsOf(after.xml)[0]).toBe('<w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Item</w:t></w:r></w:p>');
});

test('borders a cell has in the file are drawn as Word draws them', async ({ page }) => {
  const body = TABLE.replace('<w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Item',
    '<w:tcW w:w="3000" w:type="dxa"/><w:tcBorders><w:bottom w:val="double" w:sz="12" w:space="0" w:color="C00000"/></w:tcBorders></w:tcPr><w:p><w:r><w:t>Item');
  await openInEditor(page, await buildDocx(body));
  const first = page.locator('.docx-page td').first();
  await expect(first).toHaveCSS('border-bottom-style', 'double');
  await expect(first).toHaveCSS('border-bottom-color', 'rgb(192, 0, 0)');
});

test('a picture still can’t go in a text box, also one in a footer', async ({ page }) => {
  const box = '<w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml" style="width:200pt;height:40pt"><v:textbox><w:txbxContent><w:p><w:r><w:t>Boxed address</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r>';
  const opts = { ...HEADED, parts: { ...HEADED.parts, 'word/footer1.xml': part('ftr', `<w:p>${box}</w:p>`) } };
  await openInEditor(page, await buildDocx(BODY, opts));
  await page.getByText('Boxed address').click();
  await insertPicture(page, { name: 'logo.png', mimeType: 'image/png', buffer: await png(page, 10, 10) });
  await expect(status(page)).toHaveText('A picture can’t go in a text box here: add it in Word');
  await expect(page.locator('.docx-page img')).toHaveCount(0);
});
