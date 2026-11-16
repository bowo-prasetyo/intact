// Picture styles, shapes, 3-D effects and artistic effects, as Word sets them, and the page showing
// effects with the file's own values (a shadow's distance, direction and colour; a turned picture
// taking the room of its turned outline).
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const PICTURE = (spPrExtra = '') => '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="381000" cy="190500"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
  + '<wp:docPr id="7" name="Picture 7"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
  + '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="7" name="logo.png"/><pic:cNvPicPr/></pic:nvPicPr>'
  + `<pic:blipFill><a:blip r:embed="rId4"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="381000" cy="190500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${spPrExtra}</pic:spPr></pic:pic>`
  + '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
const DOC = (extra) => `<w:p><w:r><w:t xml:space="preserve">Before </w:t></w:r>${PICTURE(extra)}<w:r><w:t xml:space="preserve"> after.</w:t></w:r></w:p>`;

const status = (page) => page.locator('.statusbar [role=status]');
const panel = (page) => page.getByRole('dialog', { name: 'Layout Options' });
const pic = (page) => page.locator('.docx-page .picture');

async function openPanel(page, section) {
  await pic(page).click();
  await page.getByTitle('Layout Options: wrapping, position, size and alt text').click();
  if (section) await panel(page).locator('summary').getByText(section, { exact: true }).click();
}

test('a picture style from Word\'s gallery: its border, shadow and shape written out, and shown', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC()));
  await openPanel(page, 'Style and 3-D');
  await panel(page).getByRole('combobox', { name: 'Picture style' }).selectOption('Simple Frame, White');
  await expect(status(page)).toHaveText('Picture style: Simple Frame, White');
  expect(await pic(page).locator('.pic-box').evaluate((b) => b.style.outlineColor)).toBe('rgb(255, 255, 255)');
  let { xml } = await download(page);
  expect(xml).toContain('<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln w="88900"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:miter lim="800000"/></a:ln><a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="tl" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="43000"/></a:srgbClr></a:outerShdw></a:effectLst></pic:spPr>');

  await openPanel(page, 'Style and 3-D');
  await panel(page).getByRole('combobox', { name: 'Picture style' }).selectOption('none');
  ({ xml } = await download(page));
  expect(xml).toContain(PICTURE().replace('<w:r>', '').replace('</w:r>', ''));
});

test('a shape: Oval cuts the picture to an ellipse', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC()));
  await openPanel(page, 'Style and 3-D');
  await panel(page).getByRole('combobox', { name: 'Shape' }).selectOption('ellipse');
  await expect(status(page)).toHaveText('Picture shape: Oval');
  expect(await pic(page).locator('.pic-box').evaluate((b) => b.style.borderRadius)).toBe('50%');
  const { xml } = await download(page);
  expect(xml).toContain('<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>');
});

test('3-D: a bevel and a 3-D rotation, written as Word\'s 3-D format and shown in perspective', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC()));
  await openPanel(page, 'Style and 3-D');
  await panel(page).getByRole('combobox', { name: 'Bevel' }).selectOption('circle');
  await expect(status(page)).toHaveText('3-D format changed');
  await openPanel(page, 'Style and 3-D');
  await panel(page).getByRole('combobox', { name: 'Preset' }).selectOption('perspectiveRight');
  await expect(status(page)).toHaveText('3-D rotation: Perspective: right');
  await expect(pic(page).locator('.pic-bevel')).toHaveCount(1);
  // Word's preset: turned 20° about the vertical axis, a 45° field of view (40 px / tan 22.5° ≈ 97 px).
  expect(await pic(page).evaluate((s) => s.style.transform)).toBe('perspective(97px) rotateY(-20deg)');
  const { xml } = await download(page);
  expect(xml).toContain('<a:scene3d><a:camera prst="perspectiveRight"/><a:lightRig rig="threePt" dir="t"/></a:scene3d><a:sp3d><a:bevelT w="76200" h="25400" prst="circle"/></a:sp3d></pic:spPr>');
});

test('an artistic effect is drawn into the picture, the original kept; taking it off gives the original back', async ({ page }) => {
  await openInEditor(page, await buildDocx('<w:p><w:r><w:t>Text.</w:t></w:r></w:p>'));
  const base64 = await page.evaluate(() => {
    const c = document.createElement('canvas');
    c.width = 60;
    c.height = 40;
    const g = c.getContext('2d');
    g.fillStyle = '#fff';
    g.fillRect(0, 0, 60, 40);
    g.fillStyle = '#c00';
    g.fillRect(15, 10, 30, 20);
    return c.toDataURL('image/png').split(',')[1];
  });
  await page.locator('.docx-page p').click();
  await page.keyboard.press('End');
  const [chooser] = await Promise.all([page.waitForEvent('filechooser'), page.getByTitle('Insert picture').click()]);
  await chooser.setFiles({ name: 'box.png', mimeType: 'image/png', buffer: Buffer.from(base64, 'base64') });
  const before = await pic(page).locator('img').first().getAttribute('src');
  await openPanel(page, 'Artistic effects');
  await panel(page).getByRole('combobox', { name: 'Artistic effect' }).selectOption('lineDrawing');
  await expect(status(page)).toHaveText('Artistic effect: Line Drawing');
  expect(await pic(page).locator('img').first().getAttribute('src')).not.toBe(before);
  let { zip, xml } = await download(page);
  const blip = xml.match(/<a:blip r:embed="(rIdImage\d+)"><a:extLst><a:ext uri="\{6E8D1F2A-7B3C-4D5E-9F60-A1B2C3D4E5F6\}"><intact:original xmlns:intact="urn:intact:picture" r:embed="(rIdImage\d+)" effect="lineDrawing" amount="50"\/><\/a:ext><\/a:extLst><\/a:blip>/);
  expect(blip).toBeTruthy();
  const rels = await zip.file('word/_rels/document.xml.rels').async('string');
  expect(rels).toContain(`Id="${blip[1]}"`);
  expect(rels).toContain(`Id="${blip[2]}"`); // both pictures are in the file
  expect(zip.file('word/media/image2.png')).toBeTruthy();
  expect(zip.file('word/media/image3.png')).toBeTruthy();

  await openPanel(page, 'Artistic effects');
  await panel(page).getByRole('combobox', { name: 'Artistic effect' }).selectOption('');
  await expect(status(page)).toHaveText('Artistic effect taken off');
  ({ zip, xml } = await download(page));
  expect(xml).toContain(`<a:blip r:embed="${blip[2]}"/>`);
  expect(zip.file('word/media/image3.png')).toBeNull(); // the drawn picture is no longer used
});

test('effects are shown with the file\'s values, and turning one on keeps the others as they are', async ({ page }) => {
  const shadow = '<a:effectLst><a:outerShdw blurRad="19050" dist="95250" dir="5400000" algn="t" rotWithShape="0"><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:outerShdw></a:effectLst>';
  await openInEditor(page, await buildDocx(DOC(shadow)));
  // 10 px straight down (90°), 1 px blur, red at 50%.
  expect(await pic(page).locator('.pic-shadow').evaluate((s) => s.style.filter)).toBe('drop-shadow(rgba(255, 0, 0, 0.5) 0px 10px 1px)');
  await openPanel(page, 'Effects');
  await panel(page).getByRole('checkbox', { name: 'Glow' }).check();
  const { xml } = await download(page);
  expect(xml).toContain('<a:outerShdw blurRad="19050" dist="95250" dir="5400000" algn="t" rotWithShape="0"><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:outerShdw>');
  expect(xml).toContain('<a:glow rad="63500">');
});

test('a turned picture in line with the text takes the room of its turned outline, on the page and in the file', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC()));
  await openPanel(page, 'Rotate and flip');
  await panel(page).getByRole('button', { name: 'Rotate right 90°' }).click();
  // 40 × 20 px turned upright: 10 px more on the top and bottom, 10 px less on the sides (none taken).
  expect(await pic(page).evaluate((s) => s.style.margin)).toBe('10px 0px');
  const { xml } = await download(page);
  expect(xml).toContain('<wp:effectExtent l="0" t="95250" r="0" b="95250"/>');
});
