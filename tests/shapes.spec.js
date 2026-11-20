// Lines, arrows and text boxes, as Word's Insert > Shapes makes them: a line or an arrow in front of
// the text (a DrawingML shape, wps:wsp, in a wp:anchor), drawn here and moved, resized and turned as
// a picture is; a text box in line with the text, to type in. Lines, arrows, rectangles and ellipses
// already in a document are drawn too, instead of a placeholder.
import fs from 'node:fs';
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, downloadFile } from './helpers.js';

const P = (t) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
const status = (page) => page.locator('.statusbar [role=status]');
const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';

async function insert(page, item) {
  await page.getByRole('button', { name: /Insert a check box or date/ }).click();
  await page.getByRole('menuitem', { name: new RegExp(`^${item}`) }).click();
}

test('an arrow: in front of the text at the caret, selected, dragged into place, written as Word’s shape', async ({ page }) => {
  await openInEditor(page, await buildDocx(P('Label the parts of the flower.') + P('')));
  await page.locator('.docx-page p').first().click();
  await page.keyboard.press('End');
  await insert(page, 'Arrow');
  await expect(status(page)).toHaveText('Arrow inserted: drag it into place, a handle to make it longer, or Layout Options to turn it');
  const arrow = page.locator('.docx-page .picture.shape');
  await expect(arrow).toHaveCount(1);
  await expect(arrow.locator('svg line[marker-end]')).toHaveCount(1);
  await expect(arrow).toHaveCSS('width', /^151\.1/); // 4 cm
  // Selected, with the picture's frame: Layout Options without a picture's crop and effects.
  await page.getByTitle('Layout Options: wrapping, position, size and alt text').click();
  const panel = page.getByRole('dialog', { name: 'Layout Options' });
  await expect(panel.getByText('Text wrapping')).toBeVisible();
  await expect(panel.getByText('Rotate and flip')).toBeVisible();
  await expect(panel.getByText('Crop', { exact: true })).toHaveCount(0);
  await expect(panel.getByText('Effects', { exact: true })).toHaveCount(0);
  await panel.press('Escape');

  // Selected with a click, then dragged 60 px right and 20 down.
  const hit = arrow.locator('svg line').first();
  // A quarter of the way along: its middle holds the frame's resize handles.
  const box = await hit.boundingBox();
  const [gx, gy] = [box.x + box.width / 4, box.y + box.height / 2];
  await page.mouse.click(gx, gy);
  const before = await arrow.boundingBox();
  await page.mouse.move(gx, gy);
  await page.mouse.down();
  await page.mouse.move(gx + 30, gy + 10, { steps: 3 });
  await page.mouse.move(gx + 60, gy + 20, { steps: 3 });
  await page.mouse.up();
  const after = await arrow.boundingBox();
  expect(Math.round(after.x - before.x)).toBe(60);
  await expect(status(page)).toHaveText('Shape moved');

  const { xml } = await download(page);
  expect(xml).toMatch(/<wp:anchor [^>]*behindDoc="0"[^>]*>.*<wp:wrapNone\/><wp:docPr id="\d+" name="Straight Arrow Connector \d+"\/>/s);
  expect(xml).toContain(`<a:graphicData uri="${WPS}"><wps:wsp`);
  expect(xml).toContain('<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom><a:ln w="19050"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:tailEnd type="triangle"/></a:ln>');
  expect(xml).toMatch(/xmlns:wps="http:\/\/schemas\.microsoft\.com\/office\/word\/2010\/wordprocessingShape"/);
});

test('a text box: in line with the text, typed in, written with its text', async ({ page }) => {
  await openInEditor(page, await buildDocx(P('Word bank: ')));
  await page.locator('.docx-page p').first().click();
  await page.keyboard.press('End');
  await insert(page, 'Text box');
  await expect(status(page)).toHaveText('Text box inserted: type in it');
  await page.keyboard.type('cat, dog, bird');
  await expect(page.locator('.docx-page .textbox')).toContainText('cat, dog, bird');
  const { xml } = await download(page);
  expect(xml).toMatch(/<wp:inline [^>]*><wp:extent cx="1800000" cy="540000"\/>.*<wp:docPr id="\d+" name="Text Box \d+"\/>/s);
  expect(xml).toMatch(/<wps:txbx><w:txbxContent><w:p>(?:<w:pPr>.*?<\/w:pPr>)?<w:r><w:t[^>]*>cat, dog, bird<\/w:t><\/w:r><\/w:p><\/w:txbxContent><\/wps:txbx>/s);
});

test('a line inserted and opened again: drawn, not a placeholder', async ({ page }) => {
  await openInEditor(page, await buildDocx(P('Underline here.')));
  await page.locator('.docx-page p').first().click();
  await insert(page, 'Line');
  await expect(page.locator('.docx-page .picture.shape svg line:not([marker-end])')).toHaveCount(2); // the line, and its wider unseen one for clicking
  const path = await downloadFile(page);
  await page.goto('./');
  await page.setInputFiles('input[type=file]', { name: 'Line.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: fs.readFileSync(path) });
  await expect(page.locator('.docx-page .picture.shape[data-shape="line"] svg')).toHaveCount(1);
});

test('shapes Word wrote: a rectangle and an ellipse drawn with their fill and outline; Undo takes an inserted one out', async ({ page }) => {
  const shape = (prst, fill) => `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="952500" cy="476250"/><wp:docPr id="7${prst.length}" name="${prst}"/>`
    + `<a:graphic><a:graphicData uri="${WPS}"><wps:wsp xmlns:wps="${WPS}"><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm><a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>`
    + `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
  await openInEditor(page, await buildDocx(`<w:p>${shape('rect', 'FFD966')}${shape('ellipse', '9DC3E6')}</w:p>`));
  await expect(page.locator('.docx-page .picture.shape svg rect')).toHaveAttribute('fill', '#FFD966');
  await expect(page.locator('.docx-page .picture.shape svg ellipse')).toHaveAttribute('fill', '#9DC3E6');
  // Kept as they were.
  const { xml } = await download(page);
  expect(xml).toContain('<a:srgbClr val="FFD966"/>');

  await page.locator('.docx-page p').first().click();
  await page.keyboard.press('End');
  await insert(page, 'Arrow');
  await expect(page.locator('.docx-page .picture.shape')).toHaveCount(3);
  await page.locator('.docx-page').focus();
  await page.keyboard.press('Control+z');
  await expect(page.locator('.docx-page .picture.shape')).toHaveCount(2);
});
