// Laying out pictures, as in Word: a click selects a picture (a frame with handles); corners resize
// it in proportion, sides stretch it; Layout Options set its text wrapping, position, exact size and
// alt text; a floating picture can be dragged. The picture's XML changes only where it must.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

// A 40 × 20 px picture (381000 × 190500 EMU), the fixture's media/image1.png, in line with the text.
const PICTURE = (frame = 'inline') => `<w:r><w:drawing><wp:${frame} distT="0" distB="0" distL="0" distR="0"><wp:extent cx="381000" cy="190500"/><wp:effectExtent l="0" t="0" r="0" b="0"/>`
  + '<wp:docPr id="7" name="Picture 7"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
  + '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="7" name="logo.png"/><pic:cNvPicPr/></pic:nvPicPr>'
  + '<pic:blipFill><a:blip r:embed="rId4"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="381000" cy="190500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
  + `</a:graphicData></a:graphic></wp:${frame}></w:drawing></w:r>`;
const DOC = `<w:p><w:r><w:t xml:space="preserve">Before </w:t></w:r>${PICTURE()}<w:r><w:t xml:space="preserve"> after the picture, and some more text to wrap around it.</w:t></w:r></w:p><w:p><w:r><w:t>Next paragraph.</w:t></w:r></w:p>`;

const status = (page) => page.locator('.statusbar [role=status]');
const panel = (page) => page.getByRole('dialog', { name: 'Layout Options' });
const pic = (page) => page.locator('.docx-page .picture');

async function selectPicture(page) {
  await pic(page).locator('img').click();
  await expect(page.locator('.picture-frame')).toBeVisible();
}

async function dragBy(page, locator, dx, dy) {
  const box = await locator.boundingBox();
  const x = box.x + box.width / 2;
  const y = box.y + box.height / 2;
  await page.mouse.move(x, y);
  await page.mouse.down();
  await page.mouse.move(x + dx / 2, y + dy / 2, { steps: 3 });
  await page.mouse.move(x + dx, y + dy, { steps: 3 });
  await page.mouse.up();
}

test('a click selects a picture; a corner resizes it in proportion, and Undo takes it back', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await selectPicture(page);
  await expect(status(page)).toHaveText('Picture: drag a corner to resize, or open Layout Options for wrapping and position');
  await expect(page.locator('.picture-frame .handle:not(.rotate)')).toHaveCount(8);
  await expect(page.locator('.picture-frame .handle.rotate')).toHaveCount(1);
  await dragBy(page, page.locator('.picture-frame .se'), 40, 5); // wider: the height follows
  await expect(status(page)).toHaveText('Picture resized to 2.1 × 1.1 cm');
  expect(await pic(page).locator('img').evaluate((i) => [i.style.width, i.style.height])).toEqual(['80px', '40px']);
  let { xml } = await download(page);
  expect(xml).toContain('<wp:extent cx="762000" cy="381000"/>');
  expect(xml).toContain('<a:ext cx="762000" cy="381000"/>');

  await page.keyboard.press('ControlOrMeta+z');
  ({ xml } = await download(page));
  expect(xml).toContain('<wp:extent cx="381000" cy="190500"/>');
});

test('a side handle stretches one way only', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await selectPicture(page);
  await dragBy(page, page.locator('.picture-frame .s'), 0, 20);
  expect(await pic(page).locator('img').evaluate((i) => [i.style.width, i.style.height])).toEqual(['40px', '40px']);
});

test('Layout Options: Square wrapping floats the picture, written as Word anchors it', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await selectPicture(page);
  await page.getByTitle('Layout Options: wrapping, position, size and alt text').click();
  await panel(page).getByRole('radio', { name: 'Square' }).check();
  await expect(status(page)).toHaveText('Text wrapping: Square');
  await expect(pic(page)).toHaveAttribute('data-wrap', 'square');
  expect(await pic(page).evaluate((s) => getComputedStyle(s).float)).toBe('left');
  const { xml } = await download(page);
  expect(xml).toContain('<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="');
  expect(xml).toMatch(/behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"\/><wp:positionH relativeFrom="column"><wp:posOffset>0<\/wp:posOffset><\/wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0<\/wp:posOffset><\/wp:positionV><wp:extent cx="381000" cy="190500"\/><wp:effectExtent l="0" t="0" r="0" b="0"\/><wp:wrapSquare wrapText="bothSides"\/><wp:docPr id="7" name="Picture 7"\/>/);
  expect(xml).not.toContain('<wp:inline');
});

test('every wrapping Word offers, and back in line with the text', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await selectPicture(page);
  await page.getByTitle('Layout Options: wrapping, position, size and alt text').click();
  const expected = {
    Tight: '<wp:wrapTight wrapText="bothSides"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/>',
    Through: '<wp:wrapThrough wrapText="bothSides"><wp:wrapPolygon edited="0">',
    'Top and bottom': '<wp:wrapTopAndBottom/>',
    'Behind text': '<wp:wrapNone/>',
    'In front of text': '<wp:wrapNone/>',
  };
  for (const [name, fragment] of Object.entries(expected)) {
    await panel(page).getByRole('radio', { name }).check();
    await expect(status(page)).toHaveText(`Text wrapping: ${name}`);
    const { xml } = await download(page);
    expect(xml).toContain(fragment);
    expect(xml).toContain(name === 'Behind text' ? 'behindDoc="1"' : 'behindDoc="0"');
    await page.getByTitle('Layout Options: wrapping, position, size and alt text').click();
  }
  await panel(page).getByRole('radio', { name: 'In line with text' }).check();
  const { xml } = await download(page);
  expect(xml).toContain('<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="381000" cy="190500"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="7" name="Picture 7"/>');
  expect(xml).not.toContain('wp:anchor');
});

test('position: centred between the margins, or at a distance; size in centimetres; alt text', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await selectPicture(page);
  await page.getByTitle('Layout Options: wrapping, position, size and alt text').click();
  await expect(panel(page).getByRole('combobox', { name: 'Horizontal' })).toBeDisabled(); // in line: no position
  await panel(page).getByRole('radio', { name: 'Top and bottom' }).check();
  await panel(page).getByRole('combobox', { name: 'Horizontal' }).selectOption('center');
  await panel(page).getByRole('combobox', { name: 'Vertical' }).selectOption('offset');
  await panel(page).getByRole('spinbutton', { name: 'Distance below the paragraph (cm)' }).fill('1.5');
  await panel(page).getByRole('spinbutton', { name: 'Distance below the paragraph (cm)' }).press('Tab');
  await panel(page).getByRole('spinbutton', { name: 'Width (cm)' }).fill('4');
  await panel(page).getByRole('spinbutton', { name: 'Width (cm)' }).press('Tab');
  await expect(panel(page).getByRole('spinbutton', { name: 'Height (cm)' })).toHaveValue('2');
  await panel(page).getByRole('textbox', { name: /Alt text/ }).fill('Company logo');
  await panel(page).getByRole('textbox', { name: /Alt text/ }).press('Tab');
  await expect(pic(page).locator('img')).toHaveAttribute('alt', 'Company logo');
  const { xml } = await download(page);
  expect(xml).toContain('<wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>540000</wp:posOffset></wp:positionV><wp:extent cx="1440000" cy="720000"/>');
  expect(xml).toContain('<wp:docPr id="7" name="Picture 7" descr="Company logo"/>');
});

test('a floating picture is dragged to a new place, measured from the margin and its paragraph', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC.replace(PICTURE(), PICTURE('anchor')
    .replace('<wp:anchor distT="0" distB="0" distL="0" distR="0">', '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="5" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>')
    .replace('<wp:docPr', '<wp:wrapSquare wrapText="bothSides"/><wp:docPr'))));
  await selectPicture(page);
  const before = await pic(page).locator('img').boundingBox();
  await dragBy(page, pic(page).locator('img'), 100, 30);
  await expect(status(page)).toHaveText('Picture moved');
  const after = await pic(page).locator('img').boundingBox();
  expect(Math.round(after.x - before.x)).toBe(100);
  const { xml } = await download(page);
  expect(xml).toMatch(/<wp:positionH relativeFrom="margin"><wp:posOffset>952500<\/wp:posOffset><\/wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>\d+<\/wp:posOffset><\/wp:positionV>/);
});

test('Delete removes a selected picture; while tracking, the deletion is tracked', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await page.getByRole('button', { name: 'Track changes' }).click();
  await page.getByRole('dialog', { name: 'Your name on tracked changes' }).getByRole('textbox').fill('Rita');
  await page.getByRole('button', { name: 'Start tracking' }).click();
  await selectPicture(page);
  await page.keyboard.press('Delete');
  const { xml } = await download(page);
  expect(xml).toMatch(/<w:del w:id="\d+" w:author="Rita" w:date="[^"]+"><w:r><w:drawing>/);
});

test('the review says a picture\'s layout changed, applied directly', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await selectPicture(page);
  await page.getByTitle('Layout Options: wrapping, position, size and alt text').click();
  await panel(page).getByRole('radio', { name: 'Square' }).check();
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  const review = page.getByRole('dialog', { name: 'Review your changes' });
  await expect(review.locator('.change')).toContainText('Layout of a picture changed (size, wrapping, position or alt text)');
  await expect(review.locator('.notes')).toContainText('Formatting changes are applied directly, never tracked.');
  await expect(review.locator('.notes')).not.toContainText('aren’t tracked');
});
