// Cropping, rotating, flipping and picture effects, as Word's Picture Format tab sets them, written
// as Word writes them (a:srcRect, a:xfrm rot/flip, a:ln, a:effectLst, a:grayscl, a:lum) with the
// rest of the picture kept; the page shows them as near as a web page can.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

// A 40 × 20 px picture, the fixture's media/image1.png.
const PICTURE = '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="381000" cy="190500"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
  + '<wp:docPr id="7" name="Picture 7"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
  + '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="7" name="logo.png"/><pic:cNvPicPr/></pic:nvPicPr>'
  + '<pic:blipFill><a:blip r:embed="rId4"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="381000" cy="190500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
  + '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
const DOC = `<w:p><w:r><w:t xml:space="preserve">Before </w:t></w:r>${PICTURE}<w:r><w:t xml:space="preserve"> after.</w:t></w:r></w:p>`;

const status = (page) => page.locator('.statusbar [role=status]');
const panel = (page) => page.getByRole('dialog', { name: 'Layout Options' });
const pic = (page) => page.locator('.docx-page .picture');

async function openPanel(page, section, at = undefined) {
  await pic(page).click({ position: at }); // `at`: a spot clear of the handles on a small picture
  await page.getByTitle('Layout Options: wrapping, position, size and alt text').click();
  if (section) await panel(page).locator('summary').getByText(section, { exact: true }).click();
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

test('cropping by numbers: srcRect, and the frame shrinks with what is cut', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await openPanel(page, 'Crop');
  await panel(page).getByRole('spinbutton', { name: 'Left (%)' }).fill('25');
  await panel(page).getByRole('spinbutton', { name: 'Left (%)' }).press('Tab');
  await expect(status(page)).toHaveText('Picture cropped');
  expect(await pic(page).evaluate((s) => [s.style.width, s.style.height, s.querySelector('img').style.width, s.querySelector('img').style.marginLeft])).toEqual(['30px', '20px', '40px', '-10px']);
  const { xml } = await download(page);
  expect(xml).toContain('<a:blip r:embed="rId4"/><a:srcRect l="25000"/><a:stretch>');
  expect(xml).toContain('<wp:extent cx="285750" cy="190500"/>');
});

test('cropping with the handles, then uncropping by numbers, gives the picture back', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await openPanel(page, 'Crop');
  await panel(page).getByRole('button', { name: 'Crop with the handles' }).click();
  await expect(page.locator('.picture-frame')).toHaveClass(/cropping/);
  await dragBy(page, page.locator('.picture-frame .s'), 0, -10); // cut half the height from the bottom
  let { xml } = await download(page);
  expect(xml).toMatch(/<a:srcRect b="(49\d\d\d|50\d\d\d)"\/>/);
  await openPanel(page, 'Crop', { x: 10, y: 5 });
  await panel(page).getByRole('spinbutton', { name: 'Bottom (%)' }).fill('0');
  await panel(page).getByRole('spinbutton', { name: 'Bottom (%)' }).press('Tab');
  ({ xml } = await download(page));
  expect(xml).not.toContain('srcRect');
  expect(xml).toContain('<wp:extent cx="381000" cy="190500"/>');
});

test('rotating and flipping: a:xfrm rot and flipH, shown turned', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await openPanel(page, 'Rotate and flip');
  await panel(page).getByRole('button', { name: 'Rotate right 90°' }).click();
  await expect(status(page)).toHaveText('Picture rotated to 90°');
  await panel(page).getByRole('button', { name: 'Flip horizontal' }).click();
  await expect(panel(page).getByRole('button', { name: 'Flip horizontal' })).toHaveAttribute('aria-pressed', 'true');
  expect(await pic(page).evaluate((s) => s.style.transform)).toBe('rotate(90deg) scaleX(-1)');
  const { xml } = await download(page);
  expect(xml).toContain('<a:xfrm rot="5400000" flipH="1"><a:off x="0" y="0"/><a:ext cx="381000" cy="190500"/></a:xfrm>');
});

test('the rotate handle turns the picture', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await pic(page).locator('img').click();
  const box = await pic(page).boundingBox();
  const handle = page.locator('.picture-frame .rotate');
  const h = await handle.boundingBox();
  await page.mouse.move(h.x + h.width / 2, h.y + h.height / 2);
  await page.mouse.down();
  await page.mouse.move(box.x + box.width + 200, box.y + box.height / 2, { steps: 5 }); // to the right: 90°
  await page.mouse.up();
  await expect(status(page)).toHaveText(/^Picture rotated to (8[5-9]|9[0-5])°$/);
});

test('effects: border, shadow, glow, soft edges, reflection, greyscale, brightness and contrast', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await openPanel(page, 'Effects');
  await panel(page).getByRole('combobox', { name: 'Border', exact: true }).selectOption('medium');
  for (const name of ['Shadow', 'Glow', 'Soft edges', 'Reflection']) await panel(page).getByRole('checkbox', { name }).check();
  await panel(page).getByRole('combobox', { name: 'Colour' }).selectOption('greyscale');
  await panel(page).getByRole('spinbutton', { name: 'Brightness (%)' }).fill('20');
  await panel(page).getByRole('spinbutton', { name: 'Brightness (%)' }).press('Tab');
  await expect(status(page)).toHaveText('Picture effects changed');
  // Shown with the file's values: the border on the picture's edge, shadow and glow, a reflection
  // below, and Word's brightness (tones moved up by 20%) in greyscale.
  const shown = await pic(page).evaluate((s) => ({
    outline: s.querySelector('.pic-box').style.outline,
    drops: s.querySelector('.pic-shadow').style.filter,
    reflection: !!s.querySelector('.pic-reflect img'),
    colour: s.querySelector('.pic-box img').style.filter,
  }));
  expect(shown.outline).toMatch(/2px/);
  expect(shown.outline).toMatch(/solid/);
  expect(shown.drops).toBe('drop-shadow(rgba(0, 0, 0, 0.4) 2.8px 2.8px 2.7px) drop-shadow(rgba(255, 192, 0, 0.6) 0px 0px 3.3px) drop-shadow(rgba(255, 192, 0, 0.6) 0px 0px 3.3px)');
  expect(shown.reflection).toBe(true);
  expect(shown.colour).toBe('url("#intact-colour-g-1000-200")');
  expect(await page.locator('#intact-colour-g-1000-200 feFuncR').getAttribute('intercept')).toBe('0.2');
  const { xml } = await download(page);
  expect(xml).toContain('<a:blip r:embed="rId4"><a:grayscl/><a:lum bright="20000"/></a:blip>');
  expect(xml).toContain('<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln w="19050"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln><a:effectLst><a:glow rad="63500"><a:srgbClr val="FFC000"><a:alpha val="60000"/></a:srgbClr></a:glow><a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="tl" rotWithShape="0"><a:prstClr val="black"><a:alpha val="40000"/></a:prstClr></a:outerShdw><a:reflection ');
  expect(xml).toContain('<a:softEdge rad="63500"/></a:effectLst></pic:spPr>');

  // Taking them off again leaves the picture as it was.
  await page.getByTitle('Layout Options: wrapping, position, size and alt text').click();
  await panel(page).locator('summary').getByText('Effects', { exact: true }).click();
  await panel(page).getByRole('combobox', { name: 'Border', exact: true }).selectOption('none');
  for (const name of ['Shadow', 'Glow', 'Soft edges', 'Reflection']) await panel(page).getByRole('checkbox', { name }).uncheck();
  await panel(page).getByRole('combobox', { name: 'Colour' }).selectOption('normal');
  await panel(page).getByRole('spinbutton', { name: 'Brightness (%)' }).fill('0');
  await panel(page).getByRole('spinbutton', { name: 'Brightness (%)' }).press('Tab');
  const back = await download(page);
  expect(back.xml).toContain(PICTURE.replace('<w:r>', '').replace('</w:r>', ''));
});

test('Tight wrapping follows the picture\'s outline: a wrap polygon, and the text flows around it', async ({ page }) => {
  await openInEditor(page, await buildDocx('<w:p><w:r><w:t xml:space="preserve">Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.</w:t></w:r></w:p>'));
  // A round picture: a circle on a transparent square.
  const base64 = await page.evaluate(() => {
    const c = document.createElement('canvas');
    c.width = 80;
    c.height = 80;
    const g = c.getContext('2d');
    g.fillStyle = '#2b579a';
    g.beginPath();
    g.arc(40, 40, 40, 0, Math.PI * 2);
    g.fill();
    return c.toDataURL('image/png').split(',')[1];
  });
  await page.locator('.docx-page p').click({ position: { x: 2, y: 5 } });
  await page.keyboard.press('Home');
  const [chooser] = await Promise.all([page.waitForEvent('filechooser'), page.getByTitle('Insert picture').click()]);
  await chooser.setFiles({ name: 'round.png', mimeType: 'image/png', buffer: Buffer.from(base64, 'base64') });
  await pic(page).locator('img').click();
  await page.getByTitle('Layout Options: wrapping, position, size and alt text').click();
  await panel(page).getByRole('radio', { name: 'Tight' }).check();
  await expect(status(page)).toHaveText('Text wrapping: Tight, around its outline');
  expect(await pic(page).evaluate((s) => getComputedStyle(s).shapeOutside)).toMatch(/^polygon\(/);
  const { xml } = await download(page);
  const polygon = xml.match(/<wp:wrapTight wrapText="bothSides"><wp:wrapPolygon edited="0">(.*?)<\/wp:wrapPolygon><\/wp:wrapTight>/)[1];
  const points = [...polygon.matchAll(/<wp:(start|lineTo) x="(\d+)" y="(\d+)"\/>/g)].map((m) => [Number(m[2]), Number(m[3])]);
  expect(points.length).toBeGreaterThan(10); // a round outline, not a rectangle
  expect(points[0]).toEqual(points[points.length - 1]); // closed, as Word writes it
  // Near the top, the outline is narrow (the circle's top), not the full width.
  const top = points.filter(([, y]) => y < 2000).map(([x]) => x);
  expect(Math.min(...top)).toBeGreaterThan(4000);
});

test('an outline Word drew is what the text wraps to', async ({ page }) => {
  const anchored = PICTURE.replace('<wp:inline distT="0" distB="0" distL="0" distR="0">', '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="5" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>')
    .replace('<wp:docPr', '<wp:wrapTight wrapText="bothSides"><wp:wrapPolygon edited="1"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="10800" y="21600"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapTight><wp:docPr')
    .replace('</wp:inline>', '</wp:anchor>');
  await openInEditor(page, await buildDocx(`<w:p>${anchored}<w:r><w:t>Text beside a triangle.</w:t></w:r></w:p>`));
  expect(await pic(page).evaluate((s) => getComputedStyle(s).shapeOutside)).toBe('polygon(0% 0%, 0% 100%, 50% 100%, 0% 0%) content-box');
});
