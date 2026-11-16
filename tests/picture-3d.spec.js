// 3-D lighting and rotation, worked out as Word's 3-D Format and 3-D Rotation describe them: a
// bevel lit per pixel by the light rig, from its direction and angle, in its material; a rotation by
// X, Y and Z angles with a perspective. And every one of Word's artistic effects, with its amount.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const PICTURE = (spPrExtra = '') => '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="1905000" cy="1428750"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
  + '<wp:docPr id="7" name="Picture 7"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
  + '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="7" name="logo.png"/><pic:cNvPicPr/></pic:nvPicPr>'
  + `<pic:blipFill><a:blip r:embed="rId4"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="1428750"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${spPrExtra}</pic:spPr></pic:pic>`
  + '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
const DOC = (extra) => `<w:p>${PICTURE(extra)}</w:p><w:p><w:r><w:t>After.</w:t></w:r></w:p>`;
// A 200 × 150 px picture with a circle bevel 12 pt wide.
const BEVEL = (light = '<a:lightRig rig="threePt" dir="t"/>') => `<a:scene3d><a:camera prst="orthographicFront"/>${light}</a:scene3d><a:sp3d><a:bevelT w="152400" h="50800" prst="circle"/></a:sp3d>`;

const status = (page) => page.locator('.statusbar [role=status]');
const panel = (page) => page.getByRole('dialog', { name: 'Layout Options' });
const pic = (page) => page.locator('.docx-page .picture');

async function openPanel(page, section) {
  await pic(page).click();
  await page.getByTitle('Layout Options: wrapping, position, size and alt text').click();
  if (section) await panel(page).locator('summary').getByText(section, { exact: true }).click();
}

/** How light (+) or dark (−) the lighting layer makes the middle of the top and the bottom edge. */
function edgeLight(page) {
  return pic(page).locator('.pic-bevel').evaluate(async (img) => {
    await img.decode();
    const c = document.createElement('canvas');
    c.width = img.naturalWidth;
    c.height = img.naturalHeight;
    const g = c.getContext('2d');
    g.drawImage(img, 0, 0);
    const at = (x, y) => {
      const [r, , , a] = g.getImageData(Math.round(x), Math.round(y), 1, 1).data;
      return (r > 128 ? 1 : -1) * a;
    };
    return { top: at(c.width / 2, 3), bottom: at(c.width / 2, c.height - 4) };
  });
}

test('a bevel is lit by its light rig: the top edge catches light from the top, the bottom is in shade', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC(BEVEL())));
  await expect(pic(page).locator('.pic-bevel')).toHaveAttribute('src', /^data:image\/png/);
  const lit = await edgeLight(page);
  expect(lit.top).toBeGreaterThan(40);
  expect(lit.bottom).toBeLessThan(-40);
});

test('turning the light round 180° lights the bottom edge instead; material and lighting are written', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC(BEVEL())));
  await openPanel(page, 'Style and 3-D');
  await panel(page).getByRole('spinbutton', { name: 'Light angle (°)' }).fill('180');
  await panel(page).getByRole('spinbutton', { name: 'Light angle (°)' }).press('Tab');
  await expect(status(page)).toHaveText('3-D format changed');
  const lit = await edgeLight(page);
  expect(lit.bottom).toBeGreaterThan(40);
  expect(lit.top).toBeLessThan(-40);

  await openPanel(page, 'Style and 3-D');
  await panel(page).getByRole('combobox', { name: 'Material' }).selectOption('metal');
  await openPanel(page, 'Style and 3-D');
  await panel(page).getByRole('combobox', { name: 'Lighting' }).selectOption('sunset');
  await openPanel(page, 'Style and 3-D');
  await panel(page).getByRole('spinbutton', { name: 'Bevel width (pt)' }).fill('8');
  await panel(page).getByRole('spinbutton', { name: 'Bevel width (pt)' }).press('Tab');
  const { xml } = await download(page);
  expect(xml).toContain('<a:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="sunset" dir="t"><a:rot lat="0" lon="0" rev="10800000"/></a:lightRig></a:scene3d><a:sp3d prstMaterial="metal"><a:bevelT w="101600" h="50800" prst="circle"/></a:sp3d>');
});

test('3-D rotation by angles: X, Y and Z rotation and perspective, written as Word writes them and shown so', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC()));
  await openPanel(page, 'Style and 3-D');
  for (const [name, value] of [['X rotation (°)', '30'], ['Y rotation (°)', '10'], ['Perspective (°)', '40']]) {
    await panel(page).getByRole('spinbutton', { name }).fill(value);
    await panel(page).getByRole('spinbutton', { name }).press('Tab');
    await expect(status(page)).toHaveText('3-D rotation changed');
    if (name !== 'Perspective (°)') await openPanel(page, 'Style and 3-D');
  }
  expect(await pic(page).evaluate((s) => s.style.transform)).toBe('perspective(549px) rotateX(-10deg) rotateY(30deg)');
  const { xml } = await download(page);
  expect(xml).toContain('<a:scene3d><a:camera prst="perspectiveFront" fov="2400000"><a:rot lat="600000" lon="1800000" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"/></a:scene3d><a:sp3d/>');
});

test('a rotation Word set by angles is shown with those angles', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC('<a:scene3d><a:camera prst="isometricOffAxis1Right"><a:rot lat="1200000" lon="19800000" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"/></a:scene3d>')));
  expect(await pic(page).evaluate((s) => s.style.transform)).toBe('rotateX(-20deg) rotateY(-30deg)');
});

test('every one of Word\'s artistic effects draws, with its amount', async ({ page }) => {
  test.setTimeout(60000);
  await openInEditor(page, await buildDocx('<w:p><w:r><w:t>Text.</w:t></w:r></w:p>'));
  const base64 = await page.evaluate(() => {
    const c = document.createElement('canvas');
    c.width = 48;
    c.height = 32;
    const g = c.getContext('2d');
    const gr = g.createLinearGradient(0, 0, 48, 32);
    gr.addColorStop(0, '#2b7');
    gr.addColorStop(1, '#c33');
    g.fillStyle = gr;
    g.fillRect(0, 0, 48, 32);
    g.fillStyle = '#fff';
    g.fillRect(12, 8, 20, 14);
    return c.toDataURL('image/png').split(',')[1];
  });
  await page.locator('.docx-page p').click();
  await page.keyboard.press('End');
  const [chooser] = await Promise.all([page.waitForEvent('filechooser'), page.getByTitle('Insert picture').click()]);
  await chooser.setFiles({ name: 'test.png', mimeType: 'image/png', buffer: Buffer.from(base64, 'base64') });
  await openPanel(page, 'Artistic effects');
  const select = panel(page).getByRole('combobox', { name: 'Artistic effect' });
  const effects = await select.locator('option').evaluateAll((os) => os.map((o) => [o.value, o.textContent]).filter(([v]) => v));
  expect(effects).toHaveLength(22);
  const drawn = new Set();
  for (const [value, label] of effects) {
    await select.selectOption(value);
    await expect(status(page)).toHaveText(`Artistic effect: ${label}`);
    drawn.add(await pic(page).locator('.pic-box img').first().getAttribute('src')); // the panel stays open
  }
  expect(drawn.size).toBe(22); // each drawn afresh
  await panel(page).getByRole('slider', { name: 'Amount' }).fill('80');
  await expect(status(page)).toHaveText('Artistic effect: Glow Edges');
  const { xml } = await download(page);
  expect(xml).toMatch(/<intact:original xmlns:intact="urn:intact:picture" r:embed="rIdImage1" effect="glowEdges" amount="80"\/>/);
});
