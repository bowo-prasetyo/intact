// Zoom, as in Word's status bar: the page shown bigger or smaller, the document unchanged. The
// buttons step through Word's zooms, the percentage goes back to 100%, Ctrl+mouse wheel zooms the
// page, and the choice is remembered. What is drawn over the page (formatting marks, a picture's
// frame) stays on the text, and dragging a picture's handle resizes it by what was dragged.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

// A 40 × 20 px picture (381000 × 190500 EMU), the fixture's media/image1.png, in line with the text.
const PICTURE = '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="381000" cy="190500"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
  + '<wp:docPr id="7" name="Picture 7"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
  + '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="7" name="logo.png"/><pic:cNvPicPr/></pic:nvPicPr>'
  + '<pic:blipFill><a:blip r:embed="rId4"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="381000" cy="190500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
  + '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
const DOC = `<w:p><w:r><w:t xml:space="preserve">Before </w:t></w:r>${PICTURE}<w:r><w:t xml:space="preserve"> after the picture.</w:t></w:r></w:p><w:p><w:r><w:t>Next paragraph.</w:t></w:r></w:p>`;

const level = (page) => page.locator('.statusbar .zoom-level');
const sheetZoom = (page) => page.locator('.docx-sheet').evaluate((s) => getComputedStyle(s).zoom);

test('the zoom buttons step through Word\'s zooms; the percentage goes back to 100%; it is remembered', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await expect(level(page)).toHaveText('100%');
  const width = (await page.locator('.docx-sheet').boundingBox()).width;
  await page.getByRole('button', { name: 'Zoom in' }).click();
  await page.getByRole('button', { name: 'Zoom in' }).click();
  await expect(level(page)).toHaveText('125%');
  expect(Number(await sheetZoom(page))).toBe(1.25);
  expect((await page.locator('.docx-sheet').boundingBox()).width).toBeCloseTo(width * 1.25, 0);
  await page.reload();
  await openInEditor(page, await buildDocx(DOC));
  await expect(level(page)).toHaveText('125%');
  await page.getByRole('button', { name: 'Zoom 125%, back to 100%' }).click();
  await expect(level(page)).toHaveText('100%');
  for (let i = 0; i < 8; i++) await page.getByRole('button', { name: 'Zoom out' }).click({ force: true });
  await expect(level(page)).toHaveText('50%');
  await expect(page.getByRole('button', { name: 'Zoom out' })).toBeDisabled();
  // Zooming changes nothing in the document.
  await expect(page.locator('.app-header .dirty')).toHaveCount(0);
});

test('Ctrl+mouse wheel on the page zooms it', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  const box = await page.locator('.docx-page').boundingBox();
  await page.mouse.move(box.x + 50, box.y + 50);
  await page.keyboard.down('Control');
  await page.mouse.wheel(0, -100);
  await page.keyboard.up('Control');
  await expect(level(page)).toHaveText('110%');
  await page.keyboard.down('Control');
  await page.mouse.wheel(0, 100);
  await page.mouse.wheel(0, 100);
  await page.keyboard.up('Control');
  await expect(level(page)).toHaveText('90%');
});

test('zoomed, formatting marks stay on the text', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await page.getByRole('button', { name: 'Show formatting marks' }).click();
  for (let i = 0; i < 3; i++) await page.getByRole('button', { name: 'Zoom in' }).click();
  await expect(level(page)).toHaveText('150%');
  const para = page.locator('.docx-page p', { hasText: 'Next paragraph.' });
  const end = await para.evaluate((p) => {
    const r = document.createRange();
    r.selectNodeContents(p.firstChild.nodeType === 3 ? p.firstChild : p.querySelector('span') || p);
    const rect = r.getBoundingClientRect();
    return { right: rect.right, top: rect.top };
  });
  await expect.poll(async () => page.locator('.marks-layer .fmark.para').count()).toBe(2);
  const mark = await page.locator('.marks-layer .fmark.para').nth(1).boundingBox();
  expect(Math.abs(mark.x - end.right)).toBeLessThan(4);
  expect(Math.abs(mark.y - end.top)).toBeLessThan(4);
});

test('zoomed, a picture\'s frame fits it, and dragging a corner resizes it by what was dragged', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  for (let i = 0; i < 5; i++) await page.getByRole('button', { name: 'Zoom in' }).click();
  await expect(level(page)).toHaveText('200%');
  await expect(page.getByRole('button', { name: 'Zoom in' })).toBeDisabled();
  await page.locator('.docx-page .picture img').click();
  const img = await page.locator('.docx-page .picture img').boundingBox();
  const frame = await page.locator('.picture-frame').boundingBox();
  expect(img.width).toBeCloseTo(80, 0); // 40 px shown at 200%
  for (const k of ['x', 'y', 'width', 'height']) expect(Math.abs(frame[k] - img[k])).toBeLessThan(2.5);
  // 80 shown pixels to the right are 40 of the page's: twice as wide.
  const handle = await page.locator('.picture-frame .se').boundingBox();
  const x = handle.x + handle.width / 2;
  const y = handle.y + handle.height / 2;
  await page.mouse.move(x, y);
  await page.mouse.down();
  await page.mouse.move(x + 40, y + 5, { steps: 3 });
  await page.mouse.move(x + 80, y + 10, { steps: 3 });
  await page.mouse.up();
  expect(await page.locator('.docx-page .picture img').evaluate((i) => [i.style.width, i.style.height])).toEqual(['80px', '40px']);
  const { xml } = await download(page);
  expect(xml).toContain('<wp:extent cx="762000" cy="381000"/>');
});

test('the page prints at its size, whatever the zoom', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await page.getByRole('button', { name: 'Zoom in' }).click();
  await page.emulateMedia({ media: 'print' });
  expect(Number(await sheetZoom(page))).toBe(1);
});
