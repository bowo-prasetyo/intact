// Floating pictures sit where Word puts them: measured from the margin, the page or the left margin
// as the file says, in the text and in headers; a dragged one stays where it is dropped. On a phone,
// nothing makes the document scroll sideways: a wide picture is shown smaller (its size in the file
// unchanged), and what lies past the page's edge is cut off there, as on paper.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const EMU = 9525; // per px
const picture = (cx, cy, frame) => `<w:r><w:drawing>${frame}<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>`
  + `${frame.includes('anchor') ? '<wp:wrapNone/>' : ''}<wp:docPr id="7" name="Picture 7"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>`
  + '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="7" name="logo.png"/><pic:cNvPicPr/></pic:nvPicPr>'
  + `<pic:blipFill><a:blip r:embed="rId4"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>`
  + `</a:graphicData></a:graphic></wp:${frame.includes('anchor') ? 'anchor' : 'inline'}></w:drawing></w:r>`;
const floating = (relativeFrom, offset) => picture(40 * EMU, 20 * EMU,
  '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="5" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
  + `<wp:positionH relativeFrom="${relativeFrom}"><wp:posOffset>${offset}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>`);
const para = (inner) => `<w:p><w:r><w:t xml:space="preserve">Text </w:t></w:r>${inner}<w:r><w:t xml:space="preserve"> and more text.</w:t></w:r></w:p>`;

/** The picture's left edge, from the paragraph's (the left margin) and from the page's. */
const where = (page, n = 0) => page.evaluate((i) => {
  const pic = document.querySelectorAll('.docx-page .picture')[i];
  const x = pic.getBoundingClientRect().left;
  return {
    fromMargin: Math.round(x - pic.closest('p').getBoundingClientRect().left),
    fromPage: Math.round(x - document.querySelector('.docx-sheet').getBoundingClientRect().left),
  };
}, n);

test('a picture in front of the text is placed from the margin, or from the page edge, as the file says', async ({ page }) => {
  await openInEditor(page, await buildDocx(para(floating('column', 200 * EMU)) + para(floating('leftMargin', 10 * EMU)) + para(floating('page', 300 * EMU))));
  expect((await where(page, 0)).fromMargin).toBe(200);
  expect((await where(page, 1)).fromPage).toBe(10); // in the left margin, as Word shows it
  expect((await where(page, 2)).fromPage).toBe(300);
});

test('a floating picture dragged stays where it was dropped', async ({ page }) => {
  await openInEditor(page, await buildDocx(para(floating('column', 100 * EMU))));
  const pic = page.locator('.docx-page .picture');
  await pic.locator('img').click();
  await expect(page.locator('.picture-frame')).toBeVisible();
  const before = await where(page);
  const box = await pic.boundingBox();
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
  await page.mouse.down();
  await page.mouse.move(box.x + box.width / 2 + 30, box.y + box.height / 2, { steps: 3 });
  await page.mouse.move(box.x + box.width / 2 + 60, box.y + box.height / 2, { steps: 3 });
  await page.mouse.up();
  await expect.poll(async () => (await where(page)).fromMargin).toBe(before.fromMargin + 60);
  const { xml } = await download(page);
  expect(xml).toContain(`<wp:positionH relativeFrom="margin"><wp:posOffset>${160 * EMU}</wp:posOffset></wp:positionH>`);
});

test.describe('on a phone', () => {
  test.use({ viewport: { width: 390, height: 664 }, hasTouch: true });

  test('nothing scrolls sideways: a wide picture is shown smaller, a far one stays on the page', async ({ page }) => {
    const wide = picture(600 * EMU, 300 * EMU, '<wp:inline distT="0" distB="0" distL="0" distR="0">');
    await openInEditor(page, await buildDocx(`<w:p>${wide}</w:p>${para(floating('column', 500 * EMU))}`));
    await expect.poll(() => page.evaluate(() => document.querySelector('.canvas').scrollWidth)).toBeLessThanOrEqual(390);
    const fits = await page.evaluate(() => [...document.querySelectorAll('.docx-page .picture')].map((p) => {
      const r = p.getBoundingClientRect();
      return r.left >= 0 && r.right <= innerWidth;
    }));
    expect(fits).toEqual([true, true]);
    const { xml } = await download(page);
    expect(xml).toContain(`<wp:extent cx="${600 * EMU}" cy="${300 * EMU}"/>`); // the file's size is unchanged
  });
});
