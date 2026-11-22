// SVG pictures (Insert picture, or dropped), as Word 2016 and later insert one: the SVG file kept as
// it is, and a PNG drawn from it here for readers before Office 2016, the picture's a:blip pointing
// at the PNG with the SVG in its extension (asvg:svgBlip). Shown here as the PNG.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const SVG = '<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100" viewBox="0 0 200 100"><rect width="200" height="100" fill="#0e6467"/><circle cx="50" cy="50" r="30" fill="#fff"/></svg>';

test('an SVG picture is inserted with its PNG, written as Word writes one', async ({ page }) => {
  await openInEditor(page, await buildDocx('<w:p><w:r><w:t>Diagram: </w:t></w:r></w:p>'));
  await page.locator('.docx-page p').first().click();
  await page.keyboard.press('End');
  const [chooser] = await Promise.all([page.waitForEvent('filechooser'), page.getByTitle('Insert picture').click()]);
  await chooser.setFiles({ name: 'flow.svg', mimeType: 'image/svg+xml', buffer: Buffer.from(SVG) });
  await expect(page.locator('.docx-page .picture img')).toHaveCount(1);
  await expect(page.getByRole('status')).toHaveText('Picture inserted');

  const { zip, xml } = await download(page);
  // 200 × 100 px is 1905000 × 952500 EMU.
  expect(xml).toContain('<wp:extent cx="1905000" cy="952500"/>');
  const m = /<a:blip r:embed="(rIdImage\d+)"><a:extLst><a:ext uri="\{96DAC541-7B7A-43D3-8B79-37D633B846F1\}"><asvg:svgBlip xmlns:asvg="http:\/\/schemas.microsoft.com\/office\/drawing\/2016\/SVG\/main" r:embed="(rIdImage\d+)"\/><\/a:ext><\/a:extLst><\/a:blip>/.exec(xml);
  expect(m).toBeTruthy();
  expect(xml).toContain('<pic:cNvPr id="1" name="flow.svg"/>');
  const rels = await zip.file('word/_rels/document.xml.rels').async('string');
  const target = (rid) => new RegExp(`Id="${rid}"[^>]*Target="([^"]+)"`).exec(rels)?.[1] || new RegExp(`Target="([^"]+)"[^>]*Id="${rid}"`).exec(rels)?.[1];
  const png = target(m[1]);
  const svg = target(m[2]);
  expect(png).toMatch(/^media\/image\d+\.png$/);
  expect(svg).toMatch(/^media\/image\d+\.svg$/);
  expect(await zip.file(`word/${svg}`).async('string')).toBe(SVG);
  const pngBytes = await zip.file(`word/${png}`).async('uint8array');
  expect([...pngBytes.slice(1, 4)].map((c) => String.fromCharCode(c)).join('')).toBe('PNG');
  expect(await zip.file('[Content_Types].xml').async('string')).toContain('<Default Extension="svg" ContentType="image/svg+xml"/>');
});
