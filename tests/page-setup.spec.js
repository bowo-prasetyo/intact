// Page setup for the whole document, as Word's: paper size, orientation and margins, written into
// every section's w:pgSz / w:pgMar (only what was changed; the rest of each section kept), shown on
// the page and when printing, and undone with Ctrl+Z.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

// Two sections: the first ends in a paragraph, with its own size and margins (Letter); the final is A4.
const BODY = '<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/></w:sectPr></w:pPr><w:r><w:t>Section one.</w:t></w:r></w:p>'
  + '<w:p><w:r><w:t>Section two.</w:t></w:r></w:p>';
const dialog = (page) => page.getByRole('dialog', { name: 'Page setup' });
const sections = (xml) => xml.match(/<w:sectPr>.*?<\/w:sectPr>/g);

async function openSetup(page) {
  await page.locator('.docx-page p').first().click();
  await page.getByRole('button', { name: 'Page setup' }).click();
  await expect(dialog(page)).toBeVisible();
}

test('the dialog shows the paper and margins; changed margins go to every section, nothing else', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await openSetup(page);
  await expect(dialog(page).getByLabel('Size')).toHaveValue('A4');
  await expect(dialog(page).getByLabel('Orientation')).toHaveValue('portrait');
  await expect(dialog(page).getByLabel('Top (cm)')).toHaveValue('2.54');
  await expect(dialog(page).getByLabel('Preset')).toHaveValue('Normal');
  await dialog(page).getByLabel('Preset').selectOption('Narrow');
  await expect(dialog(page).getByLabel('Left (cm)')).toHaveValue('1.27');
  await dialog(page).getByRole('button', { name: 'OK' }).click();
  await expect(page.locator('.docx-page')).toHaveCSS('padding-left', '48px');
  const { xml } = await download(page);
  const [first, last] = sections(xml);
  expect(first).toBe('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/></w:sectPr>');
  expect(last).toBe('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>');
});

test('landscape Letter for every section; the page and the printout follow; Ctrl+Z puts it back', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  const portrait = (await page.locator('.docx-sheet').boundingBox()).width;
  await openSetup(page);
  await dialog(page).getByLabel('Size').selectOption('Letter');
  await dialog(page).getByLabel('Orientation').selectOption('landscape');
  await dialog(page).getByRole('button', { name: 'OK' }).click();
  await expect(page.locator('.docx-sheet')).toHaveCSS('width', '1056px'); // 11 in
  expect((await page.locator('.docx-sheet').boundingBox()).width).toBeGreaterThan(portrait);
  expect(await page.evaluate(() => [...document.styleSheets].flatMap((s) => [...s.cssRules]).find((r) => r.type === CSSRule.PAGE_RULE)?.style.size)).toBe('1056px 816px');
  let { xml } = await download(page);
  for (const s of sections(xml)) expect(s).toContain('<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>');
  await page.locator('.docx-page p').first().click();
  await page.keyboard.press('ControlOrMeta+z');
  await expect(page.locator('.docx-sheet')).toHaveCSS('width', /^793\.7/); // A4 again (210 mm)
  await page.locator('.docx-page p').first().click();
  await page.keyboard.press('End');
  await page.keyboard.type(' Edited');
  ({ xml } = await download(page));
  const [first, last] = sections(xml);
  expect(first).toContain('<w:pgSz w:w="12240" w:h="15840"/>');
  expect(last).toContain('<w:pgSz w:w="11906" w:h="16838"/>');
});

test('Cancel, or OK with nothing changed, changes nothing; margins leaving no room are refused', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await openSetup(page);
  await dialog(page).getByRole('button', { name: 'OK' }).click();
  await openSetup(page);
  await dialog(page).getByLabel('Left (cm)').fill('9');
  await dialog(page).getByRole('button', { name: 'Cancel' }).click();
  await expect(page.locator('.app-header .dirty')).toHaveCount(0);
  await openSetup(page);
  await dialog(page).getByLabel('Left (cm)').fill('10');
  await dialog(page).getByLabel('Right (cm)').fill('10');
  await dialog(page).getByRole('button', { name: 'OK' }).click();
  await expect(page.locator('.statusbar [role=status]')).toHaveText('The margins leave less than 2.5 cm for the text: nothing was changed');
  await expect(page.locator('.app-header .dirty')).toHaveCount(0);
});
