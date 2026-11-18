// About how many pages the document prints on, and where they break: worked out from the paper size
// and margins (editor/pages.js), counted in the status bar and drawn over the page as dashed lines,
// never part of the document.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const para = (text) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const lines = (n, from = 1) => Array.from({ length: n }, (_, i) => para(`Line ${from + i} of the report.`)).join('');
// A4 with 2.54 cm margins holds about 46 lines of 11 pt text.
const TWO_PAGES = lines(70);

const count = (page) => page.locator('.statusbar .page-count');

test('the status bar counts the pages and the page shows where they break', async ({ page }) => {
  await openInEditor(page, await buildDocx(TWO_PAGES));
  await expect(count(page)).toHaveText('≈ 2 pages');
  await expect(count(page)).toHaveAttribute('title', /About 2 pages on A4 paper.*Word lays out the pages itself/);
  const marks = page.locator('.page-breaks-layer .page-estimate');
  await expect(marks).toHaveCount(1);
  await expect(marks).toHaveText('Page 2');
  // The mark is drawn over the page, never in it.
  await expect(page.locator('.docx-page .page-estimate')).toHaveCount(0);
  // It is just above the first line of the second page.
  const markTop = (await marks.boundingBox()).y;
  const following = await page.evaluate((y) => [...document.querySelectorAll('.docx-page p')].find((p) => p.getBoundingClientRect().top > y)?.textContent, markTop);
  expect(following).toMatch(/^Line \d+ of the report\.$/);
});

test('a short document is one page, with no marks', async ({ page }) => {
  await openInEditor(page, await buildDocx(lines(5)));
  await expect(count(page)).toHaveText('≈ 1 page');
  await expect(page.locator('.page-estimate')).toHaveCount(0);
});

test('page breaks, "page break before" and section breaks start new pages', async ({ page }) => {
  const body = para('One')
    + '<w:p><w:r><w:t>Two</w:t></w:r><w:r><w:br w:type="page"/></w:r><w:r><w:t>Three</w:t></w:r></w:p>'
    + '<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>Four</w:t></w:r></w:p>'
    + '<w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr><w:r><w:t>End of section</w:t></w:r></w:p>'
    + para('Five')
    + '<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr><w:r><w:t>Continuous</w:t></w:r></w:p>'
    + para('Six');
  await openInEditor(page, await buildDocx(body));
  await expect(count(page)).toHaveText('≈ 4 pages');
  await expect(page.locator('.page-estimate span')).toHaveText(['Page 2', 'Page 3', 'Page 4']);
});

test('the count follows the text as it is typed and deleted', async ({ page }) => {
  await openInEditor(page, await buildDocx(lines(40)));
  await expect(count(page)).toHaveText('≈ 1 page');
  await page.locator('.docx-page p').last().click();
  await page.keyboard.press('End');
  for (let i = 0; i < 12; i++) await page.keyboard.press('Enter');
  await expect(count(page)).toHaveText('≈ 2 pages');
  await page.keyboard.press('Control+z');
  await expect(count(page)).toHaveText('≈ 1 page');
});

test('a heading isn\'t left alone at the bottom of a page: it goes with the text it heads', async ({ page }) => {
  // How many lines fill the first page…
  await openInEditor(page, await buildDocx(lines(80)));
  await expect(page.locator('.page-estimate')).toHaveCount(1);
  const perPage = await page.evaluate(() => {
    const y = document.querySelector('.page-estimate').getBoundingClientRect().top;
    return [...document.querySelectorAll('.docx-page p')].filter((p) => p.getBoundingClientRect().bottom < y).length;
  });
  // …then a heading at each place near the end of it: it never ends the page, and where it would
  // have, it begins the next.
  let moved = 0;
  for (let n = perPage - 3; n <= perPage; n++) {
    const body = lines(n) + '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Heading</w:t></w:r></w:p>' + lines(20, 500);
    await page.getByRole('button', { name: 'Close' }).click();
    await page.setInputFiles('input[type=file]', { name: `heading${n}.docx`, mimeType: 'application/octet-stream', buffer: await buildDocx(body) });
    await expect(page.locator('.page-estimate')).toHaveCount(1);
    const around = await page.evaluate(() => {
      const y = document.querySelector('.page-estimate').getBoundingClientRect().top;
      const blocks = [...document.querySelectorAll('.docx-page > p, .docx-page > h1')];
      return { before: blocks.filter((b) => b.getBoundingClientRect().bottom <= y + 3).pop()?.textContent, after: blocks.find((b) => b.getBoundingClientRect().top >= y - 1)?.textContent };
    });
    expect(around.before, `with ${n} lines before it`).not.toBe('Heading');
    if (around.after === 'Heading') moved++;
  }
  expect(moved).toBeGreaterThan(0);
});

test('showing the breaks can be turned off, and it is remembered', async ({ page }) => {
  await openInEditor(page, await buildDocx(TWO_PAGES));
  await expect(page.locator('.page-estimate')).toHaveCount(1);
  await count(page).click();
  await expect(count(page)).toHaveAttribute('aria-pressed', 'false');
  await expect(page.locator('.page-estimate')).toHaveCount(0);
  await expect(count(page)).toHaveText('≈ 2 pages'); // still counted
  await page.reload();
  await page.setInputFiles('input[type=file]', { name: 'again.docx', mimeType: 'application/octet-stream', buffer: await buildDocx(TWO_PAGES) });
  await expect(count(page)).toHaveText('≈ 2 pages');
  await expect(page.locator('.page-estimate')).toHaveCount(0);
  await count(page).click();
  await expect(page.locator('.page-estimate')).toHaveCount(1);
});

test('the marks follow the zoom', async ({ page }) => {
  // The browser lays text out at the zoomed size, so a page can end a line or two apart.
  await openInEditor(page, await buildDocx(TWO_PAGES));
  const firstAfter = () => page.evaluate(() => {
    const y = document.querySelector('.page-estimate')?.getBoundingClientRect().top;
    const p = [...document.querySelectorAll('.docx-page p')].find((el) => el.getBoundingClientRect().top > y);
    return Number(p?.textContent.match(/\d+/)?.[0]);
  });
  const at100 = await firstAfter();
  for (const button of ['Zoom out', 'Zoom out', 'Zoom out']) await page.getByRole('button', { name: button }).click();
  await expect(count(page)).toHaveText('≈ 2 pages');
  await expect.poll(async () => Math.abs((await firstAfter()) - at100)).toBeLessThanOrEqual(3);
  await page.getByRole('button', { name: /back to 100%/i }).click();
  for (const button of ['Zoom in', 'Zoom in']) await page.getByRole('button', { name: button }).click();
  await expect.poll(async () => Math.abs((await firstAfter()) - at100)).toBeLessThanOrEqual(3);
});

test('the estimate agrees with the printout', async ({ page, browserName }) => {
  test.skip(browserName !== 'chromium', 'Only Chromium makes PDFs');
  for (const n of [40, 70, 130]) {
    await openInEditor(page, await buildDocx(lines(n)));
    await expect(count(page)).toHaveText(/≈ \d+ pages?/);
    const estimate = Number((await count(page).textContent()).match(/\d+/)[0]);
    const pdf = (await page.pdf({ preferCSSPageSize: true })).toString('latin1');
    expect((pdf.match(/\/Type\s*\/Page[^s]/g) || []).length).toBe(estimate);
  }
});

test('on a phone, where the page takes the screen\'s width, the pages aren\'t counted', async ({ page }) => {
  await page.setViewportSize({ width: 400, height: 800 });
  await openInEditor(page, await buildDocx(TWO_PAGES));
  await expect(page.locator('.docx-page')).toBeVisible();
  await expect(count(page)).toHaveCount(0);
  await expect(page.locator('.page-estimate')).toHaveCount(0);
});

test('the marks never reach the saved file', async ({ page }) => {
  await openInEditor(page, await buildDocx(TWO_PAGES));
  await expect(page.locator('.page-estimate')).toHaveCount(1);
  await page.locator('.docx-page p').first().click();
  await page.keyboard.press('Home');
  await page.keyboard.type('Edited ');
  const { xml } = await download(page);
  expect(xml).not.toContain('Page 2');
  expect(xml).toContain('Edited Line 1');
});
