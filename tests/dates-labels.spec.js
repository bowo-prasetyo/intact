// Today's date that Word updates (a DATE field, as Word's "Update automatically"), page numbers
// shown as # and the page count as about how many pages (the file keeps its numbers), text labels
// on the toolbar's icon buttons where there is room, and the caret kept where it was after Download.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, W_NS } from './helpers.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const status = (page) => page.locator('.statusbar [role=status]');
const BODY = '<w:p><w:r><w:t>Invoice sent on </w:t></w:r></w:p>';

async function insertDate(page, { auto }) {
  await page.locator('.docx-page > p').first().click();
  await page.keyboard.press('End');
  await page.getByRole('button', { name: 'Insert a check box or date' }).click();
  const menu = page.getByRole('menu', { name: 'Insert' });
  const toggle = menu.getByRole('menuitemcheckbox', { name: /Update automatically/ });
  if ((await toggle.getAttribute('aria-checked')) !== String(auto)) await toggle.click();
  await expect(toggle).toHaveAttribute('aria-checked', String(auto));
  const first = menu.getByRole('menuitem').nth(1); // the long date
  const text = (await first.textContent()).trim();
  await first.click();
  return text;
}

// ------------------------------------------------------------------ dates

test('a date that updates is Word’s DATE field in the format chosen, with today’s date shown', async ({ page }) => {
  await page.addInitScript(() => { Object.defineProperty(navigator, 'language', { get: () => 'en-US' }); });
  await openInEditor(page, await buildDocx(BODY));
  const text = await insertDate(page, { auto: true });
  await expect(status(page)).toHaveText(`Date inserted: ${text} – Word updates it to the day it opens the file`);
  await expect(page.locator('.docx-page > p').first()).toHaveText(`Invoice sent on ${text}`);
  await expect(page.locator('.docx-page .field[data-instr^="DATE"]')).toHaveText(text);
  const { xml } = await download(page);
  expect(xml).toMatch(/<w:fldChar w:fldCharType="begin"\/><\/w:r><w:r><w:instrText xml:space="preserve"> DATE \\@ "MMMM d, yyyy" <\/w:instrText><\/w:r><w:r><w:fldChar w:fldCharType="separate"\/><\/w:r><w:r><w:t( xml:space="preserve")?>[A-Z][a-z]+ \d{1,2}, \d{4}<\/w:t><\/w:r><w:r><w:fldChar w:fldCharType="end"\/>/);

  // The choice is remembered: the next date updates too, until it is turned off.
  await page.keyboard.press('End');
  await page.getByRole('button', { name: 'Insert a check box or date' }).click();
  await expect(page.getByRole('menuitemcheckbox', { name: /Update automatically/ })).toHaveAttribute('aria-checked', 'true');
});

test('a date that doesn’t update is typed as text, as before', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  const text = await insertDate(page, { auto: false });
  await expect(status(page)).toHaveText(`Date inserted: ${text}`);
  const { xml } = await download(page);
  expect(xml).not.toContain('DATE');
  expect(xml).toContain(text);
});

test('a date that updates, typed while tracking, is a tracked insertion; Undo takes it out', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await page.getByRole('button', { name: 'Track changes' }).click();
  const dialog = page.getByRole('dialog', { name: 'Your name on tracked changes' });
  await dialog.getByRole('textbox').fill('Rita');
  await dialog.getByRole('button', { name: 'Start tracking' }).click();
  await insertDate(page, { auto: true });
  const { xml } = await download(page);
  expect(xml).toMatch(/<w:ins [^>]*w:author="Rita"[^>]*>(?:(?!<\/w:ins>).)*DATE \\@/s);

  await page.locator('.docx-page > p').first().click();
  await page.keyboard.press('ControlOrMeta+z');
  await expect(page.locator('.docx-page .field[data-instr^="DATE"]')).toHaveCount(0);
});

// ------------------------------------------------------------------ page numbers

const part = (root, inner) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${root} ${W_NS}>${inner}</w:${root}>`;
const field = (instr, result) => `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${result}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
const FOOTER = part('ftr', `<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r>${field('PAGE', '1')}<w:r><w:t xml:space="preserve"> of </w:t></w:r>${field('NUMPAGES', '1')}</w:p>`);
const LONG = Array.from({ length: 60 }, (_, i) => `<w:p><w:r><w:t>Line ${i + 1} of a long quote, to fill more than one page of paper.</w:t></w:r></w:p>`).join('');
const OPTS = {
  parts: { 'word/footer1.xml': FOOTER },
  rels: `<Relationship Id="rIdF1" Type="${REL}/footer" Target="footer1.xml"/>`,
  sectPr: '<w:footerReference w:type="default" r:id="rIdF1"/>',
};

test('a page number shows #, and the page count about how many pages; the file keeps its numbers', async ({ page }) => {
  await page.setViewportSize({ width: 1280, height: 800 });
  await openInEditor(page, await buildDocx(LONG, OPTS));
  const pageNo = page.locator('.docx-hf-band.footer .field[data-instr^="PAGE"]');
  const count = page.locator('.docx-hf-band.footer .field[data-instr^="NUMPAGES"]');
  await expect(pageNo).toHaveAttribute('data-show', '#');
  await expect(pageNo).toHaveAttribute('contenteditable', 'false');
  const pages = page.locator('.statusbar .page-count');
  await expect(pages).toHaveText(/≈ \d+ pages/);
  const n = (await pages.textContent()).match(/\d+/)[0];
  expect(Number(n)).toBeGreaterThan(1);
  await expect(count).toHaveAttribute('data-show', n);
  // What shows is the # and the count, drawn over the numbers the file holds.
  expect(await pageNo.evaluate((el) => getComputedStyle(el, '::after').content)).toBe('"#"');
  expect(await count.evaluate((el) => getComputedStyle(el, '::after').content)).toBe(`"${n}"`);

  // An edit elsewhere in the footer leaves the numbers in the file as they were.
  await page.locator('.docx-hf-band.footer').getByText('Page').first().click();
  await page.keyboard.press('Home');
  await page.keyboard.type('Acme · ');
  const { zip } = await download(page);
  const ftr = await zip.file('word/footer1.xml').async('string');
  expect(ftr).toContain('Acme · Page');
  expect(ftr).toMatch(/ PAGE <\/w:instrText><\/w:r><w:r><w:fldChar w:fldCharType="separate"\/><\/w:r><w:r><w:t( xml:space="preserve")?>1<\/w:t>/);
  expect(ftr).toMatch(/ NUMPAGES <\/w:instrText><\/w:r><w:r><w:fldChar w:fldCharType="separate"\/><\/w:r><w:r><w:t( xml:space="preserve")?>1<\/w:t>/);
});

test('a page number put in here shows # too', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await page.locator('.docx-page > p').first().click();
  await page.getByRole('button', { name: 'Page number' }).click();
  await expect(page.locator('.docx-hf-band.footer .field.page-field')).toHaveAttribute('data-show', '#');
});

// ------------------------------------------------------------------ toolbar labels

const labelled = (page) => page.locator('.toolbar button.with-label .short-label').allTextContents();

test('icon buttons get text labels where there is room, without the toolbar growing', async ({ page }) => {
  await page.setViewportSize({ width: 1920, height: 800 });
  await openInEditor(page, await buildDocx(BODY));
  await expect.poll(() => labelled(page)).toEqual(['Insert', 'Symbol', 'Picture', 'Sign', 'Table', 'Layout', 'Footnote', 'Endnote', 'Update', 'Page no.', 'Page setup', 'Find', 'Fill in']);
  // Its name stays what it was, with the label shown or not.
  await expect(page.getByRole('button', { name: 'Insert picture' })).toBeVisible();
  await expect(page.getByRole('button', { name: 'Find and replace' })).toBeVisible();

  for (const width of [1280, 1440]) {
    await page.setViewportSize({ width, height: 800 });
    // The labels are fitted again once the toolbar has its new width (a ResizeObserver, a moment
    // after the resize): wait for that, then check that as many fit as leave it as tall as without.
    const fit = () => page.locator('.toolbar').evaluate((t) => {
      const shown = [...t.querySelectorAll('.with-label')];
      const withLabels = t.offsetHeight;
      shown.forEach((b) => b.classList.remove('with-label'));
      const without = t.offsetHeight;
      shown.forEach((b) => b.classList.add('with-label'));
      return { count: shown.length, grows: withLabels > without };
    });
    await expect.poll(fit).toEqual({ count: expect.any(Number), grows: false });
    expect((await fit()).count).toBeGreaterThan(0);
  }
});

test('on a phone the toolbar keeps its icons alone', async ({ page }) => {
  await page.setViewportSize({ width: 390, height: 800 });
  await openInEditor(page, await buildDocx(BODY));
  await expect(page.locator('.toolbar')).toBeVisible();
  expect(await labelled(page)).toEqual([]);
});

// ------------------------------------------------------------------ the caret after Download

test('after Download the caret is back where it was, in a table cell too', async ({ page }) => {
  const table = '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr>'
    + '<w:tc><w:p><w:r><w:t>Item</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Price</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>';
  await openInEditor(page, await buildDocx(table));
  await page.getByText('Price').click();
  await page.keyboard.press('End');
  await page.keyboard.type('s');
  await download(page);
  await expect(page.getByRole('button', { name: 'Table layout' })).toBeEnabled();
  await page.keyboard.type(' (USD)');
  await expect(page.locator('.docx-page td').nth(1)).toHaveText('Prices (USD)');
});
