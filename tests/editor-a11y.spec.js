// The editor's own interface checked for accessibility with axe-core (WCAG 2.1 A and AA): the start
// page, the About page, the editor at a laptop's and a wide screen's width, and the dialogs and menus
// used most. The document on the page is left out: its colours and structure are its author's, and
// the document's own accessibility check (Review) looks at those. A finding fails the test, with
// what and where. (A pass here is automated; a screen reader's pass by hand is still to do.)
import { test, expect } from '@playwright/test';
import AxeBuilder from '@axe-core/playwright';
import { buildDocx, openInEditor } from './helpers.js';

async function findings(page) {
  const list = (r) => r.violations.map((v) => `${v.id} (${v.impact}): ${v.nodes.map((n) => n.target.join(' ')).join(', ')}`);
  // The interface, without the document; and, with the document (which can be focused and so scrolls
  // from the keyboard), whether its scrolling area can be reached: left out, it would look empty.
  const ui = await new AxeBuilder({ page }).withTags(['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa']).disableRules(['scrollable-region-focusable']).exclude('.docx-page').analyze();
  const scrolling = await new AxeBuilder({ page }).withRules(['scrollable-region-focusable']).analyze();
  return [...list(ui), ...list(scrolling)];
}

const DOC = '<w:p><w:r><w:t>Hello world, a short document.</w:t></w:r></w:p>';

test('the start page and the About page', async ({ page }) => {
  await page.goto('./');
  expect(await findings(page)).toEqual([]);
  await page.goto('./#/about');
  await expect(page.getByRole('heading', { name: 'About' })).toBeVisible();
  expect(await findings(page)).toEqual([]);
});

for (const width of [1366, 1920]) {
  test(`the editor at ${width} pixels wide, its toolbar's labels shown where they fit`, async ({ page }) => {
    await page.setViewportSize({ width, height: 900 });
    await openInEditor(page, await buildDocx(DOC));
    expect(await findings(page)).toEqual([]);
  });
}

test('Page setup, the right-click menu, Word count and Find', async ({ page }) => {
  await page.setViewportSize({ width: 1366, height: 900 });
  await openInEditor(page, await buildDocx(DOC));
  await page.locator('.docx-page p').first().click();
  await page.getByRole('button', { name: /Page setup/ }).first().click();
  await expect(page.getByRole('dialog', { name: /Page setup/ })).toBeVisible();
  expect(await findings(page)).toEqual([]);
  await page.keyboard.press('Escape');
  await page.locator('.docx-page p').first().click({ button: 'right' });
  await expect(page.getByRole('menu').first()).toBeVisible();
  expect(await findings(page)).toEqual([]);
  await page.keyboard.press('Escape');
  await page.getByRole('button', { name: 'Word count, length and readability' }).click();
  const count = page.getByRole('dialog', { name: 'Word count' });
  await expect(count).toBeVisible();
  expect(await findings(page)).toEqual([]);
  await count.getByRole('button', { name: 'OK' }).click();
  await page.locator('.docx-page p').first().click();
  await page.keyboard.press('ControlOrMeta+f');
  await expect(page.getByRole('textbox', { name: /Find/ }).first()).toBeVisible();
  expect(await findings(page)).toEqual([]);
});

test('Review your changes, before downloading', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await page.locator('.docx-page p').first().click();
  await page.keyboard.type(' Edited.');
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  await expect(page.getByRole('dialog', { name: 'Review your changes' })).toBeVisible();
  expect(await findings(page)).toEqual([]);
});

test('Fill in, as a template opens with it', async ({ page }) => {
  await page.goto('./');
  await page.getByRole('button', { name: /^Letter:/ }).click();
  await expect(page.getByRole('dialog', { name: /fill in/i })).toBeVisible();
  expect(await findings(page)).toEqual([]);
});
