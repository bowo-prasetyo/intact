// The ribbon, as Word's: the toolbar's tools in tabs (Home, Insert, Layout, References, Review), one
// tab's at a time, with Undo, Redo and Help in every tab; "All tools" shows them all in one toolbar
// instead (as the other tests use it), and a phone keeps its one row of every tool. The tab and the
// choice are remembered in this browser.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor } from './helpers.js';

// This file tests the ribbon itself: start as a new visitor does, with nothing remembered.
test.use({ storageState: { cookies: [], origins: [] } });

const TABLE = '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
const TRACKED = '<w:p><w:ins w:id="1" w:author="Ana" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Added</w:t></w:r></w:ins></w:p>';
const BODY = '<w:p><w:r><w:t>Text</w:t></w:r></w:p>' + TABLE + TRACKED;
const tabs = (page) => page.getByRole('tablist', { name: 'Toolbar tabs' });
const tab = (page, name) => tabs(page).getByRole('tab', { name: new RegExp(`^${name}`) });
const button = (page, name) => page.getByRole('button', { name, exact: true });

/** The names of the toolbar's controls that show. */
const shown = (page) => page.locator('.toolbar').evaluate((t) => [...t.querySelectorAll('button, select')]
  .filter((el) => el.getClientRects().length)
  .map((el) => el.getAttribute('aria-label') || el.getAttribute('title') || el.textContent.trim()));

test('the tools are in tabs: Home first, one tab’s at a time, Undo, Redo and Help in each', async ({ page }) => {
  await page.setViewportSize({ width: 1280, height: 800 });
  await openInEditor(page, await buildDocx(BODY));
  await expect(tabs(page).getByRole('tab')).toHaveText(['Home', 'Insert', 'Layout', 'References', /^Review/]);
  await expect(tab(page, 'Home')).toHaveAttribute('aria-selected', 'true');
  await expect(page.getByRole('toolbar', { name: 'Home' })).toBeVisible();
  await expect(button(page, 'Bold')).toBeVisible();
  await expect(button(page, 'Insert picture')).toBeHidden();
  // Home's tools take one row on a laptop's screen.
  expect(await page.locator('.toolbar').evaluate((t) => t.offsetHeight)).toBeLessThan(60);

  await tab(page, 'Insert').click();
  await expect(button(page, 'Insert picture')).toBeVisible();
  await expect(button(page, 'Bold')).toBeHidden();
  for (const always of ['Undo', 'Redo', 'Help']) await expect(button(page, always)).toBeVisible();
  await tab(page, 'Layout').click();
  await expect(button(page, 'Page setup')).toBeVisible();
  await tab(page, 'References').click();
  await expect(button(page, 'Insert footnote')).toBeVisible();
  await tab(page, 'Review').click();
  await expect(page.getByRole('button', { name: 'Track changes' })).toBeVisible();

  // The tab is remembered for the next document.
  await openInEditor(page, await buildDocx(BODY));
  await expect(tab(page, 'Review')).toHaveAttribute('aria-selected', 'true');
});

test('every tool of the full toolbar is in one of the tabs', async ({ page }) => {
  await page.setViewportSize({ width: 1600, height: 800 });
  await openInEditor(page, await buildDocx(BODY));
  const inTabs = new Set();
  for (const name of ['Home', 'Insert', 'Layout', 'References', 'Review']) {
    await tab(page, name).click();
    for (const n of await shown(page)) inTabs.add(n);
  }
  await page.getByRole('button', { name: 'All tools' }).click();
  await expect(tabs(page)).toHaveCount(0);
  const all = (await shown(page)).filter((n) => n !== 'Show the tools in tabs (Home, Insert, Layout, References, Review), as Word’s ribbon');
  expect(all.length).toBeGreaterThan(40);
  expect(all.filter((n) => !inTabs.has(n))).toEqual([]);
});

test('“All tools” shows them in one toolbar, remembered; “Tabs” goes back to the ribbon', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await page.getByRole('button', { name: 'All tools' }).click();
  await expect(page.getByRole('toolbar', { name: 'Editing and review' })).toBeVisible();
  await expect(button(page, 'Bold')).toBeVisible();
  await expect(button(page, 'Insert picture')).toBeVisible();
  await expect(page.getByRole('button', { name: 'Track changes' })).toBeVisible();
  await openInEditor(page, await buildDocx(BODY));
  await expect(tabs(page)).toHaveCount(0);
  await page.getByRole('button', { name: 'Tabs', exact: true }).click();
  await expect(tab(page, 'Home')).toBeVisible();
  await expect(button(page, 'Insert picture')).toBeHidden();
});

test('the tabs from the keyboard, and what they tell: tracked changes waiting, the caret in a table', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  // The Review tab counts the tracked changes; the Layout tab is marked in a table, as Word shows
  // its table tabs.
  await expect(tab(page, 'Review')).toContainText('1');
  await expect(tab(page, 'Layout')).not.toHaveClass(/context/);
  await page.getByText('Cell').click();
  await expect(tab(page, 'Layout')).toHaveClass(/context/);

  await tab(page, 'Home').focus();
  await page.keyboard.press('ArrowRight');
  await expect(tab(page, 'Insert')).toBeFocused();
  await expect(tab(page, 'Insert')).toHaveAttribute('aria-selected', 'true');
  await page.keyboard.press('End');
  await expect(tab(page, 'Review')).toHaveAttribute('aria-selected', 'true');
  await page.keyboard.press('ArrowRight'); // round to the first
  await expect(tab(page, 'Home')).toHaveAttribute('aria-selected', 'true');
  // Alt+F10 goes to the chosen tab's tools; the arrows move along those that show.
  await page.getByText('Text').click();
  await page.keyboard.press('Alt+F10');
  await expect(button(page, 'Undo')).toBeFocused();
  await page.keyboard.press('ArrowLeft');
  await expect(button(page, 'Help')).toBeFocused();
});

test('on a phone, one row of every tool and no tabs', async ({ page }) => {
  await page.setViewportSize({ width: 390, height: 800 });
  await openInEditor(page, await buildDocx(BODY));
  await expect(tabs(page)).toBeHidden();
  const names = await page.locator('.toolbar').evaluate((t) => [...t.querySelectorAll('button')]
    .filter((b) => getComputedStyle(b).display !== 'none' && !b.closest('[style*="display: none"]'))
    .map((b) => b.getAttribute('aria-label') || b.getAttribute('title')));
  expect(names).toEqual(expect.arrayContaining(['Bold', 'Insert picture', 'Page setup', 'Insert footnote', 'Help']));
});
