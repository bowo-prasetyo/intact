// Help (Ctrl+/ or the toolbar's Help button): the keyboard shortcuts, and the review's checks
// reachable without downloading, each opening the review at its part.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor } from './helpers.js';

const BODY = '<w:p><w:r><w:t xml:space="preserve">The fee is due.</w:t></w:r></w:p>';
const help = (page) => page.getByRole('dialog', { name: 'Help' });
const review = (page) => page.getByRole('dialog', { name: 'Review your changes' });

test('Ctrl+/ opens Help with the shortcuts; Esc closes it and the caret is back in the text', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await page.locator('.docx-page p').first().click();
  await page.keyboard.press('End');
  await page.keyboard.press('ControlOrMeta+/');
  await expect(help(page)).toBeVisible();
  await expect(help(page).getByRole('button', { name: /^Review your changes/ })).toBeFocused();
  for (const keys of ['Ctrl+B, Ctrl+I, Ctrl+U', 'Ctrl+Shift+E', 'Alt+F10, then Esc', 'Ctrl+/']) {
    await expect(help(page).getByRole('rowheader', { name: keys, exact: true })).toBeVisible();
  }
  await page.keyboard.press('Escape');
  await expect(help(page)).toHaveCount(0);
  await page.keyboard.type(' Today.');
  await expect(page.locator('.docx-page p').first()).toHaveText('The fee is due. Today.');
});

test('Help’s checks open the review at their part, without downloading', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await page.locator('.toolbar').getByRole('button', { name: 'Help', exact: true }).click();
  await help(page).getByRole('button', { name: /^Check accessibility/ }).click();
  await expect(help(page)).toHaveCount(0);
  const a11y = review(page).getByRole('region', { name: 'Accessibility' });
  await expect(a11y.locator('details')).toHaveAttribute('open', '');
  await expect(a11y.locator('summary')).toBeFocused();
  await review(page).getByRole('button', { name: 'Keep editing' }).click();

  await page.keyboard.press('ControlOrMeta+/');
  await help(page).getByRole('button', { name: /^Before you release/ }).click();
  const release = review(page).getByRole('region', { name: 'Before you release' });
  await expect(release.locator('details')).toHaveAttribute('open', '');
  await expect(release.locator('summary')).toBeFocused();
});

test('the review opens with a one-line overview, and folds away what needs no look', async ({ page }) => {
  // Nothing to warn about: the sections are folded behind their counts.
  await openInEditor(page, await buildDocx(BODY));
  await page.locator('.toolbar').getByRole('button', { name: 'Review', exact: true }).click();
  await expect(review(page).locator('.overview')).toHaveText(/to check below|Nothing to check/);
  const release = review(page).getByRole('region', { name: 'Before you release' });
  await expect(release.locator('summary')).toContainText('Nothing hidden');
  await expect(release.locator('details')).not.toHaveAttribute('open', '');
  await review(page).getByRole('button', { name: 'Keep editing' }).click();

  // An untracked edit: a warning, shown, and counted in the overview; the changed paragraph shown.
  await page.locator('.docx-page p').first().click();
  await page.keyboard.press('End');
  await page.keyboard.type(' Now.');
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  await expect(review(page).locator('.overview')).toHaveText(/^\d+ things? to check below\.$/);
  await expect(review(page).locator('.notes .warn').first()).toBeVisible();
  await expect(review(page).locator('.change')).toBeVisible();
  await expect(review(page).locator('.changes-section summary')).toContainText('1');
});

test.describe('on a phone', () => {
  test.use({ viewport: { width: 390, height: 664 }, hasTouch: true });

  test('the toolbar is one row that scrolls, faded where more controls are; nothing scrolls the page sideways', async ({ page }) => {
    await openInEditor(page, await buildDocx(BODY));
    const bar = page.locator('.toolbar');
    const height = await bar.evaluate((el) => el.offsetHeight);
    expect(height).toBeLessThan(60);
    await expect(bar).toHaveAttribute('data-more', 'end');
    await bar.evaluate((el) => { el.scrollLeft = el.scrollWidth; });
    await expect(bar).toHaveAttribute('data-more', 'start');
    // A control reached with the keyboard is scrolled into view.
    await page.locator('.docx-page p').first().click();
    await page.keyboard.press('Alt+F10');
    await page.keyboard.press('End');
    await expect(bar.getByRole('button', { name: 'Help', exact: true })).toBeInViewport();
    expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBeLessThanOrEqual(390);
    expect(await page.locator('.statusbar').evaluate((el) => el.offsetHeight)).toBeLessThan(30);
  });
});
