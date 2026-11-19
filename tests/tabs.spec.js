// Several documents at once, each in a tab of its own (src/tabs.js): the Open button and Ctrl+O open
// the start page in a new tab; a document dropped on the page is offered to open in a new tab and is
// handed over in memory; a draft being edited in one tab shows in another tab's start page as open
// there, not to be restored a second time.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor } from './helpers.js';

const para = (text) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

async function dropDocument(page, name, buffer) {
  const dt = await page.evaluateHandle(({ name, bytes }) => {
    const dt = new DataTransfer();
    dt.items.add(new File([new Uint8Array(bytes)], name, { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }));
    return dt;
  }, { name, bytes: [...buffer] });
  await page.locator('.docx-page p').first().dispatchEvent('drop', { dataTransfer: dt });
}

test('the Open button opens the start page in a new tab; this document stays open', async ({ page, context }) => {
  await openInEditor(page, await buildDocx(para('The CV')), 'CV.docx');
  const [tab] = await Promise.all([context.waitForEvent('page'), page.getByRole('button', { name: 'Open another document in a new tab' }).click()]);
  await expect(tab.getByRole('button', { name: /Open a \.docx file/ })).toBeVisible();
  await expect(page.locator('.docx-page')).toHaveText('The CV');
});

test('Ctrl+O, as Word\'s Open, opens another document in a new tab', async ({ page, context }) => {
  await openInEditor(page, await buildDocx(para('The CV')));
  await page.locator('.docx-page p').click();
  const [tab] = await Promise.all([context.waitForEvent('page'), page.keyboard.press('ControlOrMeta+o')]);
  await expect(tab.getByRole('button', { name: /Open a \.docx file/ })).toBeVisible();
  await expect(page).toHaveURL(/#\/edit$/);
});

test('a document dropped on the page opens in a new tab, beside this one', async ({ page, context }) => {
  await openInEditor(page, await buildDocx(para('The CV')), 'CV.docx');
  await dropDocument(page, 'Cover_letter.docx', await buildDocx(para('Dear hiring manager')));
  const offer = page.getByRole('dialog', { name: 'Open in a new tab' });
  await expect(offer).toContainText('“Cover_letter.docx” opens in a new tab, beside this document');
  const [tab] = await Promise.all([context.waitForEvent('page'), offer.getByRole('button', { name: 'Open in a new tab' }).click()]);
  await expect(tab.locator('.docx-page')).toHaveText('Dear hiring manager');
  await expect(tab).toHaveTitle('Cover_letter.docx – Intact');
  // The token in the new tab's address isn't kept: a reload doesn't ask for the file again.
  await expect(tab).toHaveURL(/#\/edit$/);
  await expect(page.locator('.docx-page')).toHaveText('The CV');
  await expect(page.locator('.statusbar [role=status]')).toHaveText('“Cover_letter.docx” is opening in a new tab');
});

test('a dropped document can be left unopened', async ({ page }) => {
  await openInEditor(page, await buildDocx(para('The CV')));
  await dropDocument(page, 'Other.docx', await buildDocx(para('Other')));
  await page.getByRole('dialog', { name: 'Open in a new tab' }).getByRole('button', { name: 'Cancel' }).click();
  await expect(page.getByRole('dialog', { name: 'Open in a new tab' })).toBeHidden();
  await expect(page.locator('.docx-page')).toHaveText('The CV');
});

test('a tab asked to open a file that no tab hands over says so', async ({ page }) => {
  await page.goto('./#/?receive=nobody-has-this');
  await expect(page.getByRole('alert')).toContainText('The document couldn’t be brought over from the other tab', { timeout: 10_000 });
  await expect(page).toHaveURL(/#\/$/);
});

test('a draft open in another tab is shown as open there, not offered to restore', async ({ page, context, browserName }) => {
  test.skip(browserName === 'firefox' && !(await page.evaluate(() => !!navigator.locks)), 'Web Locks');
  await openInEditor(page, await buildDocx(para('The CV')), 'CV.docx');
  await page.locator('.docx-page p').click();
  await page.keyboard.press('Home');
  await page.keyboard.type('Edited: ');
  await expect(page.locator('.draft-status')).toBeVisible(); // the draft is kept
  const [tab] = await Promise.all([context.waitForEvent('page'), page.getByRole('button', { name: 'Open another document in a new tab' }).click()]);
  const row = tab.locator('.drafts li', { hasText: 'CV.docx' });
  await expect(row).toContainText('Open in another tab');
  await expect(row.getByRole('button', { name: /Restore/ })).toHaveCount(0);
  // Closed there, it can be restored here.
  await page.close();
  await tab.evaluate(() => window.dispatchEvent(new Event('focus')));
  await expect(row.getByRole('button', { name: 'Restore CV.docx' })).toBeVisible();
  await row.getByRole('button', { name: 'Restore CV.docx' }).click();
  await expect(tab.locator('.docx-page')).toHaveText('Edited: The CV');
});

test('on a narrow phone the header, Open button and all, never scrolls the page sideways', async ({ page }) => {
  for (const width of [390, 340, 320]) {
    await page.setViewportSize({ width, height: 664 });
    await openInEditor(page, await buildDocx(para('The CV')), 'Alex_Morgan_Curriculum_Vitae_2026.docx');
    await expect(page.getByRole('button', { name: 'Open another document in a new tab' })).toBeInViewport();
    await expect(page.getByRole('button', { name: 'Close' })).toBeInViewport();
    expect(await page.evaluate(() => document.documentElement.scrollWidth), `at ${width}px`).toBeLessThanOrEqual(width);
  }
});
