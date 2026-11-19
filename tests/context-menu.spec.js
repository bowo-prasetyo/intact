// The right-click menu, as Word's: Cut, Copy, Paste and Paste as text, a link, a comment, the
// Paragraph dialog, a tracked change's Accept and Reject, and a table's rows and columns where it is
// opened in a table. Shift+right-click leaves the browser's own menu (with spelling suggestions).
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const TABLE = '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>'
  + '<w:tr><w:tc><w:p><w:r><w:t>Item</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Price</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
const DOC = '<w:p><w:r><w:t xml:space="preserve">Please send the invoice today.</w:t></w:r></w:p>' + TABLE
  + '<w:p><w:ins w:id="9" w:author="Ana" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Added by Ana</w:t></w:r></w:ins></w:p>';

const menu = (page) => page.getByRole('menu', { name: 'Context menu' });
const status = (page) => page.locator('.statusbar [role=status]');

/** Selects `word` in the first paragraph holding it, then right-clicks on it. */
async function rightClickOn(page, word, { select = true, selector = '.docx-page p' } = {}) {
  const box = await page.evaluate(([w, sel, pick]) => {
    const p = [...document.querySelectorAll(sel)].find((e) => e.textContent.includes(w));
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    let t = walker.nextNode();
    while (t && !t.data.includes(w)) t = walker.nextNode();
    const r = document.createRange();
    r.setStart(t, t.data.indexOf(w));
    r.setEnd(t, t.data.indexOf(w) + w.length);
    document.querySelector('.docx-page').focus();
    if (pick) {
      getSelection().removeAllRanges();
      getSelection().addRange(r);
    } else {
      getSelection().collapse(t, t.data.indexOf(w) + 1);
    }
    const b = r.getBoundingClientRect();
    return { x: b.left + b.width / 2, y: b.top + b.height / 2 };
  }, [word, selector, select]);
  await page.mouse.click(box.x, box.y, { button: 'right' });
}

test('right-click on selected text offers Cut, Copy and Paste; Cut and Paste move the text', async ({ page, context, browserName }) => {
  test.skip(browserName !== 'chromium', 'reading the clipboard from a page needs Chromium’s permission here');
  await context.grantPermissions(['clipboard-read', 'clipboard-write']);
  await openInEditor(page, await buildDocx(DOC));
  await rightClickOn(page, 'invoice ');
  await expect(menu(page)).toBeVisible();
  await expect(menu(page).getByRole('menuitem', { name: /^Cut/ })).toBeEnabled();
  await expect(menu(page).getByRole('menuitem', { name: /^Copy/ })).toBeEnabled();
  await expect(menu(page).getByRole('menuitem', { name: /^Paste as text/ })).toBeVisible();
  // Not in a table: no table items.
  await expect(menu(page).getByRole('menuitem', { name: 'Insert row below' })).toHaveCount(0);
  await menu(page).getByRole('menuitem', { name: /^Cut/ }).click();
  await expect(menu(page)).toHaveCount(0);
  await expect(page.locator('.docx-page p').first()).toHaveText('Please send the today.');

  // The right-click puts the caret where it is, as in Word; the Menu key opens the menu there.
  await rightClickOn(page, 'today', { select: false });
  await page.keyboard.press('Escape');
  await expect(menu(page)).toHaveCount(0);
  await page.keyboard.press('End');
  await page.keyboard.press('Shift+F10');
  await menu(page).getByRole('menuitem', { name: /^Paste as text/ }).click();
  await expect(page.locator('.docx-page p').first()).toHaveText('Please send the today.invoice ');
});

test('in a table the menu inserts and deletes rows and columns', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await rightClickOn(page, 'Item', { select: false, selector: '.docx-page td p' });
  await menu(page).getByRole('menuitem', { name: 'Insert row below' }).click();
  await expect(page.locator('.docx-page table tr')).toHaveCount(2);
  await rightClickOn(page, 'Price', { select: false, selector: '.docx-page td p' });
  await menu(page).getByRole('menuitem', { name: 'Insert column to the right' }).click();
  await expect(page.locator('.docx-page table tr').first().locator('td')).toHaveCount(3);
  await rightClickOn(page, 'Price', { select: false, selector: '.docx-page td p' });
  await menu(page).getByRole('menuitem', { name: 'Delete column' }).click();
  await expect(page.locator('.docx-page table tr').first().locator('td')).toHaveCount(2);
  const { xml } = await download(page);
  expect(xml.match(/<w:tr>/g)).toHaveLength(2);
});

test('the menu accepts a tracked change, opens the link dialog where it was, and adds a comment', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await rightClickOn(page, 'Added', { select: false });
  await menu(page).getByRole('menuitem', { name: 'Accept change' }).click();
  await expect(page.locator('.docx-page [data-tc]')).toHaveCount(0);
  await expect(status(page)).toContainText('Accepted');

  await rightClickOn(page, 'invoice');
  const at = await menu(page).boundingBox();
  await menu(page).getByRole('menuitem', { name: /^Link…/ }).click();
  const dialog = page.getByRole('dialog', { name: /link/i });
  await expect(dialog).toBeVisible();
  const box = await dialog.boundingBox();
  expect(Math.abs(box.x - at.x)).toBeLessThan(40); // opened where the menu was, not under the toolbar
  await page.keyboard.press('Escape');

  await rightClickOn(page, 'Please');
  await menu(page).getByRole('menuitem', { name: /^New comment/ }).click();
  await expect(page.locator('.comments textarea, .comments [contenteditable="true"]').first()).toBeVisible();
});

test('the Menu key opens the menu at the caret, the arrow keys move through it, and Shift+right-click leaves the browser’s', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await rightClickOn(page, 'send', { select: false });
  await page.keyboard.press('Escape');
  await page.keyboard.press('Shift+F10');
  await expect(menu(page)).toBeVisible();
  // Nothing selected: Cut and Copy are off, and the arrows skip them.
  await expect(menu(page).getByRole('menuitem', { name: /^Cut/ })).toBeDisabled();
  await expect(menu(page).getByRole('menuitem', { name: /^Paste/ }).first()).toBeFocused();
  await page.keyboard.press('ArrowDown');
  await expect(menu(page).getByRole('menuitem', { name: /^Paste as text/ })).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(menu(page)).toHaveCount(0);

  const prevented = await page.evaluate(() => {
    const p = document.querySelector('.docx-page p');
    const e = new MouseEvent('contextmenu', { bubbles: true, cancelable: true, shiftKey: true, clientX: 10, clientY: 10 });
    p.dispatchEvent(e);
    return e.defaultPrevented;
  });
  expect(prevented).toBe(false);
  await expect(menu(page)).toHaveCount(0);
});
