// The toolbar: named controls with their on/off state for screen readers; one Tab stop, arrow
// keys along it, Esc back to the text and Alt+F10 to it (WAI-ARIA's toolbar pattern, as in Word);
// and controls that stay put when the document gets its first tracked change.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor } from './helpers.js';

const BODY = '<w:p><w:r><w:t xml:space="preserve">Plain </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r></w:p>';
const toolbar = (page) => page.getByRole('toolbar', { name: 'Editing and review' });

async function caretIn(page, word, offset = 1) {
  await page.evaluate(([w, o]) => {
    const root = document.querySelector('.docx-page');
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    for (let t = walker.nextNode(); t; t = walker.nextNode()) {
      const at = t.data.indexOf(w);
      if (at < 0) continue;
      root.focus();
      getSelection().setBaseAndExtent(t, at + o, t, at + o);
      return;
    }
  }, [word, offset]);
}

test('toolbar buttons have names, and the on/off ones say whether they are on', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  for (const name of ['Undo', 'Redo', 'Bold', 'Italic', 'Underline', 'Strikethrough', 'Superscript', 'Subscript',
    'Highlight colour options', 'Font colour options', 'Align left', 'Center', 'Align right', 'Justify',
    'Show formatting marks', 'More ways to accept', 'More ways to reject']) {
    await expect(toolbar(page).getByRole('button', { name, exact: true })).toHaveCount(1);
  }
  const bold = toolbar(page).getByRole('button', { name: 'Bold', exact: true });
  await caretIn(page, 'Plain');
  await expect(bold).toHaveAttribute('aria-pressed', 'false');
  await expect(toolbar(page).getByRole('button', { name: 'Align left', exact: true })).toHaveAttribute('aria-pressed', 'true');
  await caretIn(page, 'bold');
  await expect(bold).toHaveAttribute('aria-pressed', 'true');
});

test('the toolbar is one Tab stop: arrows move along it, Esc goes back to the text, Alt+F10 comes back', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await expect(toolbar(page).locator('[tabindex="0"]')).toHaveCount(1);
  await caretIn(page, 'Plain', 3);
  await page.keyboard.press('Alt+F10');
  const undo = toolbar(page).getByRole('button', { name: 'Undo', exact: true });
  await expect(undo).toBeFocused();
  await page.keyboard.press('ArrowRight');
  await expect(toolbar(page).getByRole('button', { name: 'Redo', exact: true })).toBeFocused();
  await page.keyboard.press('ArrowRight');
  const style = toolbar(page).getByRole('combobox', { name: 'Paragraph style' });
  await expect(style).toBeFocused();
  await page.keyboard.press('ArrowRight'); // from a list, ← and → still move along the toolbar
  await expect(toolbar(page).getByRole('combobox', { name: 'Font', exact: true })).toBeFocused();
  await expect(style).toHaveValue('normal');
  await page.keyboard.press('Home');
  await expect(undo).toBeFocused();
  await page.keyboard.press('ArrowLeft'); // wraps around to the last control
  await expect(toolbar(page).getByRole('button', { name: 'Help', exact: true })).toBeFocused();
  await page.keyboard.press('Home');
  // Tab leaves the toolbar at once (for the text, where Tab types a tab, as in Word); Alt+F10 comes
  // back to the control last used.
  await page.keyboard.press('ArrowRight');
  await page.keyboard.press('Tab');
  await expect(page.locator('.toolbar :focus')).toHaveCount(0);
  await caretIn(page, 'Plain', 3);
  await page.keyboard.press('Alt+F10');
  await expect(toolbar(page).getByRole('button', { name: 'Redo', exact: true })).toBeFocused();
  // Esc: back to the text, where the caret was.
  await page.keyboard.press('Escape');
  await expect(page.locator('.docx-page')).toBeFocused();
  await page.keyboard.type('X');
  await expect(page.locator('.docx-page p').first()).toHaveText('PlaXin bold');
});

test('Alt+F10 straight after the caret moves: Esc comes back where it is, not where it was', async ({ page }) => {
  // The editor keeps the caret on selectionchange, which comes later; a key pressed before it came
  // took Esc back to the caret before (WebKit in CI was slow enough to show it).
  await openInEditor(page, await buildDocx(BODY));
  await caretIn(page, 'Plain', 0);
  await expect.poll(() => page.evaluate(() => getSelection().anchorOffset)).toBe(0);
  await page.waitForTimeout(100); // the editor has kept this caret
  await page.evaluate(() => {
    const root = document.querySelector('.docx-page');
    const text = document.createTreeWalker(root, NodeFilter.SHOW_TEXT).nextNode();
    getSelection().setBaseAndExtent(text, 3, text, 3);
    window.dispatchEvent(new KeyboardEvent('keydown', { key: 'F10', altKey: true, bubbles: true, cancelable: true }));
  });
  await expect(toolbar(page).getByRole('button', { name: 'Undo', exact: true })).toBeFocused();
  await page.keyboard.press('Escape');
  await page.keyboard.type('X');
  await expect(page.locator('.docx-page p').first()).toHaveText('PlaXin bold');
});

test('the controls stay where they are when the first tracked change appears, Accept and Reject labelled', async ({ page }) => {
  await page.setViewportSize({ width: 1280, height: 800 });
  await openInEditor(page, await buildDocx(BODY));
  const where = () => page.locator('.toolbar').evaluate((bar) => [...bar.querySelectorAll('button, select')]
    .map((el) => { const r = el.getBoundingClientRect(); return `${el.getAttribute('aria-label') || el.title}@${Math.round(r.x)},${Math.round(r.y)},${Math.round(r.width)}`; }));
  const accept = toolbar(page).getByRole('button', { name: 'Accept', exact: true });
  await expect(accept).toBeVisible();
  await expect(accept).toBeDisabled();
  await expect(page.locator('.changes-group .badge')).toHaveText('0');
  const before = await where();

  await page.getByRole('button', { name: 'Track changes' }).click();
  const dialog = page.getByRole('dialog', { name: 'Your name on tracked changes' });
  await dialog.getByRole('textbox').fill('Rita');
  await dialog.getByRole('button', { name: 'Start tracking' }).click();
  await caretIn(page, 'Plain', 0);
  await page.keyboard.type('New ');
  await expect(page.locator('.changes-group .badge')).toHaveText('1');
  await expect(accept).toBeEnabled();
  await expect(toolbar(page).getByRole('button', { name: 'Reject', exact: true })).toBeVisible();
  await expect(page.getByRole('button', { name: 'Track changes' })).toContainText('Track changes');
  expect(await where()).toEqual(before);
});
