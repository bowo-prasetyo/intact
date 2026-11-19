// Opening a document from the start page, by keyboard as well as by mouse, and dropping files: a file
// dropped anywhere on the start page opens; one dropped where nothing takes it is refused, so the
// browser never leaves the app to open or download it (with an open document's edits lost).
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor } from './helpers.js';

const BODY = '<w:p><w:r><w:t>Dropped document</w:t></w:r></w:p>';
const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

/** Drags a file over the element `selector` and drops it there; what the page let the browser do. */
function dropFile(page, selector, { name, type, bytes }) {
  return page.evaluate(([sel, name, type, data]) => {
    const dt = new DataTransfer();
    dt.items.add(new File([new Uint8Array(data)], name, { type }));
    // A DataTransfer made by a script ignores its drop effect being set: record what the page sets.
    let dropEffect = 'copy';
    Object.defineProperty(dt, 'dropEffect', { get: () => dropEffect, set: (v) => { dropEffect = v; } });
    const target = document.querySelector(sel);
    const fire = (kind) => {
      const e = new DragEvent(kind, { bubbles: true, cancelable: true, dataTransfer: dt });
      target.dispatchEvent(e);
      return e;
    };
    fire('dragenter');
    const over = fire('dragover');
    const effect = dropEffect;
    const drop = fire('drop');
    // A dragover cancelled with the drop effect "none" refuses the drop; cancelled otherwise, allows it.
    return { dropAllowed: over.defaultPrevented && effect !== 'none', browserDefault: !drop.defaultPrevented };
  }, [selector, name, type, [...bytes]]);
}

test('the Open tile is a button: Tab reaches it and Enter opens the file picker', async ({ page }) => {
  // Listening before the page loads, so the picker can't open before Playwright is watching for it.
  const picker = page.waitForEvent('filechooser');
  await page.goto('./');
  const open = page.getByRole('button', { name: 'Open a .docx file' });
  await expect(open).toHaveAccessibleDescription(/drop a file anywhere here/);
  await page.getByRole('button', { name: 'Blank document' }).focus();
  await page.keyboard.press('Tab');
  await expect(open).toBeFocused();
  await page.keyboard.press('Enter');
  const chooser = await picker;
  await chooser.setFiles({ name: 'keyboard.docx', mimeType: DOCX, buffer: await buildDocx(BODY) });
  await expect(page.locator('.docx-page')).toContainText('Dropped document');
});

test('a file dropped anywhere on the start page opens, not only on the Open tile', async ({ page }) => {
  await page.goto('./');
  const bytes = await buildDocx(BODY);
  const result = await dropFile(page, 'h1', { name: 'dropped.docx', type: DOCX, bytes });
  expect(result).toEqual({ dropAllowed: true, browserDefault: false });
  await expect(page.locator('.docx-page')).toContainText('Dropped document');
  await expect(page.locator('.app-header .filename')).toHaveValue('dropped');
});

test('dragging a file over the start page shows where it goes', async ({ page }) => {
  await page.goto('./');
  await page.evaluate(() => {
    const dt = new DataTransfer();
    dt.items.add(new File(['x'], 'a.docx'));
    document.querySelector('h1').dispatchEvent(new DragEvent('dragenter', { bubbles: true, cancelable: true, dataTransfer: dt }));
  });
  await expect(page.getByRole('button', { name: 'Drop to open' })).toHaveClass(/\bover\b/);
});

test('in the editor, a file dropped outside the page is refused and one the page can’t take is explained', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  const bytes = await buildDocx(BODY);
  // Outside the page: no drop at all, so the browser doesn't replace the app with the file.
  for (const where of ['.toolbar', '.statusbar', '.canvas']) {
    expect(await dropFile(page, where, { name: 'other.docx', type: DOCX, bytes }), where).toEqual({ dropAllowed: false, browserDefault: false });
  }
  await expect(page).toHaveURL(/#\/edit$/);
  // On the page: pictures go in (see insert.spec.js); a document is offered to open in a new tab
  // (see tabs.spec.js), not here.
  expect((await dropFile(page, '.docx-page p', { name: 'other.docx', type: DOCX, bytes })).browserDefault).toBe(false);
  await expect(page.getByRole('dialog', { name: 'Open in a new tab' })).toContainText('“other.docx” opens in a new tab');
  await page.getByRole('dialog', { name: 'Open in a new tab' }).getByRole('button', { name: 'Cancel' }).click();
  expect((await dropFile(page, '.docx-page p', { name: 'notes.txt', type: 'text/plain', bytes: Buffer.from('hi') })).browserDefault).toBe(false);
  await expect(page.locator('.statusbar [role=status]')).toHaveText('Only pictures can be dropped into the document');
  await expect(page.locator('.docx-page')).toHaveText('Dropped document');
});
