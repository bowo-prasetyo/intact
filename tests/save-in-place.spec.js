// Saving back to the file opened, where the browser can (the File System Access API, in Chrome and
// Edge on a computer): opened through the browser's file picker, Save writes into that file instead
// of downloading a copy; Download a copy and Save as… are in Download's menu. The browser's own
// pickers can't be driven by a test, so they are stood in for by ones that hand over a file the
// test made and keep what is written to it.
import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import { buildDocx } from './helpers.js';

const status = (page) => page.locator('.statusbar [role=status]');
const header = (page) => page.locator('.app-header');
const review = (page) => page.getByRole('dialog', { name: 'Review your changes' });

/** Stands in for the browser's pickers: `window.__open` is the file to hand over, `window.__written` what Save wrote, by name. */
async function fakePickers(page, { permission = 'granted' } = {}) {
  await page.addInitScript((perm) => {
    window.__written = {};
    window.__asked = 0;
    window.__fail = null; // 'write' or 'close': the next write fails there, as with a full disk
    window.__aborted = 0;
    const handle = (name, bytes) => ({
      kind: 'file',
      name,
      getFile: async () => new File([bytes || new Uint8Array()], name),
      queryPermission: async () => 'prompt',
      requestPermission: async () => { window.__asked += 1; return perm; },
      createWritable: async () => {
        // As the browser's: written to a copy, which takes the file's place only when closed whole.
        const parts = [];
        return {
          write: async (blob) => {
            if (window.__fail === 'write') throw new DOMException('The disk is full', 'QuotaExceededError');
            parts.push(new Uint8Array(await blob.arrayBuffer()));
          },
          close: async () => {
            if (window.__fail === 'close') throw new DOMException('The file could not be written', 'InvalidStateError');
            window.__written[name] = parts.flatMap((p) => [...p]);
          },
          abort: async () => { window.__aborted += 1; },
        };
      },
    });
    window.showOpenFilePicker = async () => [handle(window.__open.name, new Uint8Array(window.__open.bytes))];
    window.showSaveFilePicker = async ({ suggestedName }) => handle(window.__saveAs || suggestedName);
  }, permission);
}

async function openPicked(page, name, buffer) {
  await page.goto('./');
  await page.evaluate(([n, bytes]) => { window.__open = { name: n, bytes }; }, [name, [...buffer]]);
  await page.getByRole('button', { name: /Open a \.docx file/ }).click();
  await expect(page.locator('.docx-page')).toBeVisible();
}

async function written(page, name) {
  const bytes = await page.evaluate((n) => window.__written[n] || null, name);
  if (!bytes) return null;
  return (await JSZip.loadAsync(Buffer.from(bytes))).file('word/document.xml').async('string');
}

const BODY = '<w:p><w:r><w:t>Chapter one draft.</w:t></w:r></w:p>';

test('opened with the browser’s picker, Save writes the changes into the same file: no download', async ({ page }) => {
  await fakePickers(page);
  await openPicked(page, 'novel.docx', await buildDocx(BODY));
  const save = header(page).getByRole('button', { name: 'Save', exact: true });
  await expect(save).toBeVisible();
  await page.locator('.docx-page p').first().click();
  await page.keyboard.press('End');
  await page.keyboard.type(' Revised.');
  let downloaded = false;
  page.on('download', () => { downloaded = true; });
  await page.keyboard.press('Control+s');
  await review(page).getByRole('button', { name: 'Save', exact: true }).click();
  await expect(status(page)).toContainText('Saved novel.docx on this device');
  expect(await written(page, 'novel.docx')).toContain('Chapter one draft. Revised.');
  expect(await page.evaluate(() => window.__asked)).toBe(1); // the browser asked once, "Save changes to novel.docx?"
  expect(downloaded).toBe(false);

  // Saved again: into the same file.
  await page.keyboard.type(' Again.');
  await save.click();
  await review(page).getByRole('button', { name: 'Save', exact: true }).click();
  await expect.poll(() => written(page, 'novel.docx')).toContain('Revised. Again.');
});

test('Enter on the Open tile opens the browser’s picker, from the keyboard', async ({ page }) => {
  await fakePickers(page);
  await page.goto('./');
  await page.evaluate(async (bytes) => { window.__open = { name: 'keys.docx', bytes }; }, [...await buildDocx(BODY)]);
  await page.getByRole('button', { name: 'Blank document' }).focus();
  await page.keyboard.press('Tab');
  await page.keyboard.press('Enter');
  await expect(page.locator('.docx-page')).toContainText('Chapter one draft.');
  await expect(header(page).getByRole('button', { name: 'Save', exact: true })).toBeVisible();
});

test('Download a copy still downloads, and the file opened stays as it was', async ({ page }) => {
  await fakePickers(page);
  await openPicked(page, 'novel.docx', await buildDocx(BODY));
  await page.locator('.docx-page p').first().click();
  await page.keyboard.type('New ');
  await header(page).getByRole('button', { name: 'More ways to save' }).click();
  await page.getByRole('menu', { name: 'Save' }).getByRole('menuitem', { name: /Download a copy/ }).click();
  const [dl] = await Promise.all([page.waitForEvent('download'), review(page).getByRole('button', { name: 'Download', exact: true }).click()]);
  expect(dl.suggestedFilename()).toBe('novel.docx');
  expect(await written(page, 'novel.docx')).toBeNull();
});

test('when the browser isn’t allowed to write, nothing is lost: it says so, and the document stays unsaved', async ({ page }) => {
  await fakePickers(page, { permission: 'denied' });
  await openPicked(page, 'novel.docx', await buildDocx(BODY));
  await page.locator('.docx-page p').first().click();
  await page.keyboard.type('New ');
  await page.keyboard.press('Control+s');
  await review(page).getByRole('button', { name: 'Save', exact: true }).click();
  await expect(status(page)).toContainText('Not saved: the browser didn’t allow writing to novel.docx – Download a copy (▾) keeps one in Downloads');
  expect(await written(page, 'novel.docx')).toBeNull();
});

test('a new manuscript: Save as… picks a file, and later Saves write to it', async ({ page }) => {
  await fakePickers(page);
  await page.goto('./');
  await page.evaluate(() => { window.__saveAs = 'My Novel.docx'; });
  await page.getByRole('button', { name: /^Manuscript:/ }).click();
  await expect(page.locator('.docx-page')).toBeVisible();
  await page.keyboard.press('Escape'); // Fill in, later
  await expect(header(page).getByRole('button', { name: 'Download', exact: true })).toBeVisible();
  await header(page).getByRole('button', { name: 'More ways to save' }).click();
  await page.getByRole('menu', { name: 'Save' }).getByRole('menuitem', { name: /Save as…/ }).click();
  await review(page).getByRole('button', { name: 'Save as…' }).click();
  await expect(status(page)).toContainText('Saved My Novel.docx on this device');
  expect(await written(page, 'My Novel.docx')).toContain('Chapter One');
  await expect(page.getByRole('textbox', { name: 'File name' })).toHaveValue('My Novel');
  // From now on, Save.
  await expect(header(page).getByRole('button', { name: 'Save', exact: true })).toBeVisible();
});

test('without the browser’s pickers (Firefox, Safari, phones), Open and Download work as before', async ({ page }) => {
  await page.addInitScript(() => { delete window.showOpenFilePicker; delete window.showSaveFilePicker; });
  await page.goto('./');
  await page.setInputFiles('input[type=file]', { name: 'plain.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await buildDocx(BODY) });
  await expect(page.locator('.docx-page')).toBeVisible();
  await expect(header(page).getByRole('button', { name: 'Download', exact: true })).toBeVisible();
  await header(page).getByRole('button', { name: 'More ways to save' }).click();
  await expect(page.getByRole('menu', { name: 'Save' }).getByRole('menuitem', { name: /Save as…|Download a copy/ })).toHaveCount(0);
});

for (const where of ['write', 'close']) {
  test(`a save that fails part way (at ${where}) leaves the file as it was, and the changes still here`, async ({ page }) => {
    await fakePickers(page);
    await openPicked(page, 'novel.docx', await buildDocx(BODY));
    // A first Save writes the file, so there is a version to keep.
    await page.locator('.docx-page p').first().click();
    await page.keyboard.press('End');
    await page.keyboard.type(' Saved.');
    await header(page).getByRole('button', { name: 'Save', exact: true }).click();
    await review(page).getByRole('button', { name: 'Save', exact: true }).click();
    await expect(status(page)).toContainText('Saved novel.docx on this device');
    expect(await written(page, 'novel.docx')).toContain('Chapter one draft. Saved.');

    // The next Save fails part way: the copy is thrown away, the file stays as it was.
    await page.keyboard.type(' Lost?');
    await expect(page.locator('.statusbar .draft-status')).toHaveText('Changes kept on this device');
    await page.evaluate((w) => { window.__fail = w; }, where);
    await header(page).getByRole('button', { name: 'Save', exact: true }).click();
    await review(page).getByRole('button', { name: 'Save', exact: true }).click();
    await expect(status(page)).toContainText('Not saved:');
    await expect(status(page)).toContainText('novel.docx is as it was, and your changes are still here');
    expect(await page.evaluate(() => window.__aborted)).toBe(1);
    const kept = await written(page, 'novel.docx');
    expect(kept).toContain('Chapter one draft. Saved.');
    expect(kept).not.toContain('Lost?');
    // The changes are still here, still unsaved, and still kept in this browser.
    await expect(page.locator('.docx-page p').first()).toContainText('Saved. Lost?');
    await expect(page.locator('.statusbar .draft-status')).toHaveText('Changes kept on this device');

    // Once writing works again, Save writes them.
    await page.evaluate(() => { window.__fail = null; });
    await header(page).getByRole('button', { name: 'Save', exact: true }).click();
    await review(page).getByRole('button', { name: 'Save', exact: true }).click();
    await expect(status(page)).toContainText('Saved novel.docx on this device');
    expect(await written(page, 'novel.docx')).toContain('Saved. Lost?');
  });
}
