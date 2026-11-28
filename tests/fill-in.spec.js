// Small-business templates: placeholders such as [Client name] filled in once for every place they
// appear, a template (.dotx) made into a new document (.docx) without changing the template, and
// toolbar popups that always fit in the window.
import fs from 'node:fs';
import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import { buildDocx, openInEditor, download } from './helpers.js';

const MAIN = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';
const DOTX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml';
const DOTM = 'application/vnd.ms-word.template.macroEnabledTemplate.main+xml';
const DOCM = 'application/vnd.ms-word.document.macroEnabled.main+xml';

const p = (inner) => `<w:p>${inner}</w:p>`;
const r = (text, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;

// A quote: [Client name] four times (once split across formatting, once locked), {{date}}, «Amount», and
// brackets that aren't placeholders (a citation, a check box, an editorial note, a reference).
const QUOTE = [
  p(r('Quote for [Client name]')),
  p(r('Dear ') + r('[Client', '<w:b/>') + r(' name], thank you.')),
  p(r('Date: {{date}} · Total: «Amount»')),
  p(r('As agreed [1], tick [x] and [sic] per [Smith 2020].')),
  '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc>' + p(r('Signed for [Client name]')) + '</w:tc></w:tr></w:tbl>',
  '<w:sdt><w:sdtPr><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent>' + p(r('Locked: [Client name]')) + '</w:sdtContent></w:sdt>',
].join('');

async function asTemplate(body, type = DOTX) {
  const zip = await JSZip.loadAsync(await buildDocx(body));
  const ct = await zip.file('[Content_Types].xml').async('string');
  zip.file('[Content_Types].xml', ct.replace(MAIN, type).replace('</Types>', type === DOTM ? '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>' : '</Types>'));
  if (type === DOTM) zip.file('word/vbaProject.bin', Buffer.from('VBA project bytes'));
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function downloaded(page) {
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  const review = page.getByRole('dialog', { name: 'Review your changes' });
  const [dl] = await Promise.all([page.waitForEvent('download'), review.getByRole('button', { name: 'Download', exact: true }).click()]);
  return { name: dl.suggestedFilename(), zip: await JSZip.loadAsync(fs.readFileSync(await dl.path())) };
}

const fillButton = (page) => page.getByRole('button', { name: 'Fill in placeholders' });
const dialog = (page) => page.getByRole('dialog', { name: 'Fill in placeholders' });

// ------------------------------------------------------------------ fill in placeholders

test('placeholders are listed once each, with where they are, and not citations or check boxes', async ({ page }) => {
  await openInEditor(page, await buildDocx(QUOTE));
  // The count is in the button's name (and drawn beside its label).
  await expect(fillButton(page)).toHaveAccessibleName('Fill in placeholders (3)');
  await fillButton(page).click();
  const d = dialog(page);
  await expect(d.locator('label')).toHaveText(['Client name', 'date', 'Amount']);
  // Three in the text and the table can be filled in; the one in the locked field can't.
  await expect(d.getByText('4 places, 1 read-only')).toBeVisible();
  await expect(d.getByLabel('Client name')).toBeFocused();
  // Focusing one highlights every place it is.
  const highlighted = await page.evaluate(() => [...(CSS.highlights.get('find-match') || [])].map((x) => x.toString()));
  expect(highlighted).toEqual(Array(4).fill('[Client name]'));
});

test('filling in replaces every place at once, as one undoable edit, and keeps the formatting', async ({ page }) => {
  await openInEditor(page, await buildDocx(QUOTE));
  await fillButton(page).click();
  const d = dialog(page);
  await d.getByLabel('Client name').fill('Acme Ltd');
  await d.getByLabel('Amount').fill('$1,250.00');
  await d.getByRole('button', { name: 'Fill in', exact: true }).click();
  await expect(page.locator('footer.statusbar').getByText('Filled in 2 placeholders in 4 places')).toBeVisible();
  // What is left: the date, and the locked one, which can't be filled in here.
  await expect(d.locator('label')).toHaveText(['date', 'Client name']);
  await expect(d.getByRole('status')).toHaveText('Filled in 2 placeholders in 4 places. 1 left.');
  await expect(d.getByLabel('Client name')).toBeDisabled();
  const text = page.locator('.docx-page');
  await expect(text).toContainText('Quote for Acme Ltd');
  await expect(text).toContainText('Dear Acme Ltd, thank you.');
  await expect(text).toContainText('Total: $1,250.00');
  await expect(text).toContainText('Signed for Acme Ltd');
  await expect(text).toContainText('Locked: [Client name]'); // read-only, left as it is
  await expect(text).toContainText('As agreed [1], tick [x] and [sic] per [Smith 2020].');
  await expect(fillButton(page)).toHaveAccessibleName('Fill in placeholders (1)');

  await d.getByRole('button', { name: 'Cancel' }).click();
  await expect(d).toHaveCount(0);
  const { xml } = await download(page);
  // The replacement takes the formatting of the placeholder's first character, as Replace does.
  expect(xml).toMatch(/<w:b\/><\/w:rPr><w:t[^>]*>Acme Ltd<\/w:t>/);
  expect(xml.match(/\[Client/g)).toHaveLength(1); // only the locked one
  expect(xml).toContain('{{date}}');

  await page.locator('.docx-page p').first().click();
  await page.keyboard.press('ControlOrMeta+z');
  await expect(text).toContainText('Quote for [Client name]');
  await expect(text).toContainText('Total: «Amount»');
});

test('placeholders filled in while tracking are tracked changes', async ({ page }) => {
  await openInEditor(page, await buildDocx(p(r('Dear [Client name],'))));
  await page.getByRole('button', { name: 'Track changes' }).click();
  const name = page.getByRole('dialog', { name: 'Your name on tracked changes' });
  await name.getByRole('textbox').fill('Rita');
  await name.getByRole('button', { name: 'Start tracking' }).click();
  await expect(page.getByRole('button', { name: 'Track changes' })).toHaveAttribute('aria-pressed', 'true');
  await fillButton(page).click();
  await dialog(page).getByLabel('Client name').fill('Acme Ltd');
  await dialog(page).getByRole('button', { name: 'Fill in', exact: true }).click();
  await expect(dialog(page)).toHaveCount(0); // nothing left to fill in
  const { xml } = await download(page);
  expect(xml).toMatch(/<w:del [^>]*>.*\[Client name\].*<\/w:del>/s);
  expect(xml).toMatch(/<w:ins [^>]*>.*Acme Ltd.*<\/w:ins>/s);
});

test('a document without placeholders says what they look like', async ({ page }) => {
  await openInEditor(page, await buildDocx(p(r('Nothing to fill in here [1].'))));
  await expect(fillButton(page)).toHaveAccessibleName('Fill in placeholders');
  await fillButton(page).click();
  await expect(dialog(page)).toContainText('No placeholders found. Placeholders are text in brackets, such as [Client name], {{date}}, «Amount» or <<Amount>>.');
  await page.keyboard.press('Escape');
  await expect(dialog(page)).toHaveCount(0);
});

// ------------------------------------------------------------------ new document from a template

test('a template is made into a new .docx, and the template itself is unchanged', async ({ page }) => {
  await openInEditor(page, await asTemplate(QUOTE), 'Quote.dotx');
  const bar = page.getByRole('region', { name: 'Template' });
  await expect(bar).toContainText('Word template. Downloading saves the template itself.');
  await expect(page.locator('.app-header .ext')).toHaveText('.dotx');
  await bar.getByRole('button', { name: 'Start a new document' }).click();
  await expect(page.locator('.app-header .ext')).toHaveText('.docx');
  await expect(bar).toContainText('New document from this template. It downloads as a .docx; the template (Quote.dotx) stays as it is');

  // The bar offers the placeholders too.
  await bar.getByRole('button', { name: 'Fill in 3 placeholders…' }).click();
  await dialog(page).getByLabel('Client name').fill('Acme Ltd');
  await dialog(page).getByRole('button', { name: 'Fill in', exact: true }).click();
  await dialog(page).getByRole('button', { name: 'Cancel' }).click();

  const { name, zip } = await downloaded(page);
  expect(name).toBe('Quote.docx');
  const ct = await zip.file('[Content_Types].xml').async('string');
  expect(ct).toContain(MAIN);
  expect(ct).not.toContain(DOTX);
  expect(await zip.file('word/document.xml').async('string')).toContain('Acme Ltd');
});

test('a new document made from an unedited template still downloads as a .docx', async ({ page }) => {
  await openInEditor(page, await asTemplate(p(r('Letterhead'))), 'Letter.dotx');
  await page.getByRole('region', { name: 'Template' }).getByRole('button', { name: 'Start a new document' }).click();
  const { name, zip } = await downloaded(page);
  expect(name).toBe('Letter.docx');
  expect(await zip.file('[Content_Types].xml').async('string')).toContain(MAIN);
});

test('going back to the template saves the template, and a .dotm makes a .docm with its macros', async ({ page }) => {
  await openInEditor(page, await asTemplate(p(r('Body')), DOTM), 'Form.dotm');
  const bar = page.getByRole('region', { name: 'Template' });
  await bar.getByRole('button', { name: 'Start a new document' }).click();
  await expect(page.locator('.app-header .ext')).toHaveText('.docm');
  await bar.getByRole('button', { name: 'Edit the template instead' }).click();
  await expect(page.locator('.app-header .ext')).toHaveText('.dotm');
  let { name, zip } = await downloaded(page);
  expect(name).toBe('Form.dotm');
  expect(await zip.file('[Content_Types].xml').async('string')).toContain(DOTM);

  await bar.getByRole('button', { name: 'Start a new document' }).click();
  ({ name, zip } = await downloaded(page));
  expect(name).toBe('Form.docm');
  expect(await zip.file('[Content_Types].xml').async('string')).toContain(DOCM);
  expect(Buffer.from(await zip.file('word/vbaProject.bin').async('uint8array')).toString()).toBe('VBA project bytes');
});

test('a document that is not a template has no template bar', async ({ page }) => {
  await openInEditor(page, await buildDocx(p(r('Body'))));
  await expect(page.getByRole('region', { name: 'Template' })).toHaveCount(0);
});

// ------------------------------------------------------------------ popups fit in the window

for (const width of [1280, 1440, 1600]) {
  test(`toolbar popups stay inside a ${width}px window`, async ({ page }) => {
    await page.setViewportSize({ width, height: 800 });
    await openInEditor(page, await buildDocx(p(r('Text [Client name]'))));
    await page.locator('.docx-page p').first().click();
    for (const [button, popup] of [
      ['Insert table', '.table-picker'],
      ['Fill in placeholders', '.fill-in'],
      ['Insert a check box or date', '[role=menu]'],
      ['More ways to save', '[role=menu]'],
    ]) {
      await page.getByRole('button', { name: button }).first().click();
      const box = await page.locator(popup).first().boundingBox();
      expect(box.x, `${button} popup left edge`).toBeGreaterThanOrEqual(0);
      expect(box.x + box.width, `${button} popup right edge`).toBeLessThanOrEqual(width);
      await page.keyboard.press('Escape');
    }
  });
}
