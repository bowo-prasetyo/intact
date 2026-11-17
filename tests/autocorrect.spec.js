// Word's AutoCorrect for symbols as you type: (c) ©, (r) ®, (tm) ™, (e) €, --> →, <-- ←, ==> ⇒,
// <== ⇐ and <=> ⇔. Ctrl+Z straight after gives back what was typed; a list item's letter "(c)" is
// left as it is; the AutoFormat switch turns it off; tracked while tracking.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const DOC = '<w:p><w:r><w:t>Start</w:t></w:r></w:p>';
const para = (page) => page.locator('.docx-page > p').first();

async function typeAtEnd(page, text) {
  await para(page).click();
  await page.keyboard.press('End');
  await page.keyboard.type(text);
}

test('symbols and arrows are replaced as the last character is typed', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await typeAtEnd(page, ' (c) 2026 Acme(TM), (r) and (e)5; go --> back <-- if ==> then <== iff <=> done');
  await expect(para(page)).toHaveText('Start © 2026 Acme™, ® and €5; go → back ← if ⇒ then ⇐ iff ⇔ done');
  const { xml } = await download(page);
  expect(xml).toContain('Start © 2026 Acme™, ® and €5; go → back ← if ⇒ then ⇐ iff ⇔ done');
});

test('Ctrl+Z straight after gives back what was typed', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await typeAtEnd(page, ' (c)');
  await expect(para(page)).toHaveText('Start ©');
  await page.keyboard.press('ControlOrMeta+z');
  await expect(para(page)).toHaveText('Start (c)');
  await page.keyboard.type(' kept');
  await expect(para(page)).toHaveText('Start (c) kept');
});

test('a list item\'s letter (c) stays as it is', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await typeAtEnd(page, ': (a) one, (b) two and (c) three; (d) and (e) too');
  await expect(para(page)).toHaveText('Start: (a) one, (b) two and (c) three; (d) and (e) too');
});

test('with AutoFormat off, nothing is replaced', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await page.getByRole('button', { name: /^AutoFormat: on$/ }).click();
  await typeAtEnd(page, ' (c) -->');
  await expect(para(page)).toHaveText('Start (c) -->');
});

test('while tracking, the symbol is a tracked insertion of its own and nothing is left deleted', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await page.getByRole('button', { name: 'Track changes' }).click();
  const dialog = page.getByRole('dialog', { name: 'Your name on tracked changes' });
  await dialog.getByRole('textbox').fill('Rita');
  await dialog.getByRole('button', { name: 'Start tracking' }).click();
  await typeAtEnd(page, ' (c) Acme');
  await expect(para(page)).toHaveText('Start © Acme');
  const { xml } = await download(page);
  expect(xml).toMatch(/<w:ins w:id="\d+" w:author="Rita" w:date="[^"]+"><w:r><w:t xml:space="preserve"> © Acme<\/w:t><\/w:r><\/w:ins>/);
  expect(xml).not.toContain('<w:del ');
});
