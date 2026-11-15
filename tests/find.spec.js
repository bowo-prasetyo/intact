// Find and replace: Ctrl+F / Ctrl+H, matches across formatting but not across paragraphs, Word's
// rule that a replacement takes the formatting of the match's first character, and read-only
// content that is never changed.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const DOC = `<w:p><w:r><w:t xml:space="preserve">Agreement between ACME and </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>John</w:t></w:r><w:r><w:t xml:space="preserve"> Smith.</w:t></w:r></w:p>
<w:p><w:r><w:t>John</w:t></w:r></w:p>
<w:p><w:r><w:t>Smith signs for john smith ltd; Johnson too.</w:t></w:r></w:p>
<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Signed: John Smith</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
<w:sdt><w:sdtPr><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Locked: John Smith</w:t></w:r></w:p></w:sdtContent></w:sdt>
<w:p><w:r><w:t>Line one John</w:t><w:br/><w:t>Smith line two</w:t></w:r></w:p>`;

const findBox = (page) => page.getByRole('textbox', { name: 'Find' });
const status = (page) => page.locator('.findbar .status');
const highlighted = (page) => page.evaluate(() => [...(CSS.highlights.get('find-match') || [])].map((r) => r.toString()));
const currentMatch = (page) => page.evaluate(() => [...(CSS.highlights.get('find-current') || [])].map((r) => r.toString()).join());

async function open(page) {
  await openInEditor(page, await buildDocx(DOC));
  await page.locator('.docx-page p').first().click();
  await page.keyboard.press('ControlOrMeta+Home');
}

test('Ctrl+F finds text across formatting, but not across paragraphs or line breaks', async ({ page }) => {
  await open(page);
  await page.keyboard.press('ControlOrMeta+f');
  await expect(findBox(page)).toBeFocused();
  await page.keyboard.type('john smith');
  // Paragraph 1 (bold "John" + " Smith"), "john smith ltd", the table cell; the locked field is
  // counted separately; "John" / "Smith" split by a paragraph or a line break are not matches.
  await expect(status(page)).toHaveText('1 of 3 · 1 in read-only content');
  expect(await highlighted(page)).toEqual(['John Smith', 'john smith', 'John Smith']);
  expect(await currentMatch(page)).toBe('John Smith');

  await page.keyboard.press('Enter');
  await expect(status(page)).toHaveText('2 of 3 · 1 in read-only content');
  expect(await currentMatch(page)).toBe('john smith');
  await page.keyboard.press('Enter');
  await page.keyboard.press('Enter');
  await expect(status(page)).toHaveText('1 of 3 · 1 in read-only content'); // wraps around
  await page.keyboard.press('Shift+Enter');
  await expect(status(page)).toHaveText('3 of 3 · 1 in read-only content');
});

test('Match case and Whole words narrow the search', async ({ page }) => {
  await open(page);
  await page.keyboard.press('ControlOrMeta+f');
  await page.keyboard.type('John');
  await expect(status(page)).toHaveText('1 of 6 · 1 in read-only content'); // includes "john" and "Johnson"
  await page.getByLabel('Match case').check();
  await expect(status(page)).toHaveText('1 of 5 · 1 in read-only content');
  await page.getByLabel('Whole words').check();
  await expect(status(page)).toHaveText('1 of 4 · 1 in read-only content');
});

test('Replace takes the formatting of the first matched character and moves on', async ({ page }) => {
  await open(page);
  await page.keyboard.press('ControlOrMeta+h');
  await findBox(page).fill('John Smith');
  await page.getByLabel('Match case').check();
  await page.getByRole('textbox', { name: 'Replace with' }).fill('Jane Doe');
  await page.getByRole('button', { name: 'Replace', exact: true }).click();
  await expect(page.locator('.docx-page p').first()).toHaveText('Agreement between ACME and Jane Doe.');
  await expect(status(page)).toHaveText('1 of 1 · 1 in read-only content'); // now on the table cell

  const { xml } = await download(page);
  expect(xml).toContain('<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Jane Doe</w:t></w:r><w:r><w:t xml:space="preserve">.</w:t></w:r>');
  expect(xml).toContain('Signed: John Smith');
});

test('Replace all is one undo step and never touches read-only content', async ({ page }) => {
  await open(page);
  await page.keyboard.press('ControlOrMeta+h');
  await findBox(page).fill('john smith');
  await page.getByRole('textbox', { name: 'Replace with' }).fill('Jane Doe');
  await page.getByRole('button', { name: 'Replace all' }).click();
  await expect(status(page)).toHaveText('Replaced 3 matches');
  await expect(page.locator('td')).toHaveText('Signed: Jane Doe');
  await expect(page.locator('.sdt-locked')).toHaveText('Locked: John Smith');

  let { xml } = await download(page);
  expect(xml).not.toMatch(/>(Signed: |Smith signs for )?john smith/i);
  expect(xml).toContain('Locked: John Smith');

  await page.getByTitle('Undo (Ctrl+Z)').click();
  await expect(page.locator('td')).toHaveText('Signed: John Smith');
  await expect(page.locator('.docx-page p').first()).toHaveText('Agreement between ACME and John Smith.');
  ({ xml } = await download(page));
  expect(xml).not.toContain('Jane Doe');
});

test('selected text pre-fills Find; Esc closes it and leaves the match selected', async ({ page }) => {
  await open(page);
  await page.locator('.docx-page p', { hasText: 'Smith signs' }).click();
  await page.keyboard.press('Home');
  await page.keyboard.press('Shift+End');
  await page.keyboard.press('ControlOrMeta+f');
  await expect(findBox(page)).toHaveValue('Smith signs for john smith ltd; Johnson too.');
  await findBox(page).fill('Johnson');
  await page.keyboard.press('Escape');
  await expect(page.locator('.findbar')).toHaveCount(0);
  expect(await page.evaluate(() => getSelection().toString())).toBe('Johnson');
  expect(await page.evaluate(() => CSS.highlights.get('find-match'))).toBeUndefined();
  await page.keyboard.type('Jones');
  await expect(page.locator('.docx-page p', { hasText: 'Smith signs' })).toHaveText('Smith signs for john smith ltd; Jones too.');
});

test('matches follow the document as it is edited', async ({ page }) => {
  await open(page);
  await page.keyboard.press('ControlOrMeta+f');
  await page.keyboard.type('Doe');
  await expect(status(page)).toHaveText('No results');
  await page.locator('.docx-page p', { hasText: /^John$/ }).click();
  await page.keyboard.press('End');
  await page.keyboard.type(' Doe');
  await expect(status(page)).toHaveText('1 of 1');
});
