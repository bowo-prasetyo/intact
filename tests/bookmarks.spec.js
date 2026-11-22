// Bookmarks (Insert > Bookmark…), as Word's Bookmark dialog: name the selected text or the place at
// the caret (w:bookmarkStart / w:bookmarkEnd, never tracked), go to one, delete one; a name already
// used moves that bookmark. Cross-reference… then points to a bookmark (its text, or its paragraph's
// number), to a numbered item (its number, REF \r, or its text) and to a numbered heading's number.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const P = (t) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
const STEP = (t) => `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
const DOC = P('Safety first.') + STEP('Unplug the unit.') + STEP('Open the cover.') + STEP('Replace the fuse.') + P('See also the warranty terms.');

async function insertMenu(page, item) {
  await page.getByRole('button', { name: /Insert a check box or date/ }).click();
  await page.getByRole('menuitem', { name: item }).click();
}

async function select(page, text, words) {
  await page.evaluate(([t, w]) => {
    const p = [...document.querySelectorAll('.docx-page p')].find((e) => e.textContent.includes(t));
    const node = [...p.childNodes].map((n) => (n.nodeType === 3 ? n : n.firstChild)).find((n) => n?.nodeType === 3 && n.data.includes(w));
    const at = node.data.indexOf(w);
    document.querySelector('.docx-page').focus();
    getSelection().setBaseAndExtent(node, at, node, at + w.length);
  }, [text, words]);
}

async function caretAtEnd(page, text) {
  await page.locator('.docx-page p', { hasText: text }).first().click();
  await page.keyboard.press('End');
}

test('a bookmark around the selected text, written as Word writes one, and refused names say why', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await select(page, 'See also', 'warranty terms');
  await insertMenu(page, /^Bookmark…/);
  const dialog = page.getByRole('dialog', { name: 'Bookmark' });
  await expect(dialog.getByText('Marks the selected text')).toBeVisible();
  await dialog.getByLabel('Bookmark name').fill('2nd');
  await dialog.getByRole('button', { name: 'Add' }).click();
  await expect(dialog.getByRole('alert')).toHaveText('A bookmark’s name starts with a letter');
  await dialog.getByLabel('Bookmark name').fill('Warranty terms');
  await dialog.getByRole('button', { name: 'Add' }).click();
  await expect(dialog.getByRole('alert')).toHaveText(/no spaces/);
  await dialog.getByLabel('Bookmark name').fill('Warranty');
  await dialog.getByRole('button', { name: 'Add' }).click();
  await expect(dialog).toBeHidden();

  const { xml } = await download(page);
  expect(xml).toMatch(/<w:t xml:space="preserve">See also the <\/w:t><\/w:r><w:bookmarkStart w:id="\d+" w:name="Warranty"\/><w:r><w:t xml:space="preserve">warranty terms<\/w:t><\/w:r><w:bookmarkEnd w:id="\d+"\/>/);

  // Listed again, gone to, and deleted.
  await caretAtEnd(page, 'Safety first.');
  await insertMenu(page, /^Bookmark…/);
  await dialog.getByRole('option', { name: /Warranty/ }).click();
  await dialog.getByRole('button', { name: 'Go to' }).click();
  expect(await page.evaluate(() => getSelection().toString())).toBe('warranty terms');
  await insertMenu(page, /^Bookmark…/);
  await dialog.getByRole('option', { name: /Warranty/ }).click();
  await dialog.getByRole('button', { name: 'Delete' }).click();
  expect((await download(page)).xml).not.toContain('Warranty');
});

test('a name already used moves the bookmark, as in Word', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await select(page, 'Safety', 'Safety');
  await insertMenu(page, /^Bookmark…/);
  const dialog = page.getByRole('dialog', { name: 'Bookmark' });
  await dialog.getByLabel('Bookmark name').fill('Start');
  await dialog.getByRole('button', { name: 'Add' }).click();
  await select(page, 'See also', 'warranty');
  await insertMenu(page, /^Bookmark…/);
  await dialog.getByLabel('Bookmark name').fill('start');
  await dialog.getByRole('button', { name: 'Move here' }).click();
  const { xml } = await download(page);
  expect(xml.match(/<w:bookmarkStart /g)).toHaveLength(1);
  expect(xml).toMatch(/w:name="start"\/><w:r><w:t xml:space="preserve">warranty<\/w:t>/);
});

test('cross-references to a bookmark, a numbered item\'s number and its text', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await select(page, 'See also', 'warranty terms');
  await insertMenu(page, /^Bookmark…/);
  const bm = page.getByRole('dialog', { name: 'Bookmark' });
  await bm.getByLabel('Bookmark name').fill('Warranty');
  await bm.getByRole('button', { name: 'Add' }).click();

  const xref = page.getByRole('dialog', { name: 'Cross-reference' });
  await caretAtEnd(page, 'Safety first.');
  await page.keyboard.type(' Do step ');
  await insertMenu(page, /^Cross-reference…/);
  await xref.getByLabel('Reference type').selectOption('Numbered item');
  await expect(xref.getByLabel('Insert reference to')).toHaveValue('number');
  await xref.getByRole('option', { name: '3. Replace the fuse.' }).click();
  await xref.getByRole('button', { name: 'Insert' }).click();
  await page.keyboard.type(' (');
  await insertMenu(page, /^Cross-reference…/);
  await xref.getByLabel('Reference type').selectOption('Numbered item');
  await xref.getByLabel('Insert reference to').selectOption('text');
  await xref.getByRole('option', { name: '3. Replace the fuse.' }).click();
  await xref.getByRole('button', { name: 'Insert' }).click();
  await page.keyboard.type(') and read the ');
  await insertMenu(page, /^Cross-reference…/);
  await xref.getByLabel('Reference type').selectOption('Bookmark');
  await xref.getByRole('option', { name: /^Warranty: warranty terms/ }).click();
  await xref.getByRole('button', { name: 'Insert' }).click();
  await expect(page.locator('.docx-page p').first()).toHaveText('Safety first. Do step 3 (Replace the fuse.) and read the warranty terms');

  const { xml } = await download(page);
  const fuse = /<w:bookmarkStart w:id="\d+" w:name="(_Ref\d+)"\/><w:r><w:t xml:space="preserve">Replace the fuse\./.exec(xml)?.[1];
  expect(fuse).toBeTruthy();
  expect(xml).toContain(`<w:instrText xml:space="preserve"> REF ${fuse} \\r \\h </w:instrText>`);
  expect(xml).toMatch(/ REF _Ref\d+ \\h <\/w:instrText>/);
  expect(xml).toContain('<w:instrText xml:space="preserve"> REF Warranty \\h </w:instrText>');

  // F9 keeps the number in step with the list: a step added before it renumbers the reference.
  await caretAtEnd(page, 'Unplug the unit.');
  await page.keyboard.press('Enter');
  await page.keyboard.type('Wait a minute.');
  await page.keyboard.press('F9');
  await expect(page.locator('.docx-page p').first()).toHaveText('Safety first. Do step 4 (Replace the fuse.) and read the warranty terms');
});

test('a numbered heading\'s number, once headings are numbered', async ({ page }) => {
  const H = (n, t) => `<w:p><w:pPr><w:pStyle w:val="Heading${n}"/></w:pPr><w:r><w:t>${t}</w:t></w:r></w:p>`;
  await openInEditor(page, await buildDocx(H(1, 'Overview') + H(1, 'Maintenance') + P('As described in section ')));
  await caretAtEnd(page, 'As described');
  await page.getByRole('button', { name: 'Numbering library' }).click();
  await page.getByRole('menu', { name: 'Numbering library' }).getByRole('menuitem', { name: /^1 Heading 1/ }).click();
  await caretAtEnd(page, 'As described');
  await insertMenu(page, /^Cross-reference…/);
  const xref = page.getByRole('dialog', { name: 'Cross-reference' });
  await xref.getByLabel('Reference type').selectOption('Heading');
  await xref.getByLabel('Insert reference to').selectOption('number');
  await xref.getByRole('option', { name: '2 Maintenance' }).click();
  await xref.getByRole('button', { name: 'Insert' }).click();
  await expect(page.locator('.docx-page p', { hasText: 'As described' })).toHaveText('As described in section 2');
  expect((await download(page)).xml).toMatch(/ REF _Ref\d+ \\r \\h <\/w:instrText>/);
});
