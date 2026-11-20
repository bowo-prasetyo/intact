// What a long manuscript needs, as Word has it: indents in Modify Style (a first-line indent for
// Normal, so every paragraph follows), Show markup on the Review tab (all, simple or no markup:
// only how the page shows tracked changes) and Previous and Next comment.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, W_NS } from './helpers.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const status = (page) => page.locator('.statusbar [role=status]');
const para = (text) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const p = (page, text) => page.locator('.docx-page > p', { hasText: text });
const styleXml = (styles, id) => styles.match(new RegExp(`<w:style [^>]*w:styleId="${id}".*?</w:style>`, 's'))[0];

async function caretAt(page, text, offset = 0) {
  await page.evaluate(([t, o]) => {
    const root = document.querySelector('.docx-page');
    const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    for (let n = w.nextNode(); n; n = w.nextNode()) {
      if (!n.data.includes(t)) continue;
      const range = document.createRange();
      range.setStart(n, n.data.indexOf(t) + o);
      root.focus();
      getSelection().removeAllRanges();
      getSelection().addRange(range);
      return;
    }
    throw new Error(`no text "${t}"`);
  }, [text, offset]);
}

async function modifyStyleOf(page, text) {
  await p(page, text).click({ button: 'right' });
  await page.getByRole('menu', { name: 'Context menu' }).getByRole('menuitem', { name: 'Modify style…' }).click();
  return page.getByRole('dialog', { name: /Modify style/ });
}

// ------------------------------------------------------------------ indents in Modify Style

test('Modify Style gives Normal a first-line indent: every paragraph follows, and Ctrl+Z puts it back', async ({ page }) => {
  await openInEditor(page, await buildDocx(para('It was a quiet night.') + para('She walked along the river.')));
  const d = await modifyStyleOf(page, 'It was a quiet night.');
  await expect(d.getByLabel('Style special indent', { exact: true })).toHaveValue('none');
  await expect(d.getByLabel('Style special indent by')).toBeDisabled();
  await d.getByLabel('Style special indent', { exact: true }).selectOption('firstLine');
  await d.getByLabel('Style special indent by').fill('1.27'); // A4: centimetres
  await d.getByRole('button', { name: 'OK' }).click();
  for (const text of ['It was a quiet night.', 'She walked along the river.']) {
    await expect.poll(() => p(page, text).evaluate((el) => parseFloat(getComputedStyle(el).textIndent))).toBeCloseTo(48, 0); // 36 pt
  }
  await expect(status(page)).toContainText('changed: every paragraph in it follows');

  const { zip } = await download(page);
  const normal = styleXml(await zip.file('word/styles.xml').async('string'), 'Normal');
  expect(normal).toContain('<w:pPr><w:ind w:firstLine="720"/></w:pPr>');
  // The paragraphs themselves aren't changed: it is the style.
  expect(await zip.file('word/document.xml').async('string')).toContain('<w:p><w:r><w:t xml:space="preserve">It was a quiet night.</w:t></w:r></w:p>');

  await page.locator('.docx-page').focus();
  await page.keyboard.press('Control+z');
  await expect.poll(() => p(page, 'It was a quiet night.').evaluate((el) => parseFloat(getComputedStyle(el).textIndent))).toBe(0);
});

test('Modify Style shows a style’s indents as Word’s dialog does, in either unit, and makes a hanging indent', async ({ page }) => {
  const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${W_NS}>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Bibliography"><w:name w:val="Bibliography"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="1440" w:right="360" w:firstLine="720" w:firstLineChars="100"/></w:pPr></w:style>
</w:styles>`;
  const body = '<w:p><w:pPr><w:pStyle w:val="Bibliography"/></w:pPr><w:r><w:t>Doe, J. A Book. 2020.</w:t></w:r></w:p>' + para('Body text.');
  await openInEditor(page, await buildDocx(body, { parts: { 'word/styles.xml': styles } }));
  const d = await modifyStyleOf(page, 'Doe, J.');
  await expect(d.getByLabel('Style left indent')).toHaveValue('2.54');
  await expect(d.getByLabel('Style right indent')).toHaveValue('0.64');
  await expect(d.getByLabel('Style special indent', { exact: true })).toHaveValue('firstLine');
  await expect(d.getByLabel('Style special indent by')).toHaveValue('1.27');
  await d.getByLabel('Style units').selectOption('in');
  await expect(d.getByLabel('Style left indent')).toHaveValue('1');
  await expect(d.getByLabel('Style special indent by')).toHaveValue('0.5');
  // Hanging, by half an inch: Left stays where the first line starts, as in Word's dialog.
  await d.getByLabel('Style special indent', { exact: true }).selectOption('hanging');
  await d.getByRole('button', { name: 'OK' }).click();
  const entry = p(page, 'Doe, J.');
  await expect.poll(() => entry.evaluate((el) => parseFloat(getComputedStyle(el).textIndent))).toBeCloseTo(-48, 0);
  await expect.poll(() => entry.evaluate((el) => parseFloat(getComputedStyle(el).paddingLeft))).toBeCloseTo(144, 0); // 1.5 in
  await expect.poll(() => p(page, 'Body text.').evaluate((el) => parseFloat(getComputedStyle(el).textIndent))).toBe(0); // Normal isn't changed

  const { zip } = await download(page);
  const written = await zip.file('word/styles.xml').async('string');
  // The first line's indent (and its character units) give way to the hanging one; Right, left alone, stays.
  expect(styleXml(written, 'Bibliography')).toContain('<w:ind w:left="2160" w:right="360" w:hanging="720"/>');
  expect(styleXml(written, 'Normal')).not.toContain('w:ind');
  await page.evaluate(() => localStorage.removeItem('intact.unit'));
});

// ------------------------------------------------------------------ Show markup

const ins = (author, text, id) => `<w:ins w:id="${id}" w:author="${author}" w:date="2026-01-0${id}T00:00:00Z"><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:ins>`;
const del = (author, text, id) => `<w:del w:id="${id}" w:author="${author}" w:date="2026-01-0${id}T00:00:00Z"><w:r><w:delText xml:space="preserve">${text}</w:delText></w:r></w:del>`;
const EDITED = `<w:p><w:r><w:t xml:space="preserve">She walked </w:t></w:r>${ins('Editor', 'slowly ', 1)}<w:r><w:t xml:space="preserve">along the </w:t></w:r>${del('Editor', 'dark ', 2)}<w:r><w:t>river.</w:t></w:r></w:p>`
  + para('The lamps flickered.');

test('No markup shows the text as it will be once the changes are accepted; the file keeps them', async ({ page }) => {
  await openInEditor(page, await buildDocx(EDITED));
  const view = page.getByRole('combobox', { name: 'Show markup' });
  await expect(view).toHaveValue('all');
  const inserted = page.locator('.docx-page .inserted', { hasText: 'slowly' });
  const deleted = page.locator('.docx-page .deleted', { hasText: 'dark' });
  await expect(deleted).toBeVisible();
  await view.selectOption('none');
  await expect(status(page)).toContainText('No markup: the text as it will be once the changes are accepted – 2 tracked changes kept in the file');
  await expect(deleted).toBeHidden();
  await expect(inserted).toBeVisible();
  const color = await p(page, 'She walked').evaluate((el) => getComputedStyle(el).color);
  await expect(inserted).toHaveCSS('color', color);
  await expect(inserted).toHaveCSS('text-decoration-line', 'none');
  expect(await p(page, 'She walked').innerText()).toBe('She walked slowly along the river.');

  // Only the view: the changes are all in the file, as they were.
  const { xml } = await download(page);
  expect(xml).toContain('<w:ins w:id="1" w:author="Editor"');
  expect(xml).toContain('<w:delText xml:space="preserve">dark </w:delText>');

  // Going to a change shows all markup again, so it can be seen.
  await page.getByRole('group', { name: 'Tracked changes' }).getByRole('button', { name: 'Next change' }).click();
  await expect(view).toHaveValue('all');
  await expect(deleted).toBeVisible();
});

test('Simple markup draws a red line beside each changed paragraph, and your own tracked edits follow the view', async ({ page }) => {
  await openInEditor(page, await buildDocx(EDITED));
  const view = page.getByRole('combobox', { name: 'Show markup' });
  await view.selectOption('simple');
  await expect(p(page, 'She walked')).toHaveCSS('box-shadow', /rgb\(179, 38, 30\)/);
  await expect(p(page, 'The lamps')).toHaveCSS('box-shadow', 'none');

  // Tracked typing and deleting in simple markup: recorded, shown as the text will be.
  await page.getByRole('button', { name: 'Track changes' }).click();
  const named = page.getByRole('dialog', { name: 'Your name on tracked changes' });
  await expect(named.getByRole('textbox')).toHaveAttribute('autocomplete', 'off');
  await named.getByRole('textbox').fill('Rita');
  await named.getByRole('button', { name: 'Start tracking' }).click();
  await expect(page.getByRole('button', { name: 'Track changes' })).toHaveAttribute('aria-pressed', 'true');
  await caretAt(page, 'flickered');
  await page.keyboard.type('gently ');
  const own = page.locator('.docx-page .tc-ins', { hasText: 'gently' });
  await expect(own).toHaveCount(1);
  const color = await p(page, 'The lamps').evaluate((el) => getComputedStyle(el).color);
  await expect(own).toHaveCSS('color', color);
  await expect(p(page, 'The lamps')).toHaveCSS('box-shadow', /rgb\(179, 38, 30\)/);
  await caretAt(page, 'lamps', 5);
  for (let i = 0; i < 6; i++) await page.keyboard.press('Backspace'); // "lamps" and its space before
  await expect(page.locator('.docx-page .tc-del', { hasText: 'lamps' })).toBeHidden();

  await view.selectOption('all');
  await expect(page.locator('.docx-page .tc-del', { hasText: 'lamps' })).toBeVisible();
  await expect(own).toHaveCSS('color', 'rgb(111, 45, 189)');
  const { xml } = await download(page);
  expect(xml).toMatch(/<w:ins [^>]*><w:r><w:t xml:space="preserve">gently <\/w:t><\/w:r><\/w:ins>/);
  expect(xml).toMatch(/<w:del [^>]*>.*lamps.*<\/w:del>/s);
});

test('Original shows the text before the changes, read-only; All markup makes it editable again', async ({ page }) => {
  await openInEditor(page, await buildDocx(EDITED));
  const view = page.getByRole('combobox', { name: 'Show markup' });
  await view.selectOption('original');
  await expect(status(page)).toContainText('Original: the text before the changes, read-only – 2 tracked changes kept in the file');
  const inserted = page.locator('.docx-page .inserted', { hasText: 'slowly' });
  const deleted = page.locator('.docx-page .deleted', { hasText: 'dark' });
  await expect(inserted).toBeHidden();
  await expect(deleted).toBeVisible();
  await expect(deleted).toHaveCSS('text-decoration-line', 'none');
  expect(await p(page, 'She walked').innerText()).toBe('She walked along the dark river.');

  // Typing changes nothing, and says why.
  await p(page, 'The lamps').click();
  await page.keyboard.type('xyz');
  await page.keyboard.press('Enter');
  await expect(status(page)).toContainText('Showing the original, before the tracked changes: choose All markup (Review) to edit');
  expect(await p(page, 'The lamps').innerText()).toBe('The lamps flickered.');
  await expect(page.locator('.docx-page > p')).toHaveCount(2);

  await view.selectOption('all');
  await caretAt(page, 'flickered', 0);
  await page.keyboard.type('gently ');
  await expect(p(page, 'The lamps')).toHaveText('The lamps gently flickered.');
});

// ------------------------------------------------------------------ Previous and Next comment

const comment = (id, text) => `<w:comment w:id="${id}" w:author="Editor" w:date="2026-01-01T00:00:00Z"><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:comment>`;
const commented = (id, text) => `<w:commentRangeStart w:id="${id}"/><w:r><w:t xml:space="preserve">${text}</w:t></w:r><w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r>`;
const COMMENTED = `<w:p>${commented(1, 'The night')}<w:r><w:t xml:space="preserve"> was quiet and </w:t></w:r>${commented(2, 'she walked')}<w:r><w:t xml:space="preserve"> on.</w:t></w:r></w:p>`
  + para('Nothing here.') + `<w:p><w:r><w:t xml:space="preserve">Later, </w:t></w:r>${commented(3, 'the storm')}<w:r><w:t xml:space="preserve"> came.</w:t></w:r></w:p>`;
const COMMENTS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments ${W_NS}>${comment(1, 'Too quiet?')}${comment(2, 'Who is she?')}${comment(3, 'Foreshadow this.')}</w:comments>`;

test('Previous and Next comment go through the comments in the text, select what they are on and show their thread', async ({ page }) => {
  await page.setViewportSize({ width: 1280, height: 800 });
  await openInEditor(page, await buildDocx(COMMENTED, { parts: { 'word/comments.xml': COMMENTS }, rels: `<Relationship Id="rIdC" Type="${REL}/comments" Target="comments.xml"/>` }));
  const nav = page.getByRole('group', { name: 'Go to a comment' });
  const selected = () => page.evaluate(() => getSelection().toString());
  const active = page.locator('.comments .thread.active');

  // From the caret: the next one after it, in the text.
  await caretAt(page, 'Nothing here.', 3);
  await nav.getByRole('button', { name: 'Next comment' }).click();
  await expect(status(page)).toContainText('Comment 3 of 3, by Editor');
  expect(await selected()).toBe('the storm');
  await expect(active).toContainText('Foreshadow this.');

  // From the comment shown: round to the first, then back.
  await nav.getByRole('button', { name: 'Next comment' }).click();
  await expect(status(page)).toContainText('Comment 1 of 3');
  expect(await selected()).toBe('The night');
  await expect(active).toContainText('Too quiet?');
  await nav.getByRole('button', { name: 'Previous comment' }).click();
  await expect(status(page)).toContainText('Comment 3 of 3');
  await nav.getByRole('button', { name: 'Previous comment' }).click();
  await expect(status(page)).toContainText('Comment 2 of 3');
  expect(await selected()).toBe('she walked');
  await expect(active).toContainText('Who is she?');
});

test('No markup hides the comments’ highlights, as Word does; the panel still lists them', async ({ page, browserName }) => {
  test.skip(browserName !== 'chromium', 'a highlight’s computed style (getComputedStyle with ::highlight) is read in Chromium');
  await page.setViewportSize({ width: 1280, height: 800 });
  await openInEditor(page, await buildDocx(COMMENTED, { parts: { 'word/comments.xml': COMMENTS }, rels: `<Relationship Id="rIdC" Type="${REL}/comments" Target="comments.xml"/>` }));
  const highlight = () => page.evaluate(() => {
    const el = document.querySelector('.docx-page > p');
    return getComputedStyle(el, '::highlight(comment)').backgroundColor;
  });
  expect(await highlight()).toBe('rgb(251, 227, 220)');
  await page.getByRole('combobox', { name: 'Show markup' }).selectOption('none');
  expect(await highlight()).toBe('rgba(0, 0, 0, 0)');
  await expect(page.locator('.comments .thread')).toHaveCount(3);
});
