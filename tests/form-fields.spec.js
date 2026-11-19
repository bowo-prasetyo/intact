// Legacy form fields (Word's Developer → Legacy Tools: FORMTEXT, FORMCHECKBOX, FORMDROPDOWN), as
// school, medical and club forms still use them: check boxes and lists are shown and can be filled
// in, text fields typed in, also while Word allows only filling in forms; saving changes their
// w:ffData and nothing else.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, canonicalize, W_NS } from './helpers.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const status = (page) => page.locator('.statusbar [role=status]');

const begin = (ff) => `<w:r><w:fldChar w:fldCharType="begin"><w:ffData>${ff}</w:ffData></w:fldChar></w:r>`;
const instr = (code) => `<w:r><w:instrText xml:space="preserve"> ${code} </w:instrText></w:r>`;
const SEP = '<w:r><w:fldChar w:fldCharType="separate"/></w:r>';
const END = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
const BLANK = ' '.repeat(5);

const TEXT_FIELD = begin('<w:name w:val="Child"/><w:enabled/><w:calcOnExit w:val="0"/><w:textInput/>')
  + instr('FORMTEXT') + SEP + `<w:r><w:t>${BLANK}</w:t></w:r>` + END;
const checkBox = (name, extra = '') => begin(`<w:name w:val="${name}"/><w:enabled/><w:calcOnExit w:val="0"/><w:checkBox><w:sizeAuto/><w:default w:val="0"/>${extra}</w:checkBox>`)
  + instr('FORMCHECKBOX') + END;
const DROPDOWN = begin('<w:name w:val="Size"/><w:enabled/><w:calcOnExit w:val="0"/><w:ddList><w:listEntry w:val="Small"/><w:listEntry w:val="Medium"/><w:listEntry w:val="Large"/></w:ddList>')
  + instr('FORMDROPDOWN') + END;

const SLIP = `<w:p><w:r><w:t>Field trip permission slip</w:t></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">Child’s name: </w:t></w:r>${TEXT_FIELD}</w:p>
<w:p>${checkBox('Permission')}<w:r><w:t xml:space="preserve"> I give permission for my child to attend.</w:t></w:r></w:p>
<w:p>${checkBox('Photos', '<w:checked/>')}<w:r><w:t xml:space="preserve"> Photos may be taken.</w:t></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">T-shirt size: </w:t></w:r>${DROPDOWN}</w:p>`;

function protectedForForms(body) {
  return buildDocx(body, {
    parts: { 'word/settings.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ${W_NS}><w:documentProtection w:edit="forms" w:enforcement="1"/></w:settings>` },
    rels: `<Relationship Id="rIdS" Type="${REL}/settings" Target="settings.xml"/>`,
  });
}

/** The <w:ffData> of each form field in saved document.xml, in order. */
const ffData = (xml) => [...xml.matchAll(/<w:ffData>([\s\S]*?)<\/w:ffData>/g)].map((m) => m[1]);

test('check boxes and lists are shown as Word shows them, and saving unchanged keeps them', async ({ page }) => {
  const docx = await buildDocx(SLIP);
  await openInEditor(page, docx);
  const boxes = page.locator('.docx-page .ff-check');
  await expect(boxes).toHaveCount(2);
  await expect(boxes.nth(0)).toHaveAttribute('aria-checked', 'false');
  await expect(boxes.nth(1)).toHaveAttribute('aria-checked', 'true');
  await expect(boxes.nth(0)).toBeVisible();
  const glyph = (i) => boxes.nth(i).evaluate((el) => getComputedStyle(el, '::before').content);
  expect(await glyph(0)).toBe('"☐"');
  expect(await glyph(1)).toBe('"☒"');
  await expect(boxes.nth(0)).toHaveAttribute('title', 'Check box (Permission) – click to tick or clear it');
  const list = page.locator('.docx-page .ff-drop');
  expect(await list.evaluate((el) => getComputedStyle(el, '::before').content)).toBe('"Small"');
  // Drawn, not typed: the paragraph's text has no box in it.
  await expect(page.locator('.docx-page p').nth(2)).toHaveText(' I give permission for my child to attend.');

  const { xml } = await download(page);
  const original = await (await import('./helpers.js')).documentXmlOf(docx);
  expect(await canonicalize(page, xml)).toEqual(await canonicalize(page, original));
});

test('a click ticks a check box and a list offers its entries; Undo puts them back', async ({ page }) => {
  await openInEditor(page, await buildDocx(SLIP));
  const boxes = page.locator('.docx-page .ff-check');
  await boxes.nth(0).click();
  await expect(boxes.nth(0)).toHaveAttribute('aria-checked', 'true');
  await expect(status(page)).toContainText('Permission ticked');
  await boxes.nth(1).click();
  await expect(boxes.nth(1)).toHaveAttribute('aria-checked', 'false');

  await page.locator('.docx-page .ff-drop').click();
  const chooser = page.getByRole('listbox', { name: 'Choose a value' });
  await expect(chooser.getByRole('option')).toHaveText(['Small', 'Medium', 'Large']);
  await chooser.getByRole('option', { name: 'Large' }).click();
  expect(await page.locator('.docx-page .ff-drop').evaluate((el) => getComputedStyle(el, '::before').content)).toBe('"Large"');

  const { xml } = await download(page);
  const [, permission, photos, size] = ffData(xml);
  expect(permission).toContain('<w:checkBox><w:sizeAuto/><w:default w:val="0"/><w:checked/></w:checkBox>');
  expect(photos).toContain('<w:checkBox><w:sizeAuto/><w:default w:val="0"/></w:checkBox>');
  expect(size).toContain('<w:ddList><w:result w:val="2"/><w:listEntry w:val="Small"/>');
  // Nothing but the field data changed in those paragraphs.
  expect(xml).toContain('<w:t xml:space="preserve"> I give permission for my child to attend.</w:t>');

  await page.keyboard.press('Escape');
  await page.locator('.docx-page').focus();
  await page.keyboard.press('Control+z');
  expect(await page.locator('.docx-page .ff-drop').evaluate((el) => getComputedStyle(el, '::before').content)).toBe('"Small"');
  await page.keyboard.press('Control+z');
  await page.keyboard.press('Control+z');
  await expect(boxes.nth(0)).toHaveAttribute('aria-checked', 'false');
  await expect(boxes.nth(1)).toHaveAttribute('aria-checked', 'true');
});

test('clicking a blank text form field selects it, so typing replaces the blank', async ({ page }) => {
  await openInEditor(page, await buildDocx(SLIP));
  await page.locator('.docx-page .ff-text').click();
  await page.keyboard.type('Sam Lee');
  await expect(page.locator('.docx-page .ff-text')).toHaveText('Sam Lee');
  const { xml } = await download(page);
  expect(xml).toMatch(/FORMTEXT <\/w:instrText><\/w:r><w:r><w:fldChar w:fldCharType="separate"\/><\/w:r><w:r><w:t[^>]*>Sam Lee<\/w:t><\/w:r><w:r><w:fldChar w:fldCharType="end"\/>/);
});

test('a form protected for filling in: its fields are filled in, Tab goes from field to field', async ({ page }) => {
  await openInEditor(page, await protectedForForms(SLIP));
  // The rest of the text stays locked.
  await page.locator('.docx-page p').first().click();
  await page.keyboard.type('x');
  await expect(page.locator('.docx-page p').first()).toHaveText('Field trip permission slip');

  // Tab from the start: the text field, selected.
  await page.locator('.docx-page .ff-text').click();
  await page.keyboard.type('Sam');
  await page.keyboard.press('Tab');
  await expect(page.locator('.docx-page .ff-check').nth(0)).toBeFocused();
  await page.keyboard.press('Space');
  await expect(page.locator('.docx-page .ff-check').nth(0)).toHaveAttribute('aria-checked', 'true');
  await page.keyboard.press('Tab');
  await page.keyboard.press('Tab');
  await expect(page.locator('.docx-page .ff-drop')).toBeFocused();
  await page.keyboard.press('Alt+ArrowDown');
  await page.getByRole('option', { name: 'Medium', exact: true }).click();
  await page.keyboard.press('Shift+Tab');
  await expect(page.locator('.docx-page .ff-check').nth(1)).toBeFocused();
  // Enter in a text field goes on, too, rather than splitting the paragraph.
  await page.locator('.docx-page .ff-text').click();
  await page.keyboard.press('End');
  await page.keyboard.press('Enter');
  await expect(page.locator('.docx-page .ff-check').nth(0)).toBeFocused();
  await expect(page.locator('.docx-page p')).toHaveCount(5);

  const { xml } = await download(page);
  const [, permission, , size] = ffData(xml);
  expect(xml).toMatch(/<w:fldChar w:fldCharType="separate"\/><\/w:r><w:r><w:t[^>]*>Sam<\/w:t><\/w:r><w:r><w:fldChar w:fldCharType="end"\/>/);
  expect(permission).toContain('<w:checked/>');
  expect(size).toContain('<w:result w:val="1"/>');
  expect(xml).toMatch(/<w:t[^>]*>Field trip permission slip<\/w:t>/);
});

test('a read-only document says so when its check box is clicked', async ({ page }) => {
  const docx = await buildDocx(SLIP, {
    parts: { 'word/settings.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ${W_NS}><w:documentProtection w:edit="readOnly" w:enforcement="1"/></w:settings>` },
    rels: `<Relationship Id="rIdS" Type="${REL}/settings" Target="settings.xml"/>`,
  });
  await openInEditor(page, docx);
  await page.locator('.docx-page .ff-check').first().click();
  await expect(page.locator('.docx-page .ff-check').first()).toHaveAttribute('aria-checked', 'false');
  await expect(status(page)).toContainText('Read-only');
});

test('a field turned off in Word is shown but not changed', async ({ page }) => {
  const off = begin('<w:name w:val="Office"/><w:enabled w:val="0"/><w:checkBox><w:sizeAuto/><w:default w:val="0"/></w:checkBox>') + instr('FORMCHECKBOX') + END;
  await openInEditor(page, await buildDocx(`<w:p>${off}<w:r><w:t xml:space="preserve"> Office use only</w:t></w:r></w:p>`));
  const box = page.locator('.docx-page .ff-check');
  await box.click();
  await expect(box).toHaveAttribute('aria-checked', 'false');
  await expect(status(page)).toContainText('turned off in Word');
});
