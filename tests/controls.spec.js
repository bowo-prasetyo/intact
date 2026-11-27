// Content controls, the fill-in fields of templates and forms: text, checkbox, drop-down and date
// fields can be filled in; locked and data-bound ones stay read-only and are kept exactly.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, FIELDS } from './helpers.js';

const FORM = `<w:p><w:r><w:t xml:space="preserve">Client: </w:t></w:r>${FIELDS}</w:p>
<w:sdt><w:sdtPr><w:alias w:val="Terms"/><w:id w:val="201"/><w:showingPlcHdr/></w:sdtPr><w:sdtContent><w:p><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>Enter the terms here.</w:t></w:r></w:p></w:sdtContent></w:sdt>
<w:sdt><w:sdtPr><w:alias w:val="Boilerplate"/><w:lock w:val="sdtContentLocked"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Locked boilerplate</w:t></w:r></w:p></w:sdtContent></w:sdt>
<w:p><w:r><w:t>Last</w:t></w:r></w:p>`;

const control = (page, label) => page.locator(`.docx-page .sdt[title*="${label}"]`);
/** The saved <w:sdt> whose properties contain `marker`. */
const savedControl = (xml, marker) => xml.match(/<w:sdt>.*?<\/w:sdt>/gs).find((s) => s.includes(marker));

test('typing into a placeholder replaces it and the field keeps its properties', async ({ page }) => {
  await openInEditor(page, await buildDocx(FORM));
  const name = control(page, 'Client name');
  await expect(name).toHaveClass(/showing-placeholder/);
  await name.locator('.sdt-content').click();
  await page.keyboard.type('Jane Doe');
  await expect(name).toHaveText('Jane Doe');
  await expect(name).not.toHaveClass(/showing-placeholder/);

  const { xml } = await download(page);
  const sdt = savedControl(xml, 'w:val="client"');
  expect(sdt).toContain('<w:alias w:val="Client name"/>');
  expect(sdt).toContain('<w:id w:val="101"/>');
  expect(sdt).not.toContain('showingPlcHdr');
  expect(sdt).not.toContain('PlaceholderText');
  // New text takes the field's own formatting (bold here), as in Word.
  expect(sdt).toMatch(/<w:sdtContent><w:r><w:rPr><w:b\/><\/w:rPr><w:t xml:space="preserve">Jane Doe<\/w:t><\/w:r><\/w:sdtContent>/);
});

test('a placeholder that was cleared is written back so Word shows it again', async ({ page }) => {
  await openInEditor(page, await buildDocx(FORM));
  await control(page, 'Client name').locator('.sdt-content').click();
  await page.keyboard.press('Delete');
  const { xml } = await download(page);
  const sdt = savedControl(xml, 'w:val="client"');
  expect(sdt).toContain('<w:showingPlcHdr/>');
  expect(sdt).toContain('Click or tap here to enter text.');
});

test('a paragraph field holds paragraphs; a field inside a line does not', async ({ page }) => {
  await openInEditor(page, await buildDocx(FORM));
  await control(page, 'Terms').locator('.sdt-content p').click();
  await page.keyboard.type('First term');
  await page.keyboard.press('Enter');
  await page.keyboard.type('Second term');

  await control(page, 'Client name').locator('.sdt-content').click();
  await page.keyboard.type('Jane');
  await page.keyboard.press('Enter');
  await page.keyboard.type(' Doe');
  await expect(control(page, 'Client name')).toHaveText('Jane Doe');

  const { xml } = await download(page);
  const terms = savedControl(xml, 'w:val="Terms"');
  expect(terms).toMatch(/<w:sdtContent><w:p><w:r><w:t xml:space="preserve">First term<\/w:t><\/w:r><\/w:p><w:p><w:r><w:t xml:space="preserve">Second term<\/w:t><\/w:r><\/w:p><\/w:sdtContent>/);
  expect(savedControl(xml, 'w:val="client"')).not.toContain('<w:br/>');
});

test('plain text fields refuse character formatting', async ({ page }) => {
  await openInEditor(page, await buildDocx(FORM));
  await control(page, 'Client name').locator('.sdt-content').click();
  await page.keyboard.type('Jane');
  await page.keyboard.press('Shift+Home');
  await page.getByTitle('Bold (Ctrl+B)').click();
  await page.keyboard.press('ControlOrMeta+i');
  await expect(page.locator('.statusbar')).toContainText('Plain text fields can’t be formatted');
  await expect(control(page, 'Client name').locator('b, strong, i, em')).toHaveCount(0);
});

test('clicking a check box ticks it, and undo clears it again', async ({ page }) => {
  await openInEditor(page, await buildDocx(FORM));
  const box = page.locator('.docx-page .sdt-checkbox');
  await box.click();
  await expect(box).toHaveText('☒');
  let { xml } = await download(page);
  let sdt = savedControl(xml, 'w14:checkbox');
  expect(sdt).toContain('<w14:checked w14:val="1"/>');
  expect(sdt).toContain('>☒</w:t>');

  await page.getByTitle('Undo (Ctrl+Z)').click();
  await expect(box).toHaveText('☐');
  ({ xml } = await download(page));
  expect(savedControl(xml, 'w14:checkbox')).toContain('<w14:checked w14:val="0"/>');
});

test('a drop-down list offers its items and saves the choice', async ({ page }) => {
  await openInEditor(page, await buildDocx(FORM));
  await control(page, 'Plan').click();
  await expect(page.getByRole('listbox').getByRole('option')).toHaveText(['Basic plan', 'Premium plan']);
  await page.getByRole('option', { name: 'Premium plan' }).click();
  await expect(control(page, 'Plan')).toHaveText('Premium plan');
  await expect(page.getByRole('listbox').getByRole('option')).toHaveCount(0);

  const { xml } = await download(page);
  const sdt = savedControl(xml, 'w:val="Plan"');
  expect(sdt).toContain('w:lastValue="premium"');
  expect(sdt).not.toContain('showingPlcHdr');
  expect(sdt).toContain('<w:t xml:space="preserve">Premium plan</w:t>');
  expect(sdt).not.toContain('PlaceholderText');
});

test('a date field is filled in from a date picker, in the field\'s own format', async ({ page }) => {
  await openInEditor(page, await buildDocx(FORM));
  await control(page, 'Start date').click();
  await page.getByLabel('Date', { exact: true }).fill('2026-03-05');
  await expect(control(page, 'Start date')).toHaveText('5 March 2026');
  const { xml } = await download(page);
  const sdt = savedControl(xml, 'w:val="Start date"');
  expect(sdt).toContain('w:fullDate="2026-03-05T00:00:00Z"');
  expect(sdt).toContain('>5 March 2026</w:t>');
});

test('a combo box takes typed text or one of its items', async ({ page }) => {
  const COMBO = '<w:p><w:sdt><w:sdtPr><w:alias w:val="Country"/><w:comboBox><w:listItem w:displayText="Indonesia" w:value="ID"/>'
    + '<w:listItem w:displayText="Netherlands" w:value="NL"/></w:comboBox></w:sdtPr><w:sdtContent><w:r><w:t>Spain</w:t></w:r></w:sdtContent></w:sdt></w:p>';
  await openInEditor(page, await buildDocx(COMBO));
  const combo = control(page, 'Country');
  await combo.locator('.sdt-content').click();
  await page.keyboard.press('End');
  await page.keyboard.type(' (mainland)');
  await expect(combo).toHaveText('Spain (mainland)');
  await expect(page.getByRole('listbox')).toHaveCount(0); // typing doesn't open the list

  const box = await combo.boundingBox();
  await combo.click({ position: { x: box.width - 3, y: box.height / 2 } }); // the ▾ after the text
  await page.getByRole('listbox').getByRole('option', { name: 'Netherlands' }).click();
  await expect(combo).toHaveText('Netherlands');
  const { xml } = await download(page);
  expect(savedControl(xml, 'w:val="Country"')).toContain('<w:comboBox w:lastValue="NL">');
});

test('locked and data-bound fields are read-only and say why', async ({ page }) => {
  await openInEditor(page, await buildDocx(FORM));
  const locked = control(page, 'Boilerplate');
  await expect(locked).toHaveAttribute('title', /locked in Word/);
  await expect(locked).toHaveAttribute('contenteditable', 'false');
  await expect(locked.locator('[contenteditable="true"]')).toHaveCount(0);
  await expect(control(page, 'Company')).toHaveAttribute('title', /linked to document data/);
  await expect(control(page, 'Company').locator('.sdt-content')).toHaveCount(0);
});

// ------------------------------------------------------------------ keyboard

/** Puts the caret just before (or after) the field whose title contains `label`. */
async function caretBy(page, label, after = false) {
  await page.evaluate(({ label, after }) => {
    const el = [...document.querySelectorAll('.docx-page .sdt')].find((e) => e.title.includes(label));
    document.querySelector('.docx-page').focus();
    const range = document.createRange();
    if (after) range.setStartAfter(el);
    else range.setStartBefore(el);
    getSelection().removeAllRanges();
    getSelection().addRange(range);
  }, { label, after });
}

const near = (page) => page.evaluate(() => [...(CSS.highlights.get('field-near') || [])].map((r) => r.toString()).join());

test('Alt+Down ticks a check box the caret is next to', async ({ page }) => {
  await openInEditor(page, await buildDocx(FORM));
  await caretBy(page, 'Check box');
  await expect.poll(() => near(page)).toBe('☐'); // the field the keys will act on is marked
  await page.keyboard.press('Alt+ArrowDown');
  await expect(page.locator('.sdt-checkbox')).toHaveText('☒');
});

test('Space ticks a selected check box', async ({ page }) => {
  await openInEditor(page, await buildDocx(FORM));
  await caretBy(page, 'Check box');
  await page.keyboard.press('Shift+ArrowRight');
  await page.keyboard.press(' ');
  await expect(page.locator('.sdt-checkbox')).toHaveText('☒');
});

test('a drop-down list is chosen from with the keyboard, and typing carries on after it', async ({ page }) => {
  await openInEditor(page, await buildDocx(FORM));
  await caretBy(page, 'Plan');
  await page.keyboard.press('Alt+ArrowDown');
  const options = page.getByRole('listbox').getByRole('option');
  await expect(options.first()).toBeFocused();
  await page.keyboard.press('ArrowDown');
  await expect(options.nth(1)).toBeFocused();
  await page.keyboard.press('Enter');
  await expect(control(page, 'Plan')).toHaveText('Premium plan');
  await expect(page.getByRole('listbox')).toHaveCount(0);
  await page.keyboard.type('!');
  expect(await control(page, 'Plan').evaluate((el) => el.nextSibling?.textContent?.startsWith('!'))).toBe(true);
});

test('a typed date is taken on Enter, not halfway through the year', async ({ page, browserName }) => {
  // Playwright's WebKit for Linux takes no typing in a date input at all (its value stays empty);
  // Safari on a Mac does, as Chromium and Firefox do. Picking a date from the picker is tested
  // above, in WebKit too.
  test.skip(browserName === 'webkit', 'the date input of WebKit for Linux takes no typed digits');
  await openInEditor(page, await buildDocx(FORM));
  await caretBy(page, 'Start date');
  await page.keyboard.press('Alt+ArrowDown');
  const input = page.getByLabel('Date', { exact: true });
  await expect(input).toBeFocused();
  await page.keyboard.type('03052');
  await expect(input).toBeVisible(); // "0002-03-05" is a complete date, but not what is meant
  await page.keyboard.type('026');
  await page.keyboard.press('Enter');
  await expect(control(page, 'Start date')).toHaveText('5 March 2026');
});

// ------------------------------------------------------------------ fields around table cells and rows

const TABLE_FIELDS = `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>
<w:tr><w:tc><w:p><w:r><w:t>Fee</w:t></w:r></w:p></w:tc><w:sdt><w:sdtPr><w:alias w:val="Fee amount"/><w:showingPlcHdr/><w:text/></w:sdtPr><w:sdtContent><w:tc><w:p><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>Amount</w:t></w:r></w:p></w:tc></w:sdtContent></w:sdt></w:tr>
<w:sdt><w:sdtPr><w:alias w:val="Signatures"/><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent><w:tr><w:tc><w:p><w:r><w:t>Signed</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Locked row</w:t></w:r></w:p></w:tc></w:tr></w:sdtContent></w:sdt>
<w:tr><w:tc><w:p><w:r><w:t>Last</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Row</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;

test('fields around table cells and rows are shown; cell fields are filled in like other fields', async ({ page }) => {
  await openInEditor(page, await buildDocx(TABLE_FIELDS));
  await expect(page.locator('td')).toHaveText(['Fee', 'Amount', 'Signed', 'Locked row', 'Last', 'Row']);
  const fee = page.locator('td').nth(1);
  await expect(fee).toHaveAttribute('title', /Fee amount/);
  await expect(page.locator('td', { hasText: 'Locked row' })).toHaveAttribute('contenteditable', 'false');

  await fee.click();
  await page.keyboard.type('EUR 900');
  await expect(fee).toHaveText('EUR 900');
  await page.locator('td', { hasText: 'Last' }).click();
  await page.keyboard.press('End');
  await page.keyboard.type(' one');

  const { xml } = await download(page);
  const feeField = savedControl(xml, 'Fee amount');
  expect(feeField).not.toContain('showingPlcHdr');
  expect(feeField).toMatch(/<w:sdtContent><w:tc><w:p><w:r><w:t xml:space="preserve">EUR 900<\/w:t><\/w:r><\/w:p><\/w:tc><\/w:sdtContent>/);
  expect(savedControl(xml, 'Signatures')).toContain('<w:t>Locked row</w:t>');
  expect(xml).toContain('>Last one</w:t>'); // the rows after the wrapped row still line up
});
