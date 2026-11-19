// Pasting a table from Excel, Google Sheets, Word or a web page gives a real table, as Word pastes
// one: its rows and columns, merged cells, relative column widths and each cell's basic formatting.
// Where a table can't go (a header, a note, a field, another table's cell) its rows are pasted as
// lines with tabs, and the status bar says so.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const DOC = '<w:p><w:r><w:t xml:space="preserve">Start  end</w:t></w:r></w:p><w:p><w:r><w:t>Other</w:t></w:r></w:p>';

/** Pastes `html` (and `text`) with the caret after `where` in the paragraph holding it (or in `selector`). */
async function pasteAt(page, { html = '', text = '' }, where = 'Start ', selector = '.docx-page p') {
  await page.evaluate(({ html, text, where, selector }) => {
    document.querySelector('.docx-page').focus();
    const p = [...document.querySelectorAll(selector)].find((e) => e.textContent.includes(where.trim()));
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    let node = walker.nextNode();
    while (node && !node.data.includes(where)) node = walker.nextNode();
    const range = document.createRange();
    range.setStart(node, node.data.indexOf(where) + where.length);
    getSelection().removeAllRanges();
    getSelection().addRange(range);
    const dt = new DataTransfer();
    if (html) dt.setData('text/html', html);
    dt.setData('text/plain', text);
    // Firefox ignores clipboardData given to a scripted ClipboardEvent's constructor.
    const event = new ClipboardEvent('paste', { bubbles: true, cancelable: true });
    Object.defineProperty(event, 'clipboardData', { value: dt });
    p.dispatchEvent(event);
  }, { html, text, where, selector });
}

async function startTracking(page, name = 'Rita') {
  await page.getByRole('button', { name: 'Track changes' }).click();
  const dialog = page.getByRole('dialog', { name: 'Your name on tracked changes' });
  await dialog.getByRole('textbox').fill(name);
  await dialog.getByRole('button', { name: 'Start tracking' }).click();
  await expect(page.getByRole('button', { name: 'Track changes' })).toHaveAttribute('aria-pressed', 'true');
}

const cellTexts = (table) => table.locator('tr').evaluateAll((rows) => rows.map((r) => [...r.cells].filter((c) => c.style.display !== 'none').map((c) => c.textContent)));

// What Excel puts on the clipboard: a class per cell format, declared in a <style>.
const EXCEL = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
<head><meta name=ProgId content=Excel.Sheet><style>
<!--table {mso-displayed-decimal-separator:"\\.";}
.xl65 {mso-style-parent:style0; font-weight:700;}
.xl66 {mso-style-parent:style0; mso-number-format:"\\#\\,\\#\\#0";}
-->
</style></head><body link="#0563C1">
<table border=0 cellpadding=0 cellspacing=0 width=256 style='border-collapse:collapse;width:192pt'>
<!--StartFragment-->
 <col width=128 style='width:96pt'>
 <col width=64 span=2 style='width:48pt'>
 <tr height=20 style='height:15.0pt'>
  <td height=20 class=xl65 width=128 style='height:15.0pt;width:96pt'>Item</td>
  <td class=xl65 width=64 style='width:48pt'>Qty</td>
  <td class=xl65 width=64 style='width:48pt'>Price</td>
 </tr>
 <tr height=20 style='height:15.0pt'>
  <td height=20 style='height:15.0pt'>Paper A4</td>
  <td align=right>10</td>
  <td class=xl66 align=right>45,000</td>
 </tr>
 <tr height=20 style='height:15.0pt'>
  <td height=20 style='height:15.0pt'>Toner</td>
  <td align=right>2</td>
  <td class=xl66 align=right>&nbsp;</td>
 </tr>
<!--EndFragment-->
</table></body></html>`;

test('a range copied from Excel is pasted as a table, with its columns’ widths and its bold heading row', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await pasteAt(page, { html: EXCEL, text: 'Item\tQty\tPrice\nPaper A4\t10\t45,000\nToner\t2\t\n' });
  const table = page.locator('.docx-page table');
  await expect(table).toHaveCount(1);
  expect(await cellTexts(table)).toEqual([['Item', 'Qty', 'Price'], ['Paper A4', '10', '45,000'], ['Toner', '2', '']]);
  await expect(table.locator('tr').first().locator('strong')).toHaveText(['Item', 'Qty', 'Price']);
  await expect(page.locator('.statusbar [role=status]')).toContainText('Table pasted: 3 × 3');
  // The paragraph was split at the caret: the table goes between its halves.
  await expect(page.locator('.docx-page > p')).toHaveText(['Start ', ' end', 'Other']);

  const { xml } = await download(page);
  const tbl = xml.match(/<w:tbl>.*<\/w:tbl>/s)[0];
  expect(tbl).toContain('<w:tblStyle w:val="TableGrid"/>');
  const grid = [...tbl.matchAll(/<w:gridCol w:w="(\d+)"\/>/g)].map((m) => Number(m[1]));
  expect(grid).toHaveLength(3);
  expect(grid[0] / grid[1]).toBeCloseTo(2, 1); // 96pt : 48pt : 48pt
  expect(grid[1]).toBe(grid[2]);
  expect(tbl).toContain('<w:b/></w:rPr><w:t xml:space="preserve">Item</w:t>');
  expect(tbl).toContain('<w:t xml:space="preserve">45,000</w:t>');
  expect(tbl.match(/<w:tr>/g)).toHaveLength(3);
  // The table sits between the two halves of the paragraph it was pasted in.
  expect(xml.indexOf('>Start <')).toBeLessThan(xml.indexOf('<w:tbl>'));
  expect(xml.indexOf('</w:tbl>')).toBeLessThan(xml.indexOf('> end<'));
});

test('a range from Google Sheets keeps its bold cells and merged cells', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await pasteAt(page, {
    html: `<meta charset="utf-8"><google-sheets-html-origin><style type="text/css"><!--td {border: 1px solid #cccccc;}--></style>
<table xmlns="http://www.w3.org/1999/xhtml" cellspacing="0" cellpadding="0" dir="ltr" border="1" style="table-layout:fixed;font-size:10pt;font-family:Arial;width:0px;border-collapse:collapse;border:none">
<colgroup><col width="100"/><col width="100"/><col width="100"/></colgroup><tbody>
<tr style="height:21px;"><td style="overflow:hidden;padding:2px 3px 2px 3px;vertical-align:bottom;font-weight:bold;" rowspan="1" colspan="3">Quarterly report</td></tr>
<tr style="height:21px;"><td rowspan="2" style="vertical-align:middle;">North</td><td>Q1</td><td>120</td></tr>
<tr style="height:21px;"><td>Q2</td><td style="font-style:italic;">135</td></tr>
</tbody></table>`,
  });
  const table = page.locator('.docx-page table');
  await expect(table).toHaveCount(1);
  await expect(table.locator('td').first()).toHaveAttribute('colspan', '3');
  await expect(table.locator('strong')).toHaveText('Quarterly report');
  await expect(table.locator('em')).toHaveText('135');

  const { xml } = await download(page);
  const tbl = xml.match(/<w:tbl>.*<\/w:tbl>/s)[0];
  expect(tbl).toContain('<w:gridSpan w:val="3"/>');
  expect(tbl).toContain('<w:vMerge w:val="restart"/>');
  expect(tbl).toMatch(/<w:vMerge\/>/);
  const rows = tbl.match(/<w:tr>.*?<\/w:tr>/gs);
  expect(rows).toHaveLength(3);
  expect(rows[0].match(/<w:tc>/g)).toHaveLength(1);
  expect(rows[1]).toContain('North');
  expect(rows[2]).not.toContain('North');
});

test('text, a table and text from Word paste as a paragraph, a table and a paragraph', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await pasteAt(page, {
    html: `<html xmlns:w="urn:schemas-microsoft-com:office:word"><body><!--StartFragment-->
<p class=MsoNormal>Prices below.</p>
<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse;border:none'>
 <tr><td width=301 valign=top style='width:225.4pt'><p class=MsoNormal><b>Name</b></p></td>
  <td width=301 valign=top style='width:225.4pt'><p class=MsoNormal>Role</p></td></tr>
 <tr><td width=301 valign=top><p class=MsoNormal>Ana</p><p class=MsoNormal>Budi</p></td>
  <td width=301 valign=top><p class=MsoNormal>Secretary</p></td></tr>
</table>
<p class=MsoNormal>Thanks.</p><!--EndFragment--></body></html>`,
  }, 'Other', '.docx-page p');
  const table = page.locator('.docx-page table');
  await expect(table).toHaveCount(1);
  expect(await cellTexts(table)).toEqual([['Name', 'Role'], ['AnaBudi', 'Secretary']]);
  await expect(table.locator('tr').nth(1).locator('td').first().locator('p')).toHaveText(['Ana', 'Budi']);
  await expect(page.locator('.docx-page > p')).toHaveText(['Start  end', 'OtherPrices below.', 'Thanks.']);
});

test('a table pasted while tracking is a tracked insertion, rows and text; Undo takes it away', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await startTracking(page);
  await pasteAt(page, { html: '<table><tr><td>a</td><td>b</td></tr><tr><td>c</td><td>d</td></tr></table>' }, 'Other');
  const table = page.locator('.docx-page table');
  await expect(table.locator('tr[data-rev-row="ins"]')).toHaveCount(2);
  await expect(table.locator('[data-tc="ins"]')).toHaveText(['a', 'b', 'c', 'd']);
  const { xml } = await download(page);
  const tbl = xml.match(/<w:tbl>.*<\/w:tbl>/s)[0];
  expect(tbl.match(/<w:trPr><w:ins /g)).toHaveLength(2);
  expect(tbl.match(/<w:ins /g).length).toBeGreaterThanOrEqual(6);

  await page.locator('.docx-page').focus();
  await page.keyboard.press('Control+z');
  await expect(page.locator('.docx-page table')).toHaveCount(0);
});

test('where a table can’t go, its rows are pasted as lines with tabs, and the status bar says why', async ({ page }) => {
  await openInEditor(page, await buildDocx(`<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>${DOC}`));
  await pasteAt(page, { html: '<table><tr><td>a</td><td>b</td></tr><tr><td>c</td><td>d</td></tr></table>' }, 'Cell', '.docx-page td p');
  await expect(page.locator('.docx-page table')).toHaveCount(1);
  await expect(page.locator('.docx-page td p')).toHaveText(['Cella\tb', 'c\td']);
  await expect(page.locator('.statusbar [role=status]')).toContainText('pasted as text');
});

test('one cell copied from a spreadsheet pastes as its text, in the line', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await pasteAt(page, { html: '<table><tr><td style="font-weight:bold">42</td></tr></table>', text: '42\n' });
  await expect(page.locator('.docx-page table')).toHaveCount(0);
  await expect(page.locator('.docx-page p').first()).toHaveText('Start 42 end');
  await expect(page.locator('.docx-page p').first().locator('strong')).toHaveText('42');
});
